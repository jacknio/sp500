#!/usr/bin/env python3
"""Step 2: python build_dashboard.py — saved raw CSV -> analysis -> dashboard.html.

Run python fetch_data.py first, or use --raw PATH for an existing raw CSV.
Read in order: 1 READ -> 2 ANALYZE -> 3 DRAW -> 4 BUILD.
Uses only the standard library; no Bloomberg connection is opened here.
The HTML template is an input; dashboard.html and data/analysis/ are outputs.
"""
import argparse
import base64
import bisect
import calendar
import csv
import hashlib
import html
import json
import math
import re
import shutil
import tempfile
import statistics
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

# Share the configuration and CSV format. Importing fetch_data does NOT fetch:
# its CLI is guarded by if __name__ == "__main__" and blpapi is imported on fetch only.
from fetch_data import SERIES, RAW_COLUMNS, HISTORY_YEARS, years_before, csv_string

ROOT = Path(__file__).resolve().parent
MIN_OBSERVATIONS = {"Daily": 252, "Weekly": 52, "Monthly": 24, "Quarterly": 12}
BLUE, GRAY, LIGHT = "#244b73", "#718096", "#a8b2bf"

# ============================================================================
# 1. READ — the analysis only sees CSVs read back from disk, not API objects.
# ============================================================================
def read_raw(path):
    """Validate source rows before calculations; keep native dates and values.

    Legacy public CSVs can omit security/field/unit. A supplied unit must match
    the configured unit: changing a label must never silently rescale history.
    """
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        required = {"series_id", "observation_date", "value", "source", "retrieved_at"}
        if not required.issubset(reader.fieldnames or []):
            raise ValueError("CSV columns required: " + ", ".join(sorted(required)))
        if len(reader.fieldnames) != len(set(reader.fieldnames)):
            raise ValueError("Duplicate CSV column names")
        rows = []
        seen = set()
        for line_number, row in enumerate(reader, 2):
            if None in row or any(value is None for value in row.values()):
                raise ValueError("Malformed CSV row " + str(line_number) + ": wrong number of columns")
            row = {key: value.strip() for key, value in row.items()}
            key, day = row["series_id"], row["observation_date"]
            if key not in SERIES:
                raise ValueError("Unknown series_id; add its definition in fetch_data.py CONFIGURATION: " + key)
            if not row["source"] or not row["retrieved_at"]:
                raise ValueError("CSV source and retrieved_at must be populated: " + key + " " + day)
            if row.get("is_simulated", "").lower() in ("true", "1") or "simulat" in row["source"].lower():
                raise ValueError("Simulated observations are not accepted")
            try:
                parsed_day = date.fromisoformat(day)
            except ValueError as exc:
                raise ValueError("CSV observation_date must be YYYY-MM-DD: " + key + " " + day) from exc
            if day != parsed_day.isoformat():
                raise ValueError("CSV observation_date must be YYYY-MM-DD: " + key + " " + day)
            if (key, day) in seen:
                raise ValueError("Duplicate series/date: " + key + " " + day)
            seen.add((key, day))
            value = row["value"]
            try:
                row["value"] = None if value in ("", ".", "NA") else float(value)
            except ValueError as exc:
                raise ValueError("Invalid CSV value: " + key + " " + day + " = " + value) from exc
            if row["value"] is not None and not math.isfinite(row["value"]):
                raise ValueError("Nonfinite CSV value: " + key + " " + day)
            row.setdefault("security", "")
            row.setdefault("field", "")
            if row["source"].casefold() == "bloomberg":
                cfg = SERIES[key]
                for column, configured in (("security", cfg["ticker"]), ("field", cfg["field"])):
                    if not row[column] or " ".join(row[column].split()).casefold() != " ".join(configured.split()).casefold():
                        raise ValueError("Bloomberg CSV identity differs from configuration: " + key + " / " + column)
            row["unit"] = row.get("unit") or SERIES[key]["unit"]
            if row["unit"] != SERIES[key]["unit"]:
                raise ValueError("CSV unit differs from configured unit: " + key + " (" + row["unit"]
                                 + " versus " + SERIES[key]["unit"] + "). Use a separate ID or correct the mapping.")
            rows.append(row)
    # Mixing providers inside one history would create artificial jumps/ranges.
    for key in {r["series_id"] for r in rows}:
        identities = {(r["source"], r["security"], r["field"], r["unit"]) for r in rows if r["series_id"] == key}
        if len(identities) > 1:
            raise ValueError("Mixed sources/definitions in one series: " + key)
    return rows


def read_manifest(raw_path, rows):
    """Verify a saved snapshot without reinterpreting it under a new definition.

    Transform/name/group edits are analysis choices and remain allowed. Native
    frequency, units and Bloomberg identity changes require a new series ID.
    """
    manifest_path = raw_path.parent / "manifest.json"
    if not manifest_path.exists():
        return {}
    manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    if not isinstance(manifest, dict) or not isinstance(manifest.get("errors", {}), dict):
        raise ValueError("Invalid snapshot manifest: " + str(manifest_path))
    if manifest.get("sha256") and manifest["sha256"] != hashlib.sha256(raw_path.read_bytes()).hexdigest():
        raise ValueError("Raw CSV differs from its saved manifest hash. Restore the original snapshot or use a separate imported CSV folder.")
    if manifest.get("status") not in ("complete", "partial"):
        raise ValueError("Snapshot is not a completed download: " + str(manifest.get("status")))
    definitions = manifest.get("series", {})
    if not isinstance(definitions, dict):
        raise ValueError("Invalid series definitions in snapshot manifest")
    for key in {row["series_id"] for row in rows}:
        saved = definitions.get(key)
        if saved is None:
            continue
        if not isinstance(saved, dict):
            raise ValueError("Invalid snapshot definition: " + key)
        for field in ("ticker", "field", "unit", "frequency", "date_basis"):
            if field in saved and saved[field] != SERIES[key].get(field):
                raise ValueError("Snapshot definition differs from current configuration: " + key + " / " + field
                                 + ". Use a new series ID or restore the snapshot's mapping.")
    return manifest


# ============================================================================
# 2. ANALYZE — saved series, summaries and composite input standardization.
# The optional interactive composite combines these inputs in the HTML.
# ============================================================================
def quantile(values, probability):
    """Linear interpolation: sorted values, position = (n-1)*probability."""
    ordered = sorted(values)
    position = (len(ordered) - 1) * probability
    lower = int(position)
    upper = min(lower + 1, len(ordered) - 1)
    return ordered[lower] + (position - lower) * (ordered[upper] - ordered[lower])


def percentile(value, history):
    """Midrank empirical percentile; ties count as half, no distribution assumption."""
    return 100 * (sum(x < value for x in history) + 0.5 * sum(x == value for x in history)) / len(history)


def month_before(day, months):
    number = day.year * 12 + day.month - 1 - months
    return date(number // 12, number % 12 + 1, 1)


def period_before(day, months):
    """Same calendar day, preserving month-end series through short months."""
    target = month_before(day, months)
    last = calendar.monthrange(target.year, target.month)[1]
    is_month_end = day.day == calendar.monthrange(day.year, day.month)[1]
    return target.replace(day=last if is_month_end else min(day.day, last))


def transform(points, kind):
    if kind not in ("level", "avg_4w", "yoy", "annualized_3m", "annualized_6m", "avg_change_3m"):
        raise ValueError("Unsupported transform: " + str(kind))
    lookup = dict(points)
    output = []
    for day, value in points:
        result = None
        if value is not None:
            if kind == "level":
                result = value
            elif kind == "avg_4w":
                window = [lookup.get(day - timedelta(weeks=n)) for n in range(4)]
                if all(v is not None for v in window):
                    result = statistics.mean(window)
            else:
                if day.day != 1:
                    raise ValueError("Monthly transformations need reference-month dates (day 1), not release dates")
                months = {"yoy": 12, "annualized_3m": 3, "annualized_6m": 6, "avg_change_3m": 3}[kind]
                previous = lookup.get(month_before(day, months))
                # Compounded growth needs the exact endpoints, not intervening
                # observations. A missing middle month must not erase valid YoY.
                complete = kind != "avg_change_3m" or all(lookup.get(month_before(day, n)) is not None for n in range(months + 1))
                if previous is not None and complete:
                    if kind == "avg_change_3m":
                        result = (value - previous) / 3
                    elif value > 0 and previous > 0:
                        power = {"yoy": 1, "annualized_3m": 4, "annualized_6m": 2}[kind]
                        result = 100 * ((value / previous) ** power - 1)
        if result is not None and not math.isfinite(result):
            raise ValueError("Nonfinite transformed result: " + str(day) + " / " + kind)
        output.append((day, result))
    return output


def changes(points, frequency, horizon):
    """Daily: calendar weeks, last point <= target, at most 4 days earlier.
    Weekly/monthly/quarterly: exact prior week/calendar period. Missing bases stay missing.
    Returns date, difference, actual base date (for an auditable comparison).
    """
    valid = [(d, v) for d, v in points if v is not None]
    dates = [d for d, _ in valid]
    result = []
    for day, value in valid:
        if frequency in ("Monthly", "Quarterly"):
            target = period_before(day, horizon * (3 if frequency == "Quarterly" else 1))
        else:
            target = day - timedelta(weeks=horizon)
        index = bisect.bisect_right(dates, target) - 1
        tolerance = 4 if frequency == "Daily" else 0
        if index >= 0 and (target - dates[index]).days <= tolerance:
            result.append((day, value - valid[index][1], dates[index]))
        else:
            result.append((day, None, None))
    return result


def range_statistics(points, frequency):
    valid = [(d, v) for d, v in points if v is not None]
    if not valid:
        return None
    last_day, current = valid[-1]
    start = years_before(last_day, HISTORY_YEARS)
    history = [(d, v) for d, v in valid if start <= d < last_day]
    result = dict(date=str(last_day), value=current, window_start=str(start), count=len(history),
                  sample_start=str(history[0][0]) if history else None,
                  sample_end=str(history[-1][0]) if history else None,
                  sufficient=len(history) >= MIN_OBSERVATIONS[frequency])
    if result["sufficient"]:
        values = [v for _, v in history]
        result.update(min=min(values), p10=quantile(values, .10), median=quantile(values, .50),
                      p90=quantile(values, .90), max=max(values), percentile=percentile(current, values))
        sd = statistics.stdev(values)
        result["zscore"] = (current - statistics.mean(values)) / sd if sd else None
    return result


def number(value, signed=False):
    if value is None:
        return "—"
    return ("+" if signed and value > 0 else "") + format(value, ",.2f")


def analyze_series(key, config, points, source):
    frequency = config["frequency"]
    if frequency not in MIN_OBSERVATIONS:
        raise ValueError(key + ": unsupported frequency " + str(frequency))
    kind = config["transform"]
    if (kind == "avg_4w" and frequency != "Weekly") or (kind in ("yoy", "annualized_3m", "annualized_6m", "avg_change_3m") and frequency != "Monthly"):
        raise ValueError(key + ": transformation does not match the native frequency")
    display_points = transform(points, config["transform"])
    stats = range_statistics(display_points, config["frequency"])
    if stats is None:
        return None
    unit = "%" if config["transform"] in ("yoy", "annualized_3m", "annualized_6m") else config["unit"]
    if config["transform"] == "avg_change_3m":
        unit = config["unit"] + " / month"
    labels = {"level": "Reported level", "yoy": "Year-over-year change", "annualized_3m": "3M annualized growth",
              "annualized_6m": "6M annualized growth", "avg_change_3m": "3M average monthly change", "avg_4w": "4W average"}
    horizons = (1, 3, 12) if frequency == "Monthly" else (1, 2, 4) if frequency == "Quarterly" else (1, 4, 13)
    suffix = "M" if frequency == "Monthly" else "Q" if frequency == "Quarterly" else "W"
    rate_level = config["transform"] == "level" and unit in ("%", "Percent")
    factor = 100 if rate_level and key in ("BBG_2Y", "BBG_10Y", "DGS2", "DGS10", "DGS30", "DFII10", "T10YIE", "BAMLH0A0HYM2", "BAMLC0A0CM", "YIELD_CURVE", "IMPLIED_BE") else 1
    change_unit = "bp" if factor == 100 else "pp" if unit in ("%", "Percent") else unit
    comparisons = []
    for horizon in horizons:
        history = changes(display_points, config["frequency"], horizon)
        day, change, base = history[-1]
        comparisons.append(dict(label=str(horizon) + suffix,
                                value=change * factor if change is not None else None, base=str(base) if base else None))
    move_points = [(d, v * factor if v is not None else None) for d, v, _ in changes(display_points, config["frequency"], horizons[1])]
    move_stats = range_statistics(move_points, config["frequency"])
    # Do not use a stale previous change as the current change.
    if move_stats and move_stats["date"] != stats["date"]:
        move_stats = None
    text = config["name"] + " is " + number(stats["value"]) + " " + unit + " on " + stats["date"] + "."
    if stats["sufficient"]:
        position = "below P10" if stats["value"] < stats["p10"] else "above P90" if stats["value"] > stats["p90"] else "inside P10–P90"
        text += " Its level is " + position + " (" + number(stats["percentile"]) + "th percentile of available prior history)."
    middle = comparisons[1]
    if middle["value"] is not None:
        text += " The " + middle["label"] + " change is " + number(middle["value"], True) + " " + change_unit + " versus " + middle["base"] + "."
    if move_stats and move_stats["sufficient"]:
        text += " That change ranks at the " + number(move_stats["percentile"]) + "th percentile of earlier " + middle["label"] + " changes."
    return dict(id=key, name=config["name"], group=config["group"], unit=unit, frequency=config["frequency"],
                transformation=labels[config["transform"]], note=config["note"], source=source, points=display_points,
                stats=stats, comparisons=comparisons, change_unit=change_unit, move_points=move_points,
                move_stats=move_stats, summary=text)


def analyze(rows, asof):
    groups = defaultdict(list)
    sources = defaultdict(set)
    for row in rows:
        key, day = row["series_id"], date.fromisoformat(row["observation_date"])
        if day <= asof:
            groups[key].append((day, row["value"]))
            sources[key].add(row["source"] + (" · " + row["security"] + " / " + row["field"] if row["security"] else " · " + key))
    metrics = {}
    for key, cfg in SERIES.items():
        if key not in groups:
            continue
        metric = analyze_series(key, cfg, sorted(groups[key]), "; ".join(sorted(sources[key])))
        if metric:
            metric["raw_ids"] = [key]
            metrics[key] = metric
    emp, inf = "SURPRISE_EMPLOYMENT", "SURPRISE_INFLATION"
    if emp in metrics and inf in metrics:
        if (any(SERIES[key]["transform"] != "level" for key in (emp, inf))
                or metrics[emp]["unit"] != metrics[inf]["unit"]
                or metrics[emp]["frequency"] != metrics[inf]["frequency"]):
            raise ValueError("The default employment/inflation Average requires original levels with matching units and frequency. Use new IDs for alternative definitions.")
        frequency = metrics[emp]["frequency"]
        input_series = {key: dict(unit=metrics[key]["unit"], frequency=frequency,
                        transformation=metrics[key]["transformation"],
                        rows=[dict(date=str(d), value=v, original_value=v) for d, v in metrics[key]["points"]])
                        for key in (emp, inf)}
        aligned = calculate_composite(dict(weightMode="equal", scale="raw", alignment="auto", maxCarryDays=7,
                    components=[dict(id=key, weight=1, direction=1) for key in (emp, inf)]), input_series, str(asof))
        points = [(date.fromisoformat(p["date"]), p["value"]) for p in aligned["points"]]
        complete = [p for p in aligned["points"] if p["value"] is not None]
        e = {date.fromisoformat(p["date"]): p["parts"][0]["input"] for p in complete}
        i = {date.fromisoformat(p["date"]): p["parts"][1]["input"] for p in complete}
        days = sorted(e)
        cfg = dict(name="Employment + inflation Average", group="US surprises", unit="original index points",
                   frequency=frequency, transform="level", note="Average = (ECSULBUS + BCMPUSIF) / 2. Daily inputs use the latest prior observation, at most 7 calendar days old. Original scales are preserved; this is a custom calculation, not a published Fed index.")
        average = analyze_series("AVERAGE", cfg, points, "Calculated from " + metrics[emp]["source"] + " and " + metrics[inf]["source"])
        if average:
            average["raw_ids"] = [emp, inf]
            latest = days[-1]
            average["component_date"] = str(latest)
            average["components"] = {"employment": e[latest], "inflation": i[latest]}
            prior_dates = [d for d in days if years_before(latest, HISTORY_YEARS) <= d < latest]
            average["component_sd"] = None
            if len(prior_dates) >= MIN_OBSERVATIONS[frequency]:
                average["component_sd"] = {"employment": statistics.stdev(e[d] for d in prior_dates),
                                           "inflation": statistics.stdev(i[d] for d in prior_dates),
                                           "count": len(prior_dates)}
            if e[latest] * i[latest] < 0:
                average["summary"] += " Employment and inflation have opposite signs; the Average offsets their signals."
            base = average["comparisons"][1]["base"]
            average["contributions"] = None if not base else {"employment": (e[latest] - e[date.fromisoformat(base)]) / 2,
                                                              "inflation": (i[latest] - i[date.fromisoformat(base)]) / 2,
                                                              "base": base}
            metrics = {"AVERAGE": average, **metrics}
    return metrics

def prepare_composite_inputs(metrics, asof, raw_rows=()):
    """Prepare auditable inputs, not a selected/weighted composite.

    Each row contains the already-transformed chart value and its prior-only
    z-score at the NATIVE frequency. The current observation is excluded from
    mean/SD. Use up to ten preceding calendar years; retain warm-up history.
    Browser code chooses exact dates or completed calendar months afterwards.
    """
    columns = ["date", "value", "z", "history_count", "history_start",
               "history_end", "history_mean", "history_sd", "original_value"]
    originals = {(r["series_id"], r["observation_date"]): r["value"] for r in raw_rows}
    prepared, audit_rows = {}, []
    for key, metric in metrics.items():
        if key == "AVERAGE":  # Do not include the same two inputs twice by default.
            continue
        valid = [(d, v) for d, v in metric["points"] if v is not None and d <= asof]
        dates = [d for d, _ in valid]
        values = [v for _, v in valid]
        own_rows = []
        for index, (day, value) in enumerate(valid):
            first = bisect.bisect_left(dates, years_before(day, HISTORY_YEARS))
            history = values[first:index]  # Explicitly exclude today and the future.
            count = len(history)
            mean = sd = z = None
            if count >= MIN_OBSERVATIONS[metric["frequency"]]:
                mean = statistics.fmean(history)
                sd = math.sqrt(math.fsum((v - mean) ** 2 for v in history) / (count - 1))
                if sd > 0:
                    z = (value - mean) / sd
            row = [str(day), value, z, count, str(dates[first]) if count else None,
                   str(dates[index - 1]) if count else None, mean, sd,
                   originals.get((key, str(day)))]
            own_rows.append(row)
            audit_rows.append(dict(series_id=key, **dict(zip(columns, row))))
        prepared[key] = {name: metric[name] for name in
                         ("id", "name", "unit", "frequency", "transformation", "source")}
        prepared[key]["rows"] = own_rows
        prepared[key]["group"] = metric.get("group", "Other")
        # Derived measures have no single untransformed value. Their source
        # series are exported separately, with original dates and metadata.
        prepared[key]["raw_ids"] = metric.get("raw_ids", [key] if key in SERIES else [])
        prepared[key]["original_unit"] = SERIES[key]["unit"] if key in SERIES else ""
        cfg = SERIES.get(key, {})
        prepared[key].update(security=cfg.get("ticker", ""), field=cfg.get("field", ""),
                             date_basis=cfg.get("date_basis", "source observation date"))
    known = {key: {"name": cfg["name"]} for key, cfg in SERIES.items()}
    known.update({
        "YIELD_CURVE": {"name": "10Y minus 2Y Treasury yield"},
        "IMPLIED_BE": {"name": "10Y nominal minus real yield"},
        "LABOR_RATIO": {"name": "Job openings / unemployed persons"},
        "INCOME_GAP": {"name": "Real spending minus income momentum"},
    })
    return dict(asof=str(asof), columns=columns, series=prepared, known_series=known,
                history_years=HISTORY_YEARS, minimum_observations=MIN_OBSERVATIONS), audit_rows


def calculate_composite(recipe, inputs, cutoff):
    """Python counterpart of the HTML calculator, also used by Excel updates.

    inputs maps an ID to metadata plus rows of dictionaries (date/value/z).
    Automatic uses bounded daily matching for daily inputs, native dates for
    one other frequency, or completed months for mixed frequencies.
    Exact keeps original common dates. Daily uses weekdays and bounded backward
    matching; monthly selects the last observation in each completed month.
    Missing positive-weight inputs stay missing. No weights are redistributed.
    """
    mode, scale, alignment = (recipe.get(key) for key in ("weightMode", "scale", "alignment"))
    if mode not in ("equal", "custom") or scale not in ("raw", "standardized") or alignment not in ("auto", "exact", "daily", "monthly"):
        raise ValueError("Invalid composite weight, scale or alignment mode")
    carry = recipe.get("maxCarryDays", 7)
    if alignment == "daily" and (type(carry) is not int or not 0 <= carry <= 31):
        raise ValueError("Daily maximum age must be 0–31 calendar days")
    components, seen = [], set()
    for item in recipe.get("components", []):
        key = item.get("id")
        weight = 1 if mode == "equal" else item.get("weight")
        direction = item.get("direction")
        if not isinstance(key, str) or key in seen:
            raise ValueError("Invalid or duplicate composite indicator: " + str(key))
        if type(weight) not in (int, float) or not math.isfinite(weight) or weight < 0 or type(direction) not in (int, float) or direction not in (1, -1):
            raise ValueError("Invalid weight/direction: " + key)
        series = inputs.get(key)
        if not series and weight == 0 and key in SERIES:
            series = dict(id=key, name=SERIES[key]["name"], unit="", original_unit="", frequency=None,
                          transformation="Data not loaded (zero weight)", rows=[], raw_ids=[])
        if not series:
            raise ValueError("Data not loaded: " + key)
        seen.add(key)
        components.append(dict(id=key, enteredWeight=weight, direction=direction, series=series))
    total = math.fsum(item["enteredWeight"] for item in components)
    if not total or not math.isfinite(total):
        raise ValueError("At least one finite positive weight is required")
    for item in components:
        item["weight"] = item["enteredWeight"] / total
    active = [item for item in components if item["weight"] > 0]
    frequencies = {item["series"]["frequency"] for item in active}
    if alignment == "auto":
        alignment = "daily" if frequencies == {"Daily"} else "exact" if len(frequencies) == 1 else "monthly"
    if alignment == "daily" and (type(carry) is not int or not 0 <= carry <= 31):
        raise ValueError("Daily maximum age must be 0–31 calendar days")
    if alignment == "daily" and frequencies != {"Daily"}:
        raise ValueError("Daily alignment requires daily input series")
    if alignment == "exact" and len(frequencies) != 1:
        raise ValueError("Common dates require the same native frequency")
    if alignment == "monthly" and "Quarterly" in frequencies:
        raise ValueError("Quarterly inputs require exact dates")
    if scale == "raw" and (len({item["series"]["unit"].lower().replace("percent", "%") for item in active}) != 1
                           or len({item["series"]["transformation"] for item in active}) != 1):
        raise ValueError("Original inputs need comparable units and transformations; use standardized values")
    cutoff = str(cutoff)
    maps, ordered = [], []
    for item in components:
        lookup = {}
        for row in sorted(item["series"]["rows"], key=lambda row: row["date"]):
            if row["date"] > cutoff or (alignment == "monthly" and row["date"][:7] >= cutoff[:7]):
                continue
            if alignment == "daily" and row.get("value") is None:
                continue
            key = row["date"][:7] + "-01" if alignment == "monthly" else row["date"]
            lookup[key] = row
        maps.append(lookup)
        ordered.append(sorted(lookup))
    days = sorted({day for item, lookup in zip(components, maps) if item["weight"] > 0 for day in lookup})
    if alignment == "daily" and days:
        first, last = date.fromisoformat(days[0]), date.fromisoformat(days[-1])
        days = [str(first + timedelta(days=n)) for n in range((last-first).days+1)
                if (first + timedelta(days=n)).weekday() < 5]
    points = []
    for day in days:
        parts = []
        for item, lookup, source_dates in zip(components, maps, ordered):
            row = lookup.get(day)
            if alignment == "daily":
                index = bisect.bisect_right(source_dates, day) - 1
                row = lookup[source_dates[index]] if index >= 0 else None
                if row and (date.fromisoformat(day)-date.fromisoformat(row["date"])).days > carry:
                    row = None
            value = row.get("value" if scale == "raw" else "z") if row else None
            contribution = 0 if item["weight"] == 0 else None if value is None else item["weight"] * item["direction"] * value
            parts.append(dict(id=item["id"], sourceDate=row["date"] if row else None,
                              ageDays=(date.fromisoformat(day)-date.fromisoformat(row["date"])).days if row else None,
                              carried=bool(alignment == "daily" and row and row["date"] != day),
                              original=row.get("original_value") if row else None, raw=row.get("value") if row else None,
                              input=value, contribution=contribution))
        value = None if any(part["contribution"] is None for part in parts) else math.fsum(part["contribution"] for part in parts)
        if value is not None and not math.isfinite(value):
            raise ValueError("Composite calculation overflow")
        points.append(dict(date=day, value=value, parts=parts))
    return dict(points=points, components=components, unit=active[0]["series"]["unit"] if scale == "raw" else "weighted z-score",
                frequency="Monthly" if alignment == "monthly" else active[0]["series"]["frequency"],
                alignment=alignment, maxCarryDays=carry)


def date_coverage(rows, asof):
    """Calendar diagnostics, not an assertion that holidays are missing data."""
    output = []
    for key in sorted({row["series_id"] for row in rows}):
        own = [row for row in rows if row["series_id"] == key and row["observation_date"] <= str(asof)]
        valid = sorted(row["observation_date"] for row in own if row["value"] is not None)
        if not valid:
            continue
        first, last = date.fromisoformat(valid[0]), date.fromisoformat(valid[-1])
        weekdays = {str(first+timedelta(days=n)) for n in range((last-first).days+1) if (first+timedelta(days=n)).weekday()<5}
        missing = sorted(weekdays-set(valid)) if SERIES[key]["frequency"] == "Daily" else []
        output.append(dict(series_id=key, frequency=SERIES[key]["frequency"], first_date=valid[0], last_date=valid[-1],
                           numeric_points=len(valid), null_rows=len(own)-len(valid), missing_weekdays=len(missing),
                           recent_missing_dates="; ".join(missing[-10:])))
    return output

# ============================================================================
# 3. DRAW — static report charts. The interactive builder has its own renderer.
# ============================================================================
def escape(value):
    return html.escape(str(value), quote=True)


def svg_text(x, y, text, anchor="start", size=14, color="#656b73"):
    return '<text x="{}" y="{}" text-anchor="{}" font-size="{}" fill="{}">{}</text>'.format(x, y, anchor, size, color, escape(text))


def line_chart(series, stats, unit, height=285, frequency="Daily"):
    """series: [(label, [(date, value), ...], color)]. No JavaScript needed."""
    width = 960
    left, right, top, bottom = 62, 34, 32, 36
    end = date.fromisoformat(stats["date"])
    start = years_before(end, HISTORY_YEARS)
    values = [v for _, points, _ in series for d, v in points if start <= d <= end and v is not None]
    if not values:
        return ""
    if stats.get("sufficient"):
        values += [stats["p10"], stats["p90"]]
    low, high = min(values), max(values)
    spread = high - low or max(abs(high) * .1, 1)
    low, high = low - .12 * spread, high + .14 * spread
    x = lambda d: left + (d - start).days / max(1, (end - start).days) * (width - left - right)
    y = lambda v: top + (high - v) / (high - low) * (height - top - bottom)
    pieces = ['<svg xmlns="http://www.w3.org/2000/svg" font-family="Arial,sans-serif" viewBox="0 0 {} {}" role="img" aria-label="{} history"><rect width="100%" height="100%" fill="white"/>'.format(width, height, escape(series[0][0]))]
    for n, (label, _, color) in enumerate(series):
        pieces.append('<line x1="{}" x2="{}" y1="13" y2="13" stroke="{}" stroke-width="2"/>'.format(left + n * 235, left + 16 + n * 235, color))
        pieces.append(svg_text(left + 22 + n * 235, 17, label, size=13))
    if stats.get("sufficient"):
        pieces.append('<rect x="{}" y="{:.2f}" width="{}" height="{:.2f}" fill="#edf1f6"/>'.format(left, y(stats["p90"]), width - left - right, y(stats["p10"]) - y(stats["p90"])))
        pieces.append('<line x1="{}" x2="{}" y1="{:.2f}" y2="{:.2f}" stroke="#b5bdc8" stroke-dasharray="4 4"/>'.format(left, width - right, y(stats["median"]), y(stats["median"])))
    for n in range(5):
        value = low + (high - low) * n / 4
        pieces.append('<line x1="{}" x2="{}" y1="{:.2f}" y2="{:.2f}" stroke="#e3e7eb" stroke-width=".7"/>'.format(left, width - right, y(value), y(value)))
        label = format(value / 1000, ".1f") + "k" if abs(value) >= 10000 else number(value)
        pieces.append(svg_text(left - 9, round(y(value) + 4, 2), label, "end"))
    if low < 0 < high:
        pieces.append('<line x1="{}" x2="{}" y1="{:.2f}" y2="{:.2f}" stroke="#8d97a4" stroke-dasharray="3 3"/>'.format(left, width - right, y(0), y(0)))
    for label, points, color in series:
        path, pen, last = [], False, None
        for day, value in points:
            if not start <= day <= end:
                continue
            if value is None:
                # Daily histories may include a null for a non-common holiday.
                # Join nearby observations without adding an invented value.
                if frequency != "Daily":
                    pen = False
                continue
            gap_limit = {"Daily": 7, "Weekly": 8, "Monthly": 32, "Quarterly": 93}[frequency]
            if last and (day - last[0]).days > gap_limit:
                pen = False
            path.append(("L" if pen else "M") + "{:.2f},{:.2f}".format(x(day), y(value)))
            pen, last = True, (day, value)
        pieces.append('<path d="{}" fill="none" stroke="{}" stroke-width="{}"/>'.format(" ".join(path), color, 2 if color == BLUE else 1.2))
        if last:
            pieces.append('<circle cx="{:.2f}" cy="{:.2f}" r="3" fill="{}"><title>{}: {} on {}</title></circle>'.format(x(last[0]), y(last[1]), color, escape(label), number(last[1]), last[0]))
    ticks = sorted(set(range(0, HISTORY_YEARS + 1, max(1, HISTORY_YEARS // 5))) | {HISTORY_YEARS}, reverse=True)
    for years in ticks:
        day = years_before(end, years)
        pieces.append(svg_text(round(x(day), 1), height - 12, str(day)[:7], "middle"))
    pieces.append(svg_text(width - right, height - 1, unit, "end", 10))
    return "".join(pieces) + "</svg>"


def range_strip(stats):
    if not stats or not stats["sufficient"]:
        return ""
    lo, hi = min(stats["min"], stats["value"]), max(stats["max"], stats["value"])
    span = hi - lo or 1
    x = lambda v: 55 + (v - lo) / span * 820
    parts = ['<svg class="range-strip" viewBox="0 0 960 92" role="img" aria-label="Historical range and current value">']
    parts.append('<line x1="55" x2="875" y1="35" y2="35" stroke="#c5ccd5" stroke-width="3"/>')
    parts.append('<line x1="{:.2f}" x2="{:.2f}" y1="35" y2="35" stroke="#9cabbf" stroke-width="10"/>'.format(x(stats["p10"]), x(stats["p90"])))
    for key in ("min", "p10", "median", "p90", "max"):
        parts.append('<line x1="{0:.2f}" x2="{0:.2f}" y1="28" y2="43" stroke="#7b8794"/>'.format(x(stats[key])))
    parts.append('<circle cx="{:.2f}" cy="35" r="6" fill="{}"/>'.format(x(stats["value"]), BLUE))
    parts.append(svg_text(min(850, max(85, x(stats["value"]))), 16, "Current " + number(stats["value"]), "middle", 13, BLUE))
    # Fixed columns prevent numeric labels colliding when P10 is near the minimum.
    for n, (key, label) in enumerate([( "min", "Prior min"), ("p10", "P10"), ("median", "Median"), ("p90", "P90"), ("max", "Prior max")]):
        parts.append(svg_text(55 + n * 205, 65, label, "middle", 13))
        parts.append(svg_text(55 + n * 205, 84, number(stats[key]), "middle", 15, "#20252c"))
    return "".join(parts) + "</svg>"


def band_description(stats):
    if not stats["sufficient"]:
        return "Prior sample {}–{} · n={:,} · insufficient history for P10–P90.".format(
            stats["sample_start"] or "—", stats["sample_end"] or "—", stats["count"])
    return ("Prior sample {}–{} · n={:,} · P10 {} / P90 {} · dashed median {}. "
            "Latest excluded; fixed historical band, not a confidence interval.").format(
                stats["sample_start"], stats["sample_end"], stats["count"],
                number(stats["p10"]), number(stats["p90"]), number(stats["median"]))


def metric_panel(metric, extra_series=None, full=False):
    stats = metric["stats"]
    series = extra_series or [(metric["name"], metric["points"], BLUE)]
    changes_html = "".join('<div><span>{} change</span><strong>{} {}</strong><small>Base: {}</small></div>'.format(
        item["label"], number(item["value"], True), escape(metric["change_unit"]), item["base"] or "unavailable") for item in metric["comparisons"])
    title = ' data-id="{}"'.format(escape(metric["id"]))
    return '''<article class="chart-panel {full}"{identity}>
      <div class="chart-heading"><div><h3>{name}</h3><p>{transform} · {unit} · {frequency} · {date}</p></div>
      <div class="downloads"><button data-chart-download="{key}">Download data</button></div></div>
      {chart}<div class="readings"><div><span>Latest</span><strong>{value} {unit}</strong><small>Level percentile: {rank}</small></div>{changes}</div>
      {range}<p class="method">{method}</p>
      <p class="source">{source}</p>
    </article>'''.format(full="full" if full else "", identity=title, name=escape(metric["name"]), transform=escape(metric["transformation"]),
                          unit=escape(metric["unit"]), frequency=metric["frequency"], date=stats["date"], key=escape(metric["id"]),
                          chart=line_chart(series, stats, metric["unit"], frequency=metric["frequency"]), value=number(stats["value"]),
                          rank=number(stats.get("percentile")) + "%" if stats["sufficient"] else "insufficient history",
                          changes=changes_html, range=range_strip(stats),
                          method=escape(band_description(stats)), source=escape(metric["source"]))


# ============================================================================
# 4. BUILD — local raw -> saved analysis -> HTML. All sections are expanded.
# ============================================================================
def make_derived_metrics(metrics):
    """Matched-date arithmetic, kept separate from expectation surprises."""
    pairs = [
        ("YIELD_CURVE", "10Y minus 2Y Treasury yield", "BBG_10Y" if {"BBG_10Y", "BBG_2Y"} <= metrics.keys() else "DGS10",
         "BBG_2Y" if {"BBG_10Y", "BBG_2Y"} <= metrics.keys() else "DGS2", "subtract", "Market pricing", "%", "Daily"),
        ("IMPLIED_BE", "10Y nominal minus real yield", "DGS10", "DFII10", "subtract", "Rates", "%", "Daily"),
        ("LABOR_RATIO", "Job openings / unemployed persons", "JTSJOL", "UNEMPLOY", "ratio", "Labor", "times", "Monthly"),
        ("INCOME_GAP", "Real spending minus income momentum", "PCEC96", "DSPIC96", "subtract", "Growth", "pp", "Monthly"),
    ]
    for key, name, a, b, operation, group, unit, frequency in pairs:
        if a not in metrics or b not in metrics:
            continue
        # A configuration edit must not silently change a predefined formula.
        left_cfg, right_cfg = SERIES[a], SERIES[b]
        transforms = {left_cfg["transform"], right_cfg["transform"]}
        units = {metrics[item]["unit"].lower().replace("percent", "%") for item in (a, b)}
        compatible = all(metrics[item]["frequency"] == frequency for item in (a, b))
        if key == "INCOME_GAP":
            compatible &= len(transforms) == 1 and transforms <= {"yoy", "annualized_3m", "annualized_6m"} and units == {"%"}
        elif key == "LABOR_RATIO":
            compatible &= transforms == {"level"} and units == {"thousands of persons"}
        else:
            compatible &= transforms == {"level"} and units == {"%"}
        if not compatible:
            raise ValueError(key + ": input definitions no longer match this predefined formula. Check " + a + " / " + b + " units, frequency and transformation; use new IDs for alternative definitions.")
        left, right = dict(metrics[a]["points"]), dict(metrics[b]["points"])
        points = []
        for day in sorted(left.keys() & right.keys()):
            x, y = left[day], right[day]
            value = None if x is None or y is None else (x - y if operation == "subtract" else x / y if y else None)
            points.append((day, value))
        note = "Same-date values only. " + ("Ratio of job openings to unemployed people; openings are not hires." if operation == "ratio" else "Difference of the two plotted quantities; it does not establish causality.")
        if key == "IMPLIED_BE":
            note += " Includes inflation-risk and liquidity premia; not pure expected inflation or a term-premium estimate."
        cfg = dict(name=name, group=group, unit=unit, frequency=frequency, transform="level", note=note)
        result = analyze_series(key, cfg, points, metrics[a]["source"] + " minus " + metrics[b]["source"] if operation == "subtract" else metrics[a]["source"] + " divided by " + metrics[b]["source"])
        if result:
            result["raw_ids"] = sorted(set(metrics[a].get("raw_ids", [a]) + metrics[b].get("raw_ids", [b])))
            metrics[key] = result


def chart_export(title, traces, reference_stats, raw_rows, statistics_rows):
    """Three CSVs for one chart; plotted traces only, source history kept native.

    traces contains (series_id, transformation, points, unit) tuples. The chart
    CSV uses the same ten-year window and endpoint as the SVG. Raw input rows
    include earlier observations needed to reproduce growth rates and changes.
    """
    end = date.fromisoformat(reference_stats["date"])
    start = years_before(end, HISTORY_YEARS)
    points = [dict(series_id=key, transformation=label, observation_date=str(day), value=value, unit=unit)
              for key, label, history, unit in traces for day, value in history if start <= day <= end]
    stat_columns = ["series_id", "transformation", "date", "value", "window_start", "sample_start", "sample_end", "count",
                    "sufficient", "min", "p10", "median", "p90", "max", "percentile", "zscore", "change_label", "change",
                    "change_base", "change_unit", "change_percentile", "summary"]
    contents = {
        "chart_data.csv": csv_string(["series_id", "transformation", "observation_date", "value", "unit"], points),
        "raw_data.csv": csv_string(RAW_COLUMNS, raw_rows),
        "statistics.csv": csv_string(stat_columns, statistics_rows),
    }
    return {"title": title, "files": {name: base64.b64encode(text.encode("utf-8-sig")).decode("ascii")
                                      for name, text in contents.items()}}


def publish_report(files):
    """Stage all outputs, then replace them; restore the prior set on an error.

    Keep staging on the project filesystem so each rename is atomic. This
    handles caught write/rename failures, not a power loss or a forced kill.
    build_record.json contains hashes to detect an interrupted publication.
    Run one build at a time; close analysis CSVs in Excel before rebuilding.
    """
    staging = Path(tempfile.mkdtemp(prefix=".build-", dir=ROOT))
    committed, backups, rollback_errors = [], {}, []
    try:
        for index, (target, contents) in enumerate(files.items()):
            target.parent.mkdir(parents=True, exist_ok=True)
            with (staging / str(index)).open("w", encoding="utf-8", newline="") as stream:
                stream.write(contents)
            if target.exists():
                backup = staging / (str(index) + ".backup")
                shutil.copyfile(target, backup)
                backups[target] = backup
        for index, target in enumerate(files):
            (staging / str(index)).replace(target)
            committed.append(target)
    except BaseException:
        for target in reversed(committed):
            try:
                if target in backups:
                    backups[target].replace(target)
                else:
                    target.unlink()
            except OSError as error:
                rollback_errors.append(str(error))
        if rollback_errors:
            print("RECOVERY NEEDED: prior output backups retained in", staging, "|", "; ".join(rollback_errors))
        raise
    finally:
        if not rollback_errors:
            shutil.rmtree(staging, ignore_errors=True)


def build_report(raw_path, asof):
    # Validate all inputs before replacing any prior report outputs.
    if not isinstance(HISTORY_YEARS, int) or HISTORY_YEARS < 1:
        raise ValueError("HISTORY_YEARS must be a positive integer")
    if set(MIN_OBSERVATIONS) != {"Daily", "Weekly", "Monthly", "Quarterly"} or any(
            not isinstance(value, int) or value < 2 for value in MIN_OBSERVATIONS.values()):
        raise ValueError("MIN_OBSERVATIONS must define each frequency with an integer >= 2")
    template = (ROOT / "dashboard_template.html").read_text(encoding="utf-8-sig")
    required_tokens = ("__ASOF__", "__NAV__", "__CONTENT__", "__NOTICE__", "__DOWNLOADS__", "__CHART_EXPORTS__", "__COMPOSITE_DATA__")
    if any(token not in template for token in required_tokens):
        raise ValueError("HTML template is missing required data placeholders; use the matching dashboard_template.html")
    raw_hash = hashlib.sha256(raw_path.read_bytes()).hexdigest() if raw_path else None
    rows = read_raw(raw_path) if raw_path else []
    if not any(row["value"] is not None and row["observation_date"] <= str(asof) for row in rows):
        raise ValueError("No numeric raw observations on or before the report cutoff. Check the CSV and --asof; prior outputs were kept.")
    manifest = read_manifest(raw_path, rows) if raw_path else {}
    metrics = analyze(rows, asof)
    if not metrics:
        raise ValueError("No usable chart values after transformation. Download more history or check the definitions; prior outputs were kept.")
    make_derived_metrics(metrics)
    # Persist results used by charts. Inspect these CSVs before opening the HTML.
    analysis_rows = []
    statistics_rows = []
    for key, metric in metrics.items():
        for day, value in metric["points"]:
            analysis_rows.append(dict(series_id=key, transformation=metric["transformation"], observation_date=str(day), value=value, unit=metric["unit"]))
        for day, value in metric["move_points"]:
            analysis_rows.append(dict(series_id=key, transformation=metric["comparisons"][1]["label"] + " change",
                                      observation_date=str(day), value=value, unit=metric["change_unit"]))
        if key in ("CPILFESL", "PCEPILFE"):
            raw_points = sorted((date.fromisoformat(r["observation_date"]), r["value"]) for r in rows
                                if r["series_id"] == key and date.fromisoformat(r["observation_date"]) <= asof)
            for kind, label in (("annualized_3m", "3M annualized growth"), ("annualized_6m", "6M annualized growth")):
                analysis_rows.extend(dict(series_id=key, transformation=label, observation_date=str(day), value=value, unit="%")
                                     for day, value in transform(raw_points, kind))
        stats = metric["stats"]
        statistics_rows.append(dict(series_id=key, **stats, change_label=metric["comparisons"][1]["label"],
                                    change=metric["comparisons"][1]["value"], change_base=metric["comparisons"][1]["base"],
                                    change_unit=metric["change_unit"], change_percentile=(metric["move_stats"] or {}).get("percentile"),
                                    summary=metric["summary"]))
    output = ROOT / "data" / "analysis"
    analysis_csv = csv_string(["series_id", "transformation", "observation_date", "value", "unit"], analysis_rows)
    stat_columns = ["series_id", "date", "value", "window_start", "sample_start", "sample_end", "count", "sufficient", "min", "p10", "median", "p90", "max", "percentile", "zscore", "change_label", "change", "change_base", "change_unit", "change_percentile", "summary"]
    stats_csv = csv_string(stat_columns, statistics_rows)
    composite_data, composite_rows = prepare_composite_inputs(metrics, asof, rows)
    composite_csv = csv_string(["series_id"] + composite_data["columns"], composite_rows)
    # Per-input raw histories support saved monitors. Static charts each have
    # their own three-file bundle; no global dataset downloads are embedded.
    downloads, exports = {}, {}
    for key, metric in metrics.items():
        raw_ids = set(metric.get("raw_ids", [key]))
        own_raw = [r for r in rows if r["series_id"] in raw_ids and date.fromisoformat(r["observation_date"]) <= asof]
        downloads["raw_" + key] = csv_string(RAW_COLUMNS, own_raw)
        if key != "AVERAGE":
            own_stats = [{**r, "transformation": metric["transformation"]} for r in statistics_rows if r["series_id"] == key]
            exports[key] = chart_export(metric["name"], [(key, metric["transformation"], metric["points"], metric["unit"])],
                                        metric["stats"], own_raw, own_stats)
    sections, navigation = [], []
    average = metrics.get("AVERAGE")
    groups = list(dict.fromkeys(m["group"] for m in metrics.values() if m["id"] != "AVERAGE"))
    preferred = ["US surprises", "Growth", "Consumer", "Labor", "Inflation", "Market pricing", "Rates", "Credit", "Financial conditions", "Housing", "Markets", "Liquidity"]
    for group_number, group in enumerate(sorted(groups, key=lambda g: preferred.index(g) if g in preferred else len(preferred))):
        selected = [m for m in metrics.values() if m["group"] == group and m["id"] != "AVERAGE"]
        section_id = "group-" + str(group_number) + "-" + re.sub(r"[^a-z0-9]+", "-", group.lower()).strip("-")
        content = "".join(metric_panel(m) for m in selected)
        sections.append('<section id="{}"><div class="section-heading"><h2>{}</h2></div><div class="charts">{}</div></section>'.format(section_id, escape(group), content))
        navigation.append('<a href="#{}">{}</a>'.format(section_id, escape(group)))
    # Realized inflation momentum comparison. No estimates or fake release history.
    inflation_panels = []
    for key in ("CPILFESL", "PCEPILFE"):
        own = sorted((date.fromisoformat(r["observation_date"]), r["value"]) for r in rows if r["series_id"] == key and date.fromisoformat(r["observation_date"]) <= asof)
        if not own or key not in metrics:
            continue
        three = transform(own, "annualized_3m")
        stats = range_statistics(three, "Monthly")
        if not stats:
            continue
        traces = [("3M annualized", three, BLUE), ("6M annualized", transform(own, "annualized_6m"), GRAY), ("YoY", transform(own, "yoy"), LIGHT)]
        chart = line_chart(traces, stats, "%", frequency="Monthly")
        export_key = key + "_MOMENTUM"
        title = SERIES[key]["name"] + " · momentum horizons"
        raw_ids = set(metrics[key].get("raw_ids", [key]))
        own_raw = [r for r in rows if r["series_id"] in raw_ids and date.fromisoformat(r["observation_date"]) <= asof]
        horizon_stats = []
        readings = []
        for label, points, _ in traces:
            same_endpoint = [(day, value) for day, value in points if str(day) <= stats["date"]]
            horizon = range_statistics(same_endpoint, "Monthly")
            if horizon:
                horizon_stats.append(dict(series_id=key, transformation=label, **horizon))
                readings.append('<div><span>{}</span><strong>{} %</strong><small>{}</small></div>'.format(
                    escape(label), number(horizon["value"]), horizon["date"]))
        exports[export_key] = chart_export(title, [(key, label, points, "%") for label, points, _ in traces],
                                          stats, own_raw, horizon_stats)
        inflation_panels.append('<article class="chart-panel" data-id="' + escape(export_key) + '"><div class="chart-heading"><h3>'
            + escape(title) + '</h3><button data-chart-download="' + escape(export_key) + '">Download data</button></div>'
            + chart + '<div class="readings">' + ''.join(readings) + '</div>' + range_strip(stats)
            + '<p class="method">3M band · ' + escape(band_description(stats)) + '</p><p class="source">'
            + escape(metrics[key]["source"]) + '</p></article>')
    if inflation_panels:
        sections.append('<section id="inflation-momentum"><h2>Inflation momentum</h2><div class="charts">' + ''.join(inflation_panels) + '</div></section>')
        navigation.append('<a href="#inflation-momentum">Inflation momentum</a>')
    source = "No local observations loaded." if not raw_path else "Local source: " + str(raw_path.relative_to(ROOT) if raw_path.is_relative_to(ROOT) else raw_path)
    if not average:
        print("SOURCE STATUS: Employment/inflation pair unavailable; saved Average monitor needs both source histories.")
    notices = []
    for key, metric in metrics.items():
        lag = (asof - date.fromisoformat(metric["stats"]["date"])).days
        limit = {"Daily": 7, "Weekly": 21, "Monthly": 80, "Quarterly": 180}[metric["frequency"]]
        if lag > limit:
            notices.append(metric["name"] + ": latest observation " + metric["stats"]["date"] + " (" + str(lag) + " days before report date).")
    if manifest:
        notices.extend("Bloomberg " + key + ": " + str(error) for key, error in manifest.get("errors", {}).items())
    quality_note = (str(len(notices)) + " data freshness / availability flags. Check chart dates; details are saved in build_record.json.") if notices else ""
    replacements = {
        "__ASOF__": str(asof), "__SOURCE__": escape(source), "__NAV__": "".join(navigation),
        "__CONTENT__": "".join(sections),
        "__NOTICE__": '<div class="notice">' + escape(quality_note) + '</div>' if quality_note else "",
        "__DOWNLOADS__": json.dumps({key: base64.b64encode(value.encode("utf-8-sig")).decode("ascii") for key, value in downloads.items()}),
        "__CHART_EXPORTS__": json.dumps(exports, allow_nan=False, separators=(",", ":")).replace("<", "\\u003c"),
        # Escape '<' so an external series label cannot close the script element.
        "__COMPOSITE_DATA__": json.dumps(composite_data, allow_nan=False, separators=(",", ":")).replace("<", "\\u003c"),
    }
    # One pass: a label containing a placeholder must never become another
    # template substitution. All inserted labels are already HTML/JSON escaped.
    template = re.sub("|".join(re.escape(token) for token in replacements), lambda match: replacements[match.group()], template)
    if raw_path and hashlib.sha256(raw_path.read_bytes()).hexdigest() != raw_hash:
        raise ValueError("Raw CSV changed during this build. Rerun using a stable snapshot; prior outputs were kept.")
    files = {output / "chart_data.csv": analysis_csv, output / "statistics.csv": stats_csv,
             output / "composite_inputs.csv": composite_csv, ROOT / "dashboard.html": template}
    coverage = date_coverage(rows, asof)
    files[output / "date_coverage.csv"] = csv_string(["series_id", "frequency", "first_date", "last_date", "numeric_points", "null_rows", "missing_weekdays", "recent_missing_dates"], coverage)
    record = dict(asof=str(asof), raw_path=str(raw_path) if raw_path else None, raw_sha256=raw_hash,
                  generated_at=datetime.now(timezone.utc).isoformat(), chart_count=len(metrics), data_flags=notices,
                  output_sha256={path.relative_to(ROOT).as_posix(): hashlib.sha256(contents.encode("utf-8")).hexdigest()
                                 for path, contents in files.items()})
    files[output / "build_record.json"] = json.dumps(record, indent=2, ensure_ascii=False, allow_nan=False)
    publish_report(files)
    print("ANALYSIS SAVED:", output)
    print("HTML SAVED:", ROOT / "dashboard.html", "|", len(metrics), "observed/derived series")
    return metrics


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--raw", type=Path, help="Read this raw CSV instead of the latest successful snapshot")
    parser.add_argument("--asof", type=date.fromisoformat, default=date.today(), help="Report cutoff YYYY-MM-DD")
    args = parser.parse_args()
    if args.raw:
        raw_path = args.raw.resolve()
    else:
        pointer = ROOT / "data" / "latest_bloomberg.json"
        if not pointer.exists():
            raise RuntimeError("No saved Bloomberg snapshot. Run python fetch_data.py first, or use --raw PATH.")
        saved = json.loads(pointer.read_text(encoding="utf-8-sig"))
        if not isinstance(saved, dict) or not isinstance(saved.get("raw_csv"), str) or not saved["raw_csv"]:
            raise ValueError("Invalid latest_bloomberg.json: expected a raw_csv path. Run fetch_data.py again.")
        # Accept historical Windows pointers after moving a project to macOS.
        relative = Path(saved["raw_csv"].replace("\\", "/"))
        if relative.is_absolute() or ".." in relative.parts:
            raise ValueError("Snapshot pointer must be relative to this project. Use --raw PATH for an external CSV.")
        raw_path = ROOT / relative
    build_report(raw_path, args.asof)


if __name__ == "__main__":
    try:
        main()
    except (ValueError, RuntimeError, OSError, ArithmeticError) as error:
        raise SystemExit("ERROR: " + str(error))
