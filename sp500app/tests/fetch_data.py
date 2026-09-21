#!/usr/bin/env python3
"""Step 1: python fetch_data.py — Bloomberg -> timestamped local raw CSV.

Edit SERIES below. Only this script calls Bloomberg; it never builds HTML.
Step 2 is python build_dashboard.py. Both scripts and dashboard_template.html
must be in the same directory. Requires Python 3.9+ and your existing blpapi.
"""
import argparse
import csv
import hashlib
import io
import json
import math
import re
import tempfile
import time
from datetime import date, datetime, timezone
from pathlib import Path

# ============================================================================
# 1. CONFIGURATION — edit tickers / labels here, not in the HTML.
# ============================================================================
ROOT = Path(__file__).resolve().parent
HISTORY_YEARS = 10
BLOOMBERG_HOST = "127.0.0.1"
BLOOMBERG_PORT = 8194
REQUEST_TIMEOUT = 90
SERIES = {}


def add_series(key, name, ticker, group, unit="index points", frequency="Daily",
               transform="level", required=False, note="", enabled=True, field="PX_LAST"):
    """One indicator definition; use field= for a confirmed non-PX_LAST field."""
    if key in SERIES:
        raise ValueError("Duplicate series ID: " + key)
    SERIES[key] = dict(name=name, ticker=ticker, field=field, group=group,
                       unit=unit, frequency=frequency, transform=transform,
                       required=required, note=note, enabled=enabled)


# User's pair. Keep both original values; do not rescale before averaging.
add_series("SURPRISE_EMPLOYMENT", "US employment surprise", "ECSULBUS Index", "US surprises", required=True,
           note="User-supplied Bloomberg ECO labor index. Verify sign and smoothing in Terminal.")
add_series("SURPRISE_INFLATION", "US inflation surprise", "BCMPUSIF Index", "US surprises", required=True,
           note="Bloomberg Economics inflation surprise; confirm current Terminal definition.")
add_series("SURPRISE_GROWTH", "US growth surprise", "BCMPUSGR Index", "US surprises",
           note="Compare growth news with employment news; they are overlapping information, not independent votes.")
add_series("SURPRISE_US_OVERALL", "US economic surprise", "ECSURPUS Index", "US surprises",
           note="Bloomberg ECO headline index. Compare direction with the growth series; do not pool the two.")

# Price context. These are separate market series, never labelled macro surprises.
add_series("BBG_2Y", "US 2Y Treasury yield", "USGG2YR Index", "Market pricing", "%",
           note="Generic market yield; distinct from the FRED constant-maturity series. Changes shown in bp.")
add_series("BBG_10Y", "US 10Y Treasury yield", "USGG10YR Index", "Market pricing", "%",
           note="Generic market yield; not a pure measure of the expected Fed policy path. Changes shown in bp.")
add_series("SPX", "S&P 500 price index", "SPX Index", "Market pricing",
           note="Price index, not total return. Changes are in index points, not percentage returns or release attribution.")
add_series("BBG_VIX", "VIX", "VIX Index", "Market pricing",
           note="Implied equity volatility. No automatic equity direction follows from a surprise sign.")

# Existing real public history remains readable. Blank ticker means no Bloomberg request.
# To add an economic series: confirm its date basis, units and seasonality first.
# The monthly transforms below require month-start ECONOMIC REFERENCE DATES,
# not announcement dates or month-end samples from a Bloomberg DAILY history.
# Safer alternative: add a new ID with a directly reported growth-rate field and level transform.
add_series('PAYEMS', 'Nonfarm payrolls', "", 'Labor', 'Thousands of persons', 'Monthly', 'avg_change_3m',
           note='Consumer income and labor demand; check revisions alongside the latest payroll change. Payroll history is revised. A three-month average smooths noise but can delay turning points.')
add_series('UNRATE', 'Unemployment rate', "", 'Labor', 'Percent', 'Monthly', 'level',
           note='Labor-market slack and risks to household income and earnings. A rising unemployment rate can reflect both weaker hiring and changes in labor-force participation.')
add_series('ICSA', 'Initial jobless claims', "", 'Labor', 'Number of claims', 'Weekly', 'avg_4w',
           note='A timely cross-check on layoffs and downside risks to consumption. Holiday timing and seasonal adjustment can move weekly claims; use the four-week average.')
add_series('CPILFESL', 'Core CPI', "", 'Inflation', 'Index 1982–1984=100', 'Monthly', 'yoy',
           note='Price pressure can affect policy expectations, discount rates and corporate margins. Core CPI excludes food and energy. A lower inflation rate still means prices are rising if inflation remains positive.')
add_series('PCEPILFE', 'Core PCE price index', "", 'Inflation', 'Index 2017=100', 'Monthly', 'yoy',
           note='An additional inflation measure for assessing the policy and valuation backdrop. PCE and CPI differ in coverage and weights; their observations may cover different months.')
add_series('PCEC96', 'Real personal consumption', "", 'Growth', 'Billions of chained 2017 dollars', 'Monthly', 'annualized_3m',
           note='Real household demand provides context for consumer-facing revenue assumptions. Three-month annualized growth can be volatile. Real spending excludes price changes and is subject to revision.')
add_series('INDPRO', 'Industrial production', "", 'Growth', 'Index 2017=100', 'Monthly', 'yoy',
           note='Production momentum helps test the demand backdrop for cyclical businesses. Industrial production covers only part of the economy and is not a complete proxy for S&P 500 earnings.')
add_series('DGS2', '2-year Treasury yield', "", 'Rates', 'Percent', 'Daily', 'level',
           note='A market-based cross-check on the near-term rates and policy backdrop. A Treasury yield is not a pure measure of the expected policy path; several components can move it.')
add_series('DGS10', '10-year Treasury yield', "", 'Rates', 'Percent', 'Daily', 'level',
           note='Long-term nominal discount rates matter for equity valuation assumptions. Nominal yields reflect real yields, inflation compensation and premia; direction alone does not establish the cause.')
add_series('DFII10', '10-year real Treasury yield', "", 'Rates', 'Percent', 'Daily', 'level',
           note='A market-based real discount-rate input for reviewing valuation assumptions. TIPS yields include liquidity and risk-premium effects; a yield move does not imply a fixed equity response.')
add_series('BAMLH0A0HYM2', 'US high-yield option-adjusted spread', "", 'Credit', 'Percent', 'Daily', 'level',
           note='Credit conditions help monitor financing pressure and risk appetite. OAS is a credit-market indicator, not an equity risk premium. Available FRED history may be limited.')
add_series('DTWEXBGS', 'Broad US dollar index', "", 'Markets', 'Index January 2006=100', 'Daily', 'level',
           note='Currency moves can affect translated overseas earnings and import costs. Company effects depend on revenue mix, production location and hedging. This is an index, not a currency return portfolio.')
add_series('DCOILWTICO', 'WTI crude oil spot price', "", 'Markets', 'US dollars per barrel', 'Daily', 'level',
           note='Energy prices can change producer revenues, business costs and household purchasing power. This is a spot-price series. Futures total returns differ; oil demand and supply shocks have different implications.')
add_series('CCSA', 'Continuing jobless claims', "", 'Labor', 'Number of persons', 'Weekly', 'level',
           note='Check whether workers take longer to find another job. Reporting lags and benefit eligibility affect the level; compare with initial claims.')
add_series('JTSQUR', 'Job quits rate', "", 'Labor', 'Percent', 'Monthly', 'level',
           note='Worker confidence and wage pressure provide context for labor-market tightness. Survey estimates are revised and can lag faster labor indicators.')
add_series('JTSJOL', 'Job openings', "", 'Labor', 'Thousands of persons', 'Monthly', 'level',
           note='Labor demand; compare openings with unemployed workers. Openings do not equal actual hiring and survey estimates are revised.')
add_series('UNEMPLOY', 'Unemployed persons', "", 'Labor', 'Thousands of persons', 'Monthly', 'level',
           note='Denominator for the openings-to-unemployed ratio. A count and a rate answer different questions; account for labor-force size.')
add_series('CES0500000003', 'Average hourly earnings, private employees', "", 'Labor', 'Dollars per hour', 'Monthly', 'yoy',
           note='Wage pressure affects household income and business costs. Average hourly earnings also move with workforce composition; compare with ECI or a wage tracker.')
add_series('DSPIC96', 'Real disposable personal income', "", 'Consumer', 'Billions of chained 2017 dollars', 'Monthly', 'annualized_3m',
           note='Compare real income momentum with consumption to assess spending support. Transfer payments can create temporary movements; divergence alone does not identify credit-financed spending.')
add_series('PSAVERT', 'Personal saving rate', "", 'Consumer', 'Percent', 'Monthly', 'level',
           note='A cross-check on the household spending and income balance. Aggregate saving masks differences across households and is revised.')
add_series('DGS30', '30-year Treasury yield', "", 'Rates', 'Percent', 'Daily', 'level',
           note='Long-duration discount rates and changes at the long end of the curve. Term-premium and policy-expectation effects need additional decomposition.')
add_series('T10YIE', '10-year breakeven inflation', "", 'Rates', 'Percent', 'Daily', 'level',
           note='Separate inflation compensation from the real-yield component of nominal rates. Breakeven includes inflation risk and liquidity premia; it is not pure expected inflation.')
add_series('NFCI', 'Chicago Fed financial conditions index', "", 'Financial conditions', 'Index', 'Weekly', 'level',
           note='A broad cross-check on risk, credit and leverage conditions. Positive is tighter than its historical average. It is a composite and includes market-based inputs.')
add_series('WALCL', 'Federal Reserve total assets', "", 'Liquidity', 'Millions of US dollars', 'Weekly', 'level',
           note='Track the size of the central bank balance sheet alongside its composition. Wednesday level. Asset changes are not a mechanical predictor of equities.')
add_series('WTREGEN', 'Treasury General Account', "", 'Liquidity', 'Millions of US dollars', 'Weekly', 'level',
           note='Treasury cash management is relevant to banking-system liquidity. Weekly average; do not subtract this unaligned series from a Wednesday Fed level and call it net liquidity.')
add_series('RRPONTSYD', 'Overnight reverse repo operations', "", 'Liquidity', 'Billions of US dollars', 'Daily', 'level',
           note='A cross-check on the allocation of money-market cash. A falling RRP balance need not flow into equities; units here are billions, not millions.')
add_series('PERMIT', 'Building permits', "", 'Housing', 'Thousands of units', 'Monthly', 'yoy',
           note='A forward-looking part of the residential activity pipeline. Permits may not lead immediately to starts; multifamily approvals can be volatile.')
add_series('HOUST', 'Housing starts', "", 'Housing', 'Thousands of units', 'Monthly', 'yoy',
           note='Rate-sensitive real activity and the residential construction pipeline. Weather and monthly sampling noise can dominate a single observation.')
add_series('VIXCLS', 'CBOE volatility index', "", 'Financial conditions', 'Index', 'Daily', 'level',
           note='Equity option-implied volatility provides risk-pricing context. This is a volatility index, not an equity return or a direct measure of funding liquidity.')
add_series('BAMLC0A0CM', 'US investment-grade corporate OAS', "", 'Credit', 'Percent', 'Daily', 'level',
           note='Investment-grade financing conditions complement high-yield stress measures. OAS is not an equity risk premium; index composition and duration differ from high yield.')

RAW_COLUMNS = ["series_id", "security", "field", "observation_date", "value", "unit", "source", "retrieved_at"]


def years_before(day, years):
    """Calendar years, including a leap-day fallback."""
    try:
        return day.replace(year=day.year - years)
    except ValueError:
        return day.replace(year=day.year - years, day=28)


def csv_string(columns, rows):
    stream = io.StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=columns, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return stream.getvalue()


def save_text_atomic(path, contents):
    """Replace a complete UTF-8 file in one step; never expose half-written CSV/JSON."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = None
    try:
        # Close before replace: Windows cannot replace an open temporary file.
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", newline="",
                                         dir=path.parent, prefix="." + path.name + ".",
                                         suffix=".tmp", delete=False) as stream:
            temporary_path = Path(stream.name)
            stream.write(contents)
        temporary_path.replace(path)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def save_csv(path, contents):
    """UTF-8 with exact CSV newlines, including on Windows (no doubled CR)."""
    save_text_atomic(path, contents)


def save_json(path, value):
    save_text_atomic(path, json.dumps(value, indent=2, ensure_ascii=False, allow_nan=False))


def configured_requests(start_day, end_day):
    """Catch editable configuration mistakes before opening a Bloomberg session."""
    if not isinstance(start_day, date) or not isinstance(end_day, date) or start_day > end_day:
        raise ValueError("Start and end must be dates, with start <= end.")
    if not isinstance(REQUEST_TIMEOUT, (int, float)) or not math.isfinite(REQUEST_TIMEOUT) or REQUEST_TIMEOUT <= 0:
        raise ValueError("REQUEST_TIMEOUT must be a positive number of seconds.")
    mappings = {}
    for key, cfg in SERIES.items():
        if not isinstance(key, str) or not re.fullmatch(r"[A-Z][A-Z0-9_]*", key):
            raise ValueError("Use uppercase letters, digits and underscores for series IDs: " + str(key))
        if not isinstance(cfg.get("enabled"), bool) or not isinstance(cfg.get("required"), bool):
            raise ValueError(key + ": enabled and required must be True or False.")
        if not isinstance(cfg.get("ticker"), str):
            raise ValueError(key + ": ticker must be a string (blank means no request).")
        if cfg["enabled"] and cfg["required"] and not cfg["ticker"].strip():
            raise ValueError(key + ": an enabled required indicator needs a ticker.")
        if not cfg["enabled"] or not cfg["ticker"].strip():
            continue
        for field in ("ticker", "field", "name", "unit", "group"):
            if not isinstance(cfg.get(field), str) or not cfg[field].strip() or cfg[field] != cfg[field].strip():
                raise ValueError(key + ": " + field + " must be nonempty text without leading/trailing spaces.")
        if cfg.get("frequency") not in ("Daily", "Weekly", "Monthly", "Quarterly"):
            raise ValueError(key + ": frequency must be Daily, Weekly, Monthly or Quarterly.")
        # A historical announcement date must not become an economic reference month.
        if cfg.get("transform") != "level":
            raise ValueError(key + ": configured Bloomberg histories must use level transform. Import validated reference-period CSV data or use a directly reported rate field.")
        mappings[key] = dict(cfg)
    if not mappings:
        raise ValueError("No enabled Bloomberg tickers in section 1.")
    return mappings


# ============================================================================
# 2. DOWNLOAD — save source data before any transformation or chart calculation.
# ============================================================================
def response_rows(message, key, cfg, retrieved_at, start_day, end_day):
    """Validate one matching response before accepting its source observations."""
    if message.hasElement("responseError"):
        raise RuntimeError(str(message.getElement("responseError")))
    if not message.hasElement("securityData"):
        raise ValueError("Bloomberg response is missing securityData.")
    security = message.getElement("securityData")
    for error_key in ("securityError", "fieldExceptions"):
        if security.hasElement(error_key):
            error = security.getElement(error_key)
            if error_key == "securityError" or error.numValues():
                raise RuntimeError(str(error))
    if security.hasElement("security"):
        returned = security.getElementAsString("security")
        if " ".join(returned.split()).casefold() != " ".join(cfg["ticker"].split()).casefold():
            raise ValueError("Unexpected security in response: " + returned + "; requested " + cfg["ticker"])
    if not security.hasElement("fieldData"):
        raise ValueError("Bloomberg response is missing fieldData.")
    values = security.getElement("fieldData")
    rows = []
    for n in range(values.numValues()):
        item = values.getValueAsElement(n)
        observed = item.getElementAsDatetime("date")
        observed = observed.date() if isinstance(observed, datetime) else observed
        if not isinstance(observed, date) or not start_day <= observed <= end_day:
            raise ValueError("Observation date outside requested range: " + str(observed))
        value = None
        if item.hasElement(cfg["field"]) and not item.getElement(cfg["field"]).isNull():
            value = item.getElementAsFloat(cfg["field"])
            if not math.isfinite(value):
                raise ValueError("Nonfinite API value")
        rows.append(dict(series_id=key, security=cfg["ticker"], field=cfg["field"],
                         observation_date=observed.isoformat(), value=value,
                         unit=cfg["unit"], source="Bloomberg", retrieved_at=retrieved_at))
    return rows


def fetch_bloomberg(end_day, start_day):
    """Bloomberg history -> an immutable snapshot -> the latest snapshot pointer.

    CSV and manifest must finish before the latest pointer changes. Failed runs
    keep API logs and all completed series for debugging, without advancing it.
    The next script builds only from a committed snapshot. No fill or fallback.
    """
    mappings = configured_requests(start_day, end_day)
    try:
        import blpapi
    except ImportError as exc:
        raise RuntimeError("Install/use blpapi in your Bloomberg-enabled Python environment. Offline build does not need it.") from exc
    retrieved = datetime.now(timezone.utc)
    folder = ROOT / "data" / "raw" / retrieved.strftime("%Y%m%dT%H%M%S%fZ")
    folder.mkdir(parents=True, exist_ok=False)
    csv_path = folder / "observations.csv"
    manifest_path = folder / "manifest.json"
    manifest = dict(status="running", retrieved_at=retrieved.isoformat(), start=str(start_day),
                    end=str(end_day), series=mappings, errors={})
    save_json(manifest_path, manifest)
    rows = []
    session = None
    try:
        options = blpapi.SessionOptions()
        options.setServerHost(BLOOMBERG_HOST)
        options.setServerPort(BLOOMBERG_PORT)
        session = blpapi.Session(options)
        if not session.start() or not session.openService("//blp/refdata"):
            raise RuntimeError("Cannot open Bloomberg //blp/refdata; check Desktop API session, port and entitlements.")
        service = session.getService("//blp/refdata")
        for request_number, (key, cfg) in enumerate(mappings.items(), 1):
            print("FETCH", key, cfg["ticker"], cfg["field"])
            own_rows = []
            correlation = blpapi.CorrelationId(request_number)
            sent = finished = False
            try:
                request = service.createRequest("HistoricalDataRequest")
                request.getElement("securities").appendValue(cfg["ticker"])
                request.getElement("fields").appendValue(cfg["field"])
                request.set("startDate", start_day.strftime("%Y%m%d"))
                request.set("endDate", end_day.strftime("%Y%m%d"))
                request.set("periodicitySelection", "DAILY")
                request.set("nonTradingDayFillOption", "ACTIVE_DAYS_ONLY")
                request.set("nonTradingDayFillMethod", "NIL_VALUE")
                session.sendRequest(request, correlationId=correlation)
                sent = True
                deadline = time.monotonic() + REQUEST_TIMEOUT
                while not finished:
                    if time.monotonic() >= deadline:
                        raise RuntimeError("Request timed out")
                    event = session.nextEvent(1000)
                    event_type = event.eventType()
                    for message in event:
                        # Session failures may have no correlation ID. Stop the run,
                        # instead of timing out once for every remaining security.
                        if event_type == getattr(blpapi.Event, "SESSION_STATUS", None):
                            if str(message.messageType()) in ("SessionTerminated", "SessionStartupFailure", "SessionConnectionDown"):
                                raise ConnectionError("Bloomberg session lost: " + str(message))
                        if not any(c.value() == request_number for c in message.correlationIds()):
                            continue
                        with (folder / (key + "_response.txt")).open("a", encoding="utf-8") as log:
                            log.write(str(message) + "\n")
                        if event_type == blpapi.Event.REQUEST_STATUS:
                            if str(message.messageType()) == "RequestFailure":
                                raise RuntimeError(str(message))
                            continue
                        if event_type not in (blpapi.Event.PARTIAL_RESPONSE, blpapi.Event.RESPONSE):
                            continue
                        own_rows.extend(response_rows(message, key, cfg, retrieved.isoformat(), start_day, end_day))
                        if event_type == blpapi.Event.RESPONSE:
                            finished = True
                if not any(row["value"] is not None for row in own_rows):
                    raise ValueError("No numeric observations")
                if len({row["observation_date"] for row in own_rows}) != len(own_rows):
                    raise ValueError("Duplicate observation dates in API response")
                rows.extend(sorted(own_rows, key=lambda row: row["observation_date"]))
            except Exception as exc:
                manifest["errors"][key] = str(exc)
                print("  FAILED:", exc)
                if sent and not finished:
                    try:
                        session.cancel(correlation)
                    except Exception as cancel_error:
                        print("  Could not cancel request:", cancel_error)
                if isinstance(exc, ConnectionError):
                    raise
            # A per-series checkpoint is also atomic. Only fully validated series
            # enter the CSV; rejected partial responses remain in the API logs.
            save_csv(csv_path, csv_string(RAW_COLUMNS, rows))
            save_json(manifest_path, manifest)
        failures = [key for key, cfg in mappings.items() if cfg["required"] and key in manifest["errors"]]
        if failures:
            raise RuntimeError("Required series failed: " + ", ".join(failures) + ". See " + str(folder))
        if not rows:
            raise RuntimeError("No observations downloaded")
        manifest["status"] = "partial" if manifest["errors"] else "complete"
        manifest["sha256"] = hashlib.sha256(csv_path.read_bytes()).hexdigest()
        save_json(manifest_path, manifest)
        # This is the commit point. Do not modify snapshot files after publishing.
        save_json(ROOT / "data" / "latest_bloomberg.json", {"raw_csv": csv_path.relative_to(ROOT).as_posix()})
    except BaseException as exc:
        manifest["status"] = "failed"
        manifest["run_error"] = str(exc) or type(exc).__name__
        try:
            save_csv(csv_path, csv_string(RAW_COLUMNS, rows))
            save_json(manifest_path, manifest)
        except Exception as save_error:
            # Disk failures must not hide the original API/configuration error.
            print("Could not finish failed-run diagnostics:", save_error)
        raise
    finally:
        if session is not None:
            try:
                session.stop()
            except Exception as stop_error:
                print("Bloomberg session cleanup:", stop_error)
    print("RAW SAVED:", csv_path)
    return csv_path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--asof", type=date.fromisoformat, default=date.today(), help="Data cutoff YYYY-MM-DD")
    parser.add_argument("--start", type=date.fromisoformat, help="History start; default 12 years before cutoff")
    args = parser.parse_args()
    start = args.start or years_before(args.asof, HISTORY_YEARS + 2)
    if start > args.asof:
        parser.error("--start must not be after --asof")
    fetch_bloomberg(args.asof, start)
    print("NEXT: python build_dashboard.py")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, RuntimeError, OSError) as error:
        raise SystemExit("ERROR: " + str(error))
