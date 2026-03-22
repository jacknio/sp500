from __future__ import annotations

import math
import statistics
import threading
import zipfile
from bisect import bisect_left, bisect_right
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

import requests


EXCEL_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": EXCEL_NS, "r": REL_NS}

YAHOO_CHART_URL = "https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
INTERVALS = ("daily", "weekly", "monthly")
CHART_BAR_LIMITS = {
    "daily": 1000,
    "weekly": 650,
    "monthly": 523,
}
TOP_EVENT_LIMIT = 5
KEY_LEVEL_LOOKBACKS = {
    "daily": 20,
    "weekly": 13,
    "monthly": 12,
}


@dataclass
class Bar:
    date: date
    open: float
    high: float
    low: float
    close: float
    source: str

    def as_dict(self) -> Dict[str, object]:
        return {
            "date": self.date.isoformat(),
            "open": round(self.open, 4),
            "high": round(self.high, 4),
            "low": round(self.low, 4),
            "close": round(self.close, 4),
        }


def _safe_float(value: Optional[str]) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _percentile(sorted_values: Sequence[float], value: float) -> float:
    if not sorted_values:
        return 0.0
    left = bisect_left(sorted_values, value)
    right = bisect_right(sorted_values, value)
    return (left + 0.5 * (right - left)) / len(sorted_values)


def _optional_percentile(sorted_values: Sequence[float], value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return _percentile(sorted_values, value)


def _histogram(values: Sequence[float], bins: int = 12) -> List[Dict[str, float]]:
    if not values:
        return []
    lo = min(values)
    hi = max(values)
    if math.isclose(lo, hi):
        lo -= 0.5
        hi += 0.5
    width = (hi - lo) / bins
    counts = [0 for _ in range(bins)]
    for value in values:
        idx = min(int((value - lo) / width), bins - 1)
        counts[idx] += 1
    output = []
    for idx, count in enumerate(counts):
        start = lo + idx * width
        end = start + width
        output.append(
            {
                "start": round(start, 4),
                "end": round(end, 4),
                "count": count,
            }
        )
    return output


def _mean(values: Sequence[float]) -> Optional[float]:
    if not values:
        return None
    return sum(values) / len(values)


def _median(values: Sequence[float]) -> Optional[float]:
    if not values:
        return None
    return statistics.median(values)


def _quantile(values: Sequence[float], q: float) -> Optional[float]:
    if not values:
        return None
    ordered = sorted(values)
    if len(ordered) == 1:
        return ordered[0]
    idx = (len(ordered) - 1) * q
    low = math.floor(idx)
    high = math.ceil(idx)
    if low == high:
        return ordered[int(idx)]
    weight = idx - low
    return ordered[low] * (1 - weight) + ordered[high] * weight


def _metric_values(rows: Sequence[Dict[str, object]], key: str) -> List[float]:
    return [float(row[key]) for row in rows if row.get(key) is not None]


def _fmt_number(value: Optional[float], decimals: int = 2) -> str:
    if value is None:
        return "N/A"
    return f"{float(value):,.{decimals}f}"


def _fmt_pct(value: Optional[float], decimals: int = 2) -> str:
    if value is None:
        return "N/A"
    return f"{float(value):+.{decimals}f}%"


def _fmt_ratio(value: Optional[float]) -> str:
    if value is None:
        return "N/A"
    return f"{float(value):.2f}x"


def _fmt_volume(value: Optional[float], language: str) -> str:
    if value is None:
        return "N/A" if language == "en" else "暂无"
    value = float(value)
    if language == "zh":
        return f"{value / 1e8:.1f}亿"
    if abs(value) >= 1e9:
        return f"{value / 1e9:.2f}B"
    if abs(value) >= 1e6:
        return f"{value / 1e6:.1f}M"
    return f"{value:,.0f}"


def _stats(values: Sequence[float]) -> Dict[str, Optional[float]]:
    if not values:
        return {
            "sample_size": 0,
            "up_probability": None,
            "mean_return": None,
            "median_return": None,
            "p10": None,
            "p25": None,
            "p75": None,
            "p90": None,
        }
    return {
        "sample_size": len(values),
        "up_probability": sum(1 for item in values if item > 0) / len(values),
        "mean_return": _mean(values),
        "median_return": _median(values),
        "p10": _quantile(values, 0.10),
        "p25": _quantile(values, 0.25),
        "p75": _quantile(values, 0.75),
        "p90": _quantile(values, 0.90),
    }


def load_clean_sheet_bars(path: Path) -> List[Bar]:
    workbook = zipfile.ZipFile(path)
    shared_strings = []
    if "xl/sharedStrings.xml" in workbook.namelist():
        root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
        for si in root:
            shared_strings.append(
                "".join(
                    text.text or ""
                    for text in si.iter("{%s}t" % EXCEL_NS)
                )
            )

    rels = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

    wb = ET.fromstring(workbook.read("xl/workbook.xml"))
    target = None
    for sheet in wb.find("a:sheets", NS):
        if sheet.attrib["name"] == "Clean":
            rid = sheet.attrib["{%s}id" % REL_NS]
            target = "xl/" + relmap[rid]
            break
    if not target:
        raise ValueError("Could not locate Clean sheet in sp500.xlsx")

    ws = ET.fromstring(workbook.read(target))
    rows = ws.find("a:sheetData", NS).findall("a:row", NS)
    bars: List[Bar] = []
    for row in rows[1:]:
        values: Dict[str, Optional[str]] = {}
        for cell in row.findall("a:c", NS):
            ref = cell.attrib.get("r", "")
            column = "".join(ch for ch in ref if ch.isalpha())
            value_node = cell.find("a:v", NS)
            if value_node is None:
                value = None
            elif cell.attrib.get("t") == "s":
                value = shared_strings[int(value_node.text)]
            else:
                value = value_node.text
            values[column] = value

        serial = _safe_float(values.get("A"))
        close = _safe_float(values.get("E"))
        if serial is None or close is None:
            continue
        dt = (datetime(1899, 12, 30) + timedelta(days=int(serial))).date()
        open_value = _safe_float(values.get("B")) or close
        high = _safe_float(values.get("C")) or max(open_value, close)
        low = _safe_float(values.get("D")) or min(open_value, close)
        bars.append(
            Bar(
                date=dt,
                open=open_value,
                high=high,
                low=low,
                close=close,
                source="local",
            )
        )
    return bars


def fetch_yahoo_history(symbol: str, start_date: date, end_date: Optional[date] = None) -> List[Dict[str, object]]:
    if end_date is None:
        end_date = datetime.now().date()
    if start_date > end_date:
        return []

    period1 = int(datetime.combine(start_date, datetime.min.time()).timestamp())
    period2 = int(datetime.combine(end_date + timedelta(days=1), datetime.min.time()).timestamp())
    response = requests.get(
        YAHOO_CHART_URL.format(symbol=requests.utils.quote(symbol, safe="")),
        params={
            "interval": "1d",
            "period1": period1,
            "period2": period2,
            "includePrePost": "false",
            "events": "div,splits",
        },
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=20,
    )
    response.raise_for_status()
    payload = response.json()
    result = (payload.get("chart") or {}).get("result") or []
    if not result:
        return []
    result = result[0]
    timestamps = result.get("timestamp") or []
    quote = ((result.get("indicators") or {}).get("quote") or [{}])[0]
    opens = quote.get("open") or []
    highs = quote.get("high") or []
    lows = quote.get("low") or []
    closes = quote.get("close") or []
    volumes = quote.get("volume") or []
    rows: List[Dict[str, object]] = []
    for idx, ts in enumerate(timestamps):
        open_value = opens[idx] if idx < len(opens) else None
        high = highs[idx] if idx < len(highs) else None
        low = lows[idx] if idx < len(lows) else None
        close = closes[idx] if idx < len(closes) else None
        if None in (open_value, high, low, close):
            continue
        dt = datetime.utcfromtimestamp(ts).date()
        volume = volumes[idx] if idx < len(volumes) else None
        rows.append(
            {
                "date": dt,
                "open": float(open_value),
                "high": float(high),
                "low": float(low),
                "close": float(close),
                "volume": int(volume) if volume is not None else None,
                "source": "yahoo",
            }
        )
    return rows


def fetch_yahoo_incremental(start_date: date, end_date: Optional[date] = None) -> List[Bar]:
    rows = fetch_yahoo_history("^GSPC", start_date, end_date)
    return [
        Bar(
            date=row["date"],
            open=row["open"],
            high=row["high"],
            low=row["low"],
            close=row["close"],
            source=row["source"],
        )
        for row in rows
    ]


def merge_bars(local_bars: Sequence[Bar], yahoo_bars: Sequence[Bar]) -> List[Bar]:
    merged: Dict[date, Bar] = {bar.date: bar for bar in local_bars}
    for bar in yahoo_bars:
        merged[bar.date] = bar
    return [merged[key] for key in sorted(merged)]


def group_bars_by_interval(bars: Sequence[Bar], interval: str) -> List[List[Bar]]:
    if interval == "daily":
        return [[bar] for bar in bars]
    grouped: Dict[Tuple[int, int], List[Bar]] = {}
    if interval == "weekly":
        for bar in bars:
            iso = bar.date.isocalendar()
            grouped.setdefault((iso[0], iso[1]), []).append(bar)
    elif interval == "monthly":
        for bar in bars:
            grouped.setdefault((bar.date.year, bar.date.month), []).append(bar)
    else:
        raise ValueError("Unsupported interval")

    return [sorted(grouped[key], key=lambda item: item.date) for key in sorted(grouped)]


def resample_bars(bars: Sequence[Bar], interval: str) -> List[Bar]:
    groups = group_bars_by_interval(bars, interval)

    output: List[Bar] = []
    for items in groups:
        output.append(
            Bar(
                date=items[-1].date,
                open=items[0].open,
                high=max(item.high for item in items),
                low=min(item.low for item in items),
                close=items[-1].close,
                source=items[-1].source,
            )
        )
    return output


def _interval_labels(interval: str) -> List[Tuple[str, int]]:
    if interval == "daily":
        return [("1d", 1), ("2d", 2), ("1w", 5)]
    return [("1bar", 1), ("2bar", 2), ("5bar", 5)]


def _format_event_copy(interval: str, event_type: str, row: Dict[str, object]) -> Tuple[str, str, str, str]:
    unit_zh = {"daily": "天", "weekly": "周", "monthly": "个月"}[interval]
    unit_en = {"daily": "day", "weekly": "week", "monthly": "month"}[interval]
    unit_en_plural = {"daily": "days", "weekly": "weeks", "monthly": "months"}[interval]
    streak = int(row.get("streak_value", 0))
    key_window = int(row.get("key_window", KEY_LEVEL_LOOKBACKS[interval]))
    if event_type == "down_streak":
        return (
            "连续下跌",
            "Down Streak",
            "最近 {0}{1}连续下跌".format(streak, unit_zh),
            "{0} consecutive down {1}".format(streak, unit_en_plural),
        )
    if event_type == "bearish_streak":
        return (
            "连续阴线",
            "Bearish Candle Streak",
            "最近 {0}{1}连续阴线".format(streak, unit_zh),
            "{0} consecutive bearish {1}".format(streak, unit_en_plural),
        )
    if event_type == "extreme_drop":
        percentile = int(round((1 - float(row["return_percentile"])) * 100))
        return (
            "极端下跌",
            "Extreme Selloff",
            "当前{0}跌幅处于历史最弱约 {1}%".format(unit_zh, percentile),
            "Current {0} drop ranks among the weakest {1}% in history".format(unit_en, percentile),
        )
    if event_type == "extreme_gain":
        percentile = int(round(float(row["return_percentile"]) * 100))
        return (
            "极端上涨",
            "Extreme Surge",
            "当前{0}涨幅处于历史最强约 {1}%".format(unit_zh, percentile),
            "Current {0} gain ranks among the strongest {1}% in history".format(unit_en, percentile),
        )
    if event_type == "outsized_range":
        percentile = int(round(float(row["range_percentile"]) * 100))
        return (
            "极端波动",
            "Oversized Range",
            "当前{0}振幅处于历史约 {1}% 分位".format(unit_zh, percentile),
            "Current {0} range sits around the {1}th percentile of history".format(unit_en, percentile),
        )
    if event_type == "gap_down":
        return (
            "向下跳空",
            "Gap Down",
            "当前{0}开盘向下跳空并跌破前一根区间".format(unit_zh),
            "Current {0} opened below the prior bar range".format(unit_en),
        )
    if event_type == "gap_up":
        return (
            "向上跳空",
            "Gap Up",
            "当前{0}开盘向上跳空并越过前一根区间".format(unit_zh),
            "Current {0} opened above the prior bar range".format(unit_en),
        )
    if event_type == "upper_shadow":
        return (
            "长上影",
            "Long Upper Shadow",
            "当前{0}出现历史级别的长上影".format(unit_zh),
            "Current {0} printed a historically large upper shadow".format(unit_en),
        )
    if event_type == "lower_shadow":
        return (
            "长下影",
            "Long Lower Shadow",
            "当前{0}出现历史级别的长下影".format(unit_zh),
            "Current {0} printed a historically large lower shadow".format(unit_en),
        )
    if event_type == "streak_vol_combo":
        return (
            "弱势放大",
            "Weakness With Range Expansion",
            "连续弱势并伴随异常波动放大",
            "Persistent weakness combined with abnormal range expansion",
        )
    if event_type == "capitulation_selloff":
        return (
            "恐慌性抛售",
            "Capitulation Selloff",
            "当前{0}出现极端跌幅并伴随异常波动放大".format(unit_zh),
            "Current {0} shows an extreme decline with range expansion".format(unit_en),
        )
    if event_type == "breakdown_low":
        return (
            "关键位跌破",
            "Key Level Breakdown",
            "当前{0}收盘跌破过去 {1}{0}收盘关键低点".format(unit_zh, key_window),
            "Current {0} close broke below the prior {1} {2} closing floor".format(unit_en, key_window, unit_en_plural),
        )
    if event_type == "failed_breakout":
        return (
            "突破失败",
            "Failed Breakout",
            "当前{0}上冲过去 {1}{0}关键高点后回落".format(unit_zh, key_window),
            "Current {0} probed above the prior {1} {2} high and failed".format(unit_en, key_window, unit_en_plural),
        )
    if event_type == "panic_reversal":
        return (
            "恐慌后反抽",
            "Panic Reversal",
            "当前{0}下探后明显收回，像一次 panic reversal".format(unit_zh),
            "Current {0} flushed lower and recovered like a panic reversal".format(unit_en),
        )
    if event_type == "outside_reversal_bear":
        return (
            "看跌外包反转",
            "Bearish Outside Reversal",
            "当前{0}形成外包结构并以弱势收盘".format(unit_zh),
            "Current {0} formed an outside bar and closed weak".format(unit_en),
        )
    if event_type == "outside_reversal_bull":
        return (
            "看涨外包反转",
            "Bullish Outside Reversal",
            "当前{0}形成外包结构并以强势收盘".format(unit_zh),
            "Current {0} formed an outside bar and closed strong".format(unit_en),
        )
    if event_type == "volume_climax":
        return (
            "放量踩踏",
            "Volume Climax",
            "当前{0}跌势伴随异常放量，接近一次 volume climax".format(unit_zh),
            "Current {0} decline is paired with unusual volume expansion".format(unit_en),
        )
    if event_type == "high_volume_breakdown":
        return (
            "放量破位",
            "High-Volume Breakdown",
            "当前{0}放量跌破过去 {1}{0}收盘关键低点".format(unit_zh, key_window),
            "Current {0} broke below the prior {1} {2} closing floor on heavy volume".format(unit_en, key_window, unit_en_plural),
        )
    if event_type == "vix_panic":
        return (
            "VIX 恐慌抬升",
            "VIX Panic",
            "当前{0}对应的 VIX 处于历史高波动区并继续抬升".format(unit_zh),
            "Current {0} lines up with a historically elevated and rising VIX".format(unit_en),
        )
    if event_type == "high_vix_breakdown":
        return (
            "高波动破位",
            "High-VIX Breakdown",
            "高 VIX 环境下，当前{0}跌破过去 {1}{0}收盘关键低点".format(unit_zh, key_window),
            "Current {0} broke below the prior {1} {2} closing floor in a high-VIX regime".format(unit_en, key_window, unit_en_plural),
        )
    return (
        "未知事件",
        "Unknown Event",
        "当前触发了一个罕见事件",
        "A rare event is active on the current bar",
    )


class SPXAnalyzer:
    def __init__(self, workbook_path: Path):
        self.workbook_path = workbook_path
        self.local_bars = load_clean_sheet_bars(workbook_path)
        self.local_latest_date = self.local_bars[-1].date
        self.yahoo_bars = fetch_yahoo_incremental(self.local_latest_date + timedelta(days=1))
        self.merged_bars = merge_bars(self.local_bars, self.yahoo_bars)
        self.latest_merged_date = self.merged_bars[-1].date
        self.gspc_history = fetch_yahoo_history("^GSPC", self.local_bars[0].date, self.latest_merged_date)
        self.gspc_volume_by_date = {
            row["date"]: row["volume"]
            for row in self.gspc_history
            if row.get("volume") is not None
        }
        self.vix_history = fetch_yahoo_history("^VIX", self.local_bars[0].date, self.latest_merged_date)
        self.vix_by_date = {row["date"]: row for row in self.vix_history}
        self._interval_cache: Dict[str, Dict[str, object]] = {}
        self._cache_lock = threading.Lock()

    def get_interval_payload(self, interval: str) -> Dict[str, object]:
        if interval not in INTERVALS:
            raise ValueError("Unsupported interval")
        if interval not in self._interval_cache:
            with self._cache_lock:
                if interval not in self._interval_cache:
                    self._interval_cache[interval] = self._build_interval_payload(interval)
        return self._interval_cache[interval]

    def _chart_bars(self, bars: Sequence[Dict[str, object]], interval: str) -> List[Dict[str, object]]:
        limit = CHART_BAR_LIMITS.get(interval, len(bars))
        if len(bars) <= limit:
            return list(bars)
        return list(bars[-limit:])

    def _build_interval_payload(self, interval: str) -> Dict[str, object]:
        resampled = resample_bars(self.merged_bars, interval)
        context_rows = self._build_context_rows(interval)
        rows = self._analyze_rows(resampled, context_rows, interval)
        bars_payload = []
        for bar, row in zip(resampled, rows):
            item = bar.as_dict()
            item["volume"] = int(round(float(row["volume"]))) if row.get("volume") is not None else None
            bars_payload.append(item)
        candidates = self._generate_candidates(rows, interval)
        triggered = [candidate for candidate in candidates if candidate["current_triggered"]]
        triggered.sort(key=lambda item: item["rarity_score"], reverse=True)
        top_events: List[Dict[str, object]] = []
        seen_families = set()
        for candidate in triggered:
            family = candidate["family"]
            if family in seen_families:
                continue
            top_events.append(candidate)
            seen_families.add(family)
            if len(top_events) == TOP_EVENT_LIMIT:
                break

        summary_zh = top_events[0]["plain_language_description"] if top_events else "当前未识别出显著罕见事件"
        summary_en = top_events[0]["plain_language_description_en"] if top_events else "No standout rare event is active right now"
        return {
            "interval": interval,
            "bars": bars_payload,
            "rows": rows,
            "events": candidates,
            "top_events": top_events,
            "summary": summary_zh,
            "summary_en": summary_en,
        }

    def _build_context_rows(self, interval: str) -> List[Dict[str, Optional[float]]]:
        groups = group_bars_by_interval(self.merged_bars, interval)
        context_rows: List[Dict[str, Optional[float]]] = []
        for group in groups:
            dates = [bar.date for bar in group]
            volumes = [
                self.gspc_volume_by_date[bar_date]
                for bar_date in dates
                if bar_date in self.gspc_volume_by_date
            ]
            vix_rows = [
                self.vix_by_date[bar_date]
                for bar_date in dates
                if bar_date in self.vix_by_date
            ]
            context_rows.append(
                {
                    "volume": float(sum(volumes)) if volumes else None,
                    "vix_open": float(vix_rows[0]["open"]) if vix_rows else None,
                    "vix_high": max(float(row["high"]) for row in vix_rows) if vix_rows else None,
                    "vix_low": min(float(row["low"]) for row in vix_rows) if vix_rows else None,
                    "vix_close": float(vix_rows[-1]["close"]) if vix_rows else None,
                }
            )
        return context_rows

    def _analyze_rows(
        self,
        bars: Sequence[Bar],
        context_rows: Sequence[Dict[str, Optional[float]]],
        interval: str,
    ) -> List[Dict[str, object]]:
        rows: List[Dict[str, object]] = []
        prev_close = None
        prev_high = None
        prev_low = None
        prior_closes: List[float] = []
        prev_vix_close = None
        key_window = KEY_LEVEL_LOOKBACKS[interval]

        down_streak = 0
        bearish_streak = 0
        for idx, bar in enumerate(bars):
            context = context_rows[idx] if idx < len(context_rows) else {}
            close_return = ((bar.close / prev_close) - 1) * 100 if prev_close else 0.0
            is_down_close = prev_close is not None and bar.close < prev_close
            is_bearish = bar.close < bar.open
            down_streak = down_streak + 1 if is_down_close else 0
            bearish_streak = bearish_streak + 1 if is_bearish else 0

            range_pct = ((bar.high - bar.low) / bar.close) * 100 if bar.close else 0.0
            body_high = max(bar.open, bar.close)
            body_low = min(bar.open, bar.close)
            full_range = max(bar.high - bar.low, 1e-9)
            upper_shadow_pct = ((bar.high - body_high) / bar.close) * 100 if bar.close else 0.0
            lower_shadow_pct = ((body_low - bar.low) / bar.close) * 100 if bar.close else 0.0
            gap_pct = ((bar.open / prev_close) - 1) * 100 if prev_close else 0.0
            gap_down = prev_low is not None and bar.open < prev_low
            gap_up = prev_high is not None and bar.open > prev_high
            close_location = (bar.close - bar.low) / full_range
            outside_bar = prev_high is not None and prev_low is not None and bar.high > prev_high and bar.low < prev_low

            trailing_window = prior_closes[-key_window:]
            prior_window_low = min(trailing_window) if len(trailing_window) == key_window else None
            prior_window_high = max(trailing_window) if len(trailing_window) == key_window else None
            breakdown_low = prior_window_low is not None and bar.close <= prior_window_low
            breakout_high = prior_window_high is not None and bar.close >= prior_window_high

            trailing_returns = [float(row["close_return"]) for row in rows[max(0, len(rows) - 19):]] + [close_return]
            rolling_vol = statistics.pstdev(trailing_returns) if len(trailing_returns) > 1 else 0.0
            trailing_volume = [
                float(row["volume"])
                for row in rows[max(0, len(rows) - 19):]
                if row.get("volume") is not None
            ]
            volume = context.get("volume")
            avg_volume = (sum(trailing_volume) / len(trailing_volume)) if trailing_volume else None
            volume_ratio = (float(volume) / avg_volume) if volume is not None and avg_volume else None
            vix_close = context.get("vix_close")
            vix_return = ((float(vix_close) / prev_vix_close) - 1) * 100 if vix_close is not None and prev_vix_close else None

            rows.append(
                {
                    "index": idx,
                    "date": bar.date.isoformat(),
                    "open": bar.open,
                    "high": bar.high,
                    "low": bar.low,
                    "close": bar.close,
                    "close_return": close_return,
                    "range_pct": range_pct,
                    "upper_shadow_pct": upper_shadow_pct,
                    "lower_shadow_pct": lower_shadow_pct,
                    "gap_pct": gap_pct,
                    "gap_down": gap_down,
                    "gap_up": gap_up,
                    "close_location": close_location,
                    "outside_bar": outside_bar,
                    "breakdown_low": breakdown_low,
                    "breakout_high": breakout_high,
                    "prior_close_floor": prior_window_low,
                    "prior_close_ceiling": prior_window_high,
                    "down_streak": down_streak,
                    "bearish_streak": bearish_streak,
                    "rolling_vol": rolling_vol,
                    "key_window": key_window,
                    "volume": float(volume) if volume is not None else None,
                    "volume_ratio": volume_ratio,
                    "vix_close": float(vix_close) if vix_close is not None else None,
                    "vix_return": vix_return,
                }
            )

            prev_close = bar.close
            prev_high = bar.high
            prev_low = bar.low
            prior_closes.append(bar.close)
            if vix_close is not None:
                prev_vix_close = float(vix_close)

        sorted_returns = sorted(float(row["close_return"]) for row in rows)
        sorted_ranges = sorted(float(row["range_pct"]) for row in rows)
        sorted_upper = sorted(float(row["upper_shadow_pct"]) for row in rows)
        sorted_lower = sorted(float(row["lower_shadow_pct"]) for row in rows)
        sorted_volumes = sorted(float(row["volume"]) for row in rows if row["volume"] is not None)
        sorted_volume_ratios = sorted(float(row["volume_ratio"]) for row in rows if row["volume_ratio"] is not None)
        sorted_vix_levels = sorted(float(row["vix_close"]) for row in rows if row["vix_close"] is not None)
        sorted_vix_returns = sorted(float(row["vix_return"]) for row in rows if row["vix_return"] is not None)

        for row in rows:
            row["return_percentile"] = _percentile(sorted_returns, float(row["close_return"]))
            row["range_percentile"] = _percentile(sorted_ranges, float(row["range_pct"]))
            row["upper_shadow_percentile"] = _percentile(sorted_upper, float(row["upper_shadow_pct"]))
            row["lower_shadow_percentile"] = _percentile(sorted_lower, float(row["lower_shadow_pct"]))
            row["volume_percentile"] = _optional_percentile(sorted_volumes, row["volume"])
            row["volume_ratio_percentile"] = _optional_percentile(sorted_volume_ratios, row["volume_ratio"])
            row["vix_level_percentile"] = _optional_percentile(sorted_vix_levels, row["vix_close"])
            row["vix_return_percentile"] = _optional_percentile(sorted_vix_returns, row["vix_return"])
        return rows

    def _generate_candidates(self, rows: Sequence[Dict[str, object]], interval: str) -> List[Dict[str, object]]:
        latest = rows[-1]
        definitions = [
            {
                "type": "down_streak",
                "family": "streak",
                "predicate": lambda row: int(row["down_streak"]) >= 3,
                "severity": lambda row: int(row["down_streak"]),
            },
            {
                "type": "bearish_streak",
                "family": "candles",
                "predicate": lambda row: int(row["bearish_streak"]) >= 3,
                "severity": lambda row: int(row["bearish_streak"]),
            },
            {
                "type": "extreme_drop",
                "family": "return",
                "predicate": lambda row: float(row["return_percentile"]) <= 0.02,
                "severity": lambda row: 1 - float(row["return_percentile"]),
            },
            {
                "type": "extreme_gain",
                "family": "return",
                "predicate": lambda row: float(row["return_percentile"]) >= 0.98,
                "severity": lambda row: float(row["return_percentile"]),
            },
            {
                "type": "outsized_range",
                "family": "range",
                "predicate": lambda row: float(row["range_percentile"]) >= 0.98,
                "severity": lambda row: float(row["range_percentile"]),
            },
            {
                "type": "gap_down",
                "family": "gap",
                "predicate": lambda row: bool(row["gap_down"]),
                "severity": lambda row: abs(float(row["gap_pct"])),
            },
            {
                "type": "gap_up",
                "family": "gap",
                "predicate": lambda row: bool(row["gap_up"]),
                "severity": lambda row: abs(float(row["gap_pct"])),
            },
            {
                "type": "upper_shadow",
                "family": "shadow",
                "predicate": lambda row: float(row["upper_shadow_percentile"]) >= 0.985,
                "severity": lambda row: float(row["upper_shadow_percentile"]),
            },
            {
                "type": "lower_shadow",
                "family": "shadow",
                "predicate": lambda row: float(row["lower_shadow_percentile"]) >= 0.985,
                "severity": lambda row: float(row["lower_shadow_percentile"]),
            },
            {
                "type": "streak_vol_combo",
                "family": "combo",
                "predicate": lambda row: int(row["down_streak"]) >= 3 and float(row["range_percentile"]) >= 0.90,
                "severity": lambda row: int(row["down_streak"]) + float(row["range_percentile"]),
            },
            {
                "type": "capitulation_selloff",
                "family": "capitulation",
                "predicate": lambda row: (
                    float(row["return_percentile"]) <= 0.05
                    and float(row["range_percentile"]) >= 0.95
                    and int(row["down_streak"]) >= 2
                ),
                "severity": lambda row: (
                    (1 - float(row["return_percentile"])) * 2
                    + float(row["range_percentile"])
                    + int(row["down_streak"])
                ),
            },
            {
                "type": "breakdown_low",
                "family": "breakdown",
                "predicate": lambda row: (
                    bool(row["breakdown_low"])
                    and float(row["close_location"]) <= 0.30
                    and float(row["close_return"]) < 0
                ),
                "severity": lambda row: abs(float(row["close_return"])) + (1 - float(row["close_location"])) * 2,
            },
            {
                "type": "failed_breakout",
                "family": "breakout",
                "predicate": lambda row: (
                    bool(row["breakout_high"])
                    and float(row["upper_shadow_percentile"]) >= 0.96
                    and float(row["close_location"]) <= 0.45
                ),
                "severity": lambda row: float(row["upper_shadow_percentile"]) + (1 - float(row["close_location"])) * 2,
            },
            {
                "type": "panic_reversal",
                "family": "reversal",
                "predicate": lambda row: (
                    float(row["lower_shadow_percentile"]) >= 0.985
                    and int(row["down_streak"]) >= 2
                    and float(row["close_location"]) >= 0.60
                ),
                "severity": lambda row: (
                    float(row["lower_shadow_percentile"]) + float(row["close_location"]) + int(row["down_streak"])
                ),
            },
            {
                "type": "outside_reversal_bear",
                "family": "outside",
                "predicate": lambda row: (
                    bool(row["outside_bar"])
                    and float(row["close_location"]) <= 0.25
                    and float(row["close_return"]) < 0
                ),
                "severity": lambda row: abs(float(row["close_return"])) + (1 - float(row["close_location"])) * 2,
            },
            {
                "type": "outside_reversal_bull",
                "family": "outside",
                "predicate": lambda row: (
                    bool(row["outside_bar"])
                    and float(row["close_location"]) >= 0.75
                    and float(row["close_return"]) > 0
                ),
                "severity": lambda row: abs(float(row["close_return"])) + float(row["close_location"]) * 2,
            },
            {
                "type": "volume_climax",
                "family": "volume",
                "predicate": lambda row: (
                    row.get("volume_percentile") is not None
                    and float(row["volume_percentile"]) >= 0.95
                    and float(row["return_percentile"]) <= 0.15
                    and float(row["close_return"]) < 0
                ),
                "severity": lambda row: (
                    float(row["volume_percentile"]) + (1 - float(row["return_percentile"])) + abs(float(row["close_return"])) / 4
                ),
            },
            {
                "type": "high_volume_breakdown",
                "family": "volume_breakdown",
                "predicate": lambda row: (
                    bool(row["breakdown_low"])
                    and row.get("volume_ratio_percentile") is not None
                    and float(row["volume_ratio_percentile"]) >= 0.90
                ),
                "severity": lambda row: (
                    float(row["volume_ratio_percentile"]) + abs(float(row["close_return"])) / 3
                ),
            },
            {
                "type": "vix_panic",
                "family": "vix",
                "predicate": lambda row: (
                    row.get("vix_level_percentile") is not None
                    and row.get("vix_return_percentile") is not None
                    and float(row["vix_level_percentile"]) >= 0.85
                    and float(row["vix_return_percentile"]) >= 0.90
                ),
                "severity": lambda row: (
                    float(row["vix_level_percentile"]) + float(row["vix_return_percentile"])
                ),
            },
            {
                "type": "high_vix_breakdown",
                "family": "vix_breakdown",
                "predicate": lambda row: (
                    bool(row["breakdown_low"])
                    and row.get("vix_level_percentile") is not None
                    and float(row["vix_level_percentile"]) >= 0.85
                ),
                "severity": lambda row: (
                    float(row["vix_level_percentile"]) + abs(float(row["close_return"])) / 3
                ),
            },
        ]

        events = []
        for definition in definitions:
            matches = [row for row in rows if definition["predicate"](row)]
            event = self._build_event(definition, matches, latest, rows, interval)
            events.append(event)
        return events

    def _build_event(
        self,
        definition: Dict[str, object],
        matches: Sequence[Dict[str, object]],
        latest: Dict[str, object],
        rows: Sequence[Dict[str, object]],
        interval: str,
    ) -> Dict[str, object]:
        event_type = str(definition["type"])
        current_triggered = bool(definition["predicate"](latest))
        severity_value = float(definition["severity"](latest)) if current_triggered else 0.0
        streak_value = 0
        if event_type == "down_streak":
            streak_value = int(latest.get("down_streak", 0))
        elif event_type == "bearish_streak":
            streak_value = int(latest.get("bearish_streak", 0))

        methodology_note_zh, methodology_note_en = self._methodology_note(interval, event_type, latest)
        rule_lines_zh, rule_lines_en = self._rule_lines(interval, event_type, latest, rows)
        title_zh, title_en, description_zh, description_en = _format_event_copy(
            interval,
            event_type,
            {
                "streak_value": streak_value,
                "return_percentile": latest.get("return_percentile", 0),
                "range_percentile": latest.get("range_percentile", 0),
                "key_window": latest.get("key_window", KEY_LEVEL_LOOKBACKS[interval]),
            },
        )

        forward_windows = _interval_labels(interval)
        conditional_stats = {}
        unconditional_stats = {}
        distributions = {}
        examples = []
        recent_occurrences = []
        effect_score = 0.0
        eligible_count = 0
        match_indices = {int(row["index"]) for row in matches}

        for label, offset in forward_windows:
            conditional_returns = []
            unconditional_returns = []
            for row in rows:
                idx = int(row["index"])
                if idx + offset >= len(rows):
                    continue
                future_close = float(rows[idx + offset]["close"])
                current_close = float(row["close"])
                fwd_return = ((future_close / current_close) - 1) * 100
                unconditional_returns.append(fwd_return)
                if idx in match_indices:
                    conditional_returns.append(fwd_return)
            conditional_stats[label] = _stats(conditional_returns)
            unconditional_stats[label] = _stats(unconditional_returns)
            distributions[label] = {
                "conditional": _histogram(conditional_returns),
                "unconditional": _histogram(unconditional_returns),
            }
            if conditional_stats[label]["up_probability"] is not None and unconditional_stats[label]["up_probability"] is not None:
                effect_score += abs(
                    conditional_stats[label]["up_probability"] - unconditional_stats[label]["up_probability"]
                )
            if label == forward_windows[0][0]:
                eligible_count = len(conditional_returns)

        for row in matches[-5:]:
            examples.append(
                {
                    "date": row["date"],
                    "close_return": round(float(row["close_return"]), 3),
                    "range_pct": round(float(row["range_pct"]), 3),
                    "down_streak": int(row["down_streak"]),
                    "bearish_streak": int(row["bearish_streak"]),
                }
            )

        for row in reversed(matches[-8:]):
            occurrence = {
                "date": row["date"],
                "close_return": round(float(row["close_return"]), 3),
                "forward_returns": {},
            }
            row_idx = int(row["index"])
            for label, offset in forward_windows:
                if row_idx + offset >= len(rows):
                    occurrence["forward_returns"][label] = None
                    continue
                future_close = float(rows[row_idx + offset]["close"])
                current_close = float(row["close"])
                occurrence["forward_returns"][label] = round(((future_close / current_close) - 1) * 100, 3)
            recent_occurrences.append(occurrence)

        rarity = 1 - (len(matches) / len(rows) if rows else 0.0)
        rarity_score = rarity * 60 + min(effect_score * 100, 30) + min(severity_value * 2, 10)
        return {
            "event_id": event_type,
            "family": definition["family"],
            "title": title_zh,
            "title_en": title_en,
            "plain_language_description": description_zh,
            "plain_language_description_en": description_en,
            "methodology_note": methodology_note_zh,
            "methodology_note_en": methodology_note_en,
            "rule_lines": rule_lines_zh,
            "rule_lines_en": rule_lines_en,
            "current_triggered": current_triggered,
            "historical_count": len(matches),
            "sample_size": eligible_count,
            "rarity_score": round(rarity_score, 3),
            "forward_windows": [label for label, _ in forward_windows],
            "conditional_stats": conditional_stats,
            "unconditional_stats": unconditional_stats,
            "distribution_bins": distributions,
            "historical_examples": examples,
            "recent_occurrences": recent_occurrences,
            "current_context": {
                "key_window": int(latest.get("key_window", KEY_LEVEL_LOOKBACKS[interval])),
                "volume_percentile": latest.get("volume_percentile"),
                "vix_level_percentile": latest.get("vix_level_percentile"),
            },
            "latest_date": latest["date"],
        }

    def _methodology_note(self, interval: str, event_type: str, latest: Dict[str, object]) -> Tuple[str, str]:
        unit = {"daily": "day", "weekly": "week", "monthly": "month"}[interval]
        unit_zh = {"daily": "天", "weekly": "周", "monthly": "个月"}[interval]
        key_window = int(latest.get("key_window", KEY_LEVEL_LOOKBACKS[interval]))
        if event_type in {"breakdown_low", "failed_breakout", "high_volume_breakdown", "high_vix_breakdown"}:
            return (
                f"关键位定义 = 不包含当前 bar 的过去 {key_window}{unit_zh}收盘区间。",
                f"Key level = prior {key_window} {unit} closing range, excluding the current bar.",
            )
        if event_type in {"volume_climax"}:
            return (
                "量能事件 = 价格走弱并伴随指数成交量处于历史高分位。",
                "Volume event = price weakness combined with top-percentile index volume.",
            )
        if event_type in {"vix_panic"}:
            return (
                "VIX 事件 = VIX 水平本身偏高，并且本期继续明显上冲。",
                "VIX event = elevated VIX level plus a sharp VIX upswing versus history.",
            )
        return (
            "条件统计 = 用同一种模式回看历史，并与无条件基准比较。",
            "Conditional stats use the same pattern across history and compare it with the unconditional baseline.",
        )

    def _rule_lines(
        self,
        interval: str,
        event_type: str,
        latest: Dict[str, object],
        rows: Sequence[Dict[str, object]],
    ) -> Tuple[List[str], List[str]]:
        key_window = int(latest.get("key_window", KEY_LEVEL_LOOKBACKS[interval]))
        unit_zh = {"daily": "天", "weekly": "周", "monthly": "个月"}[interval]
        unit_en = {"daily": "day", "weekly": "week", "monthly": "month"}[interval]
        unit_en_plural = {"daily": "days", "weekly": "weeks", "monthly": "months"}[interval]

        return_values = _metric_values(rows, "close_return")
        range_values = _metric_values(rows, "range_pct")
        upper_shadow_values = _metric_values(rows, "upper_shadow_pct")
        lower_shadow_values = _metric_values(rows, "lower_shadow_pct")
        volume_values = _metric_values(rows, "volume")
        volume_ratio_values = _metric_values(rows, "volume_ratio")
        vix_level_values = _metric_values(rows, "vix_close")
        vix_return_values = _metric_values(rows, "vix_return")

        return_p15 = _quantile(return_values, 0.15)
        return_p02 = _quantile(return_values, 0.02)
        return_p98 = _quantile(return_values, 0.98)
        range_p98 = _quantile(range_values, 0.98)
        upper_shadow_p96 = _quantile(upper_shadow_values, 0.96)
        upper_shadow_p985 = _quantile(upper_shadow_values, 0.985)
        lower_shadow_p985 = _quantile(lower_shadow_values, 0.985)
        volume_p95 = _quantile(volume_values, 0.95)
        volume_ratio_p90 = _quantile(volume_ratio_values, 0.90)
        vix_level_p85 = _quantile(vix_level_values, 0.85)
        vix_return_p90 = _quantile(vix_return_values, 0.90)

        close_value = float(latest.get("close", 0))
        close_return = latest.get("close_return")
        close_location = latest.get("close_location")
        prior_floor = latest.get("prior_close_floor")
        prior_ceiling = latest.get("prior_close_ceiling")
        volume = latest.get("volume")
        volume_ratio = latest.get("volume_ratio")
        vix_close = latest.get("vix_close")
        vix_return = latest.get("vix_return")
        down_streak = int(latest.get("down_streak", 0))
        bearish_streak = int(latest.get("bearish_streak", 0))
        range_pct = latest.get("range_pct")
        upper_shadow_pct = latest.get("upper_shadow_pct")
        lower_shadow_pct = latest.get("lower_shadow_pct")

        if event_type == "down_streak":
            return (
                [
                    f"连续下跌计数 = {down_streak}，触发阈值 >= 3",
                    f"本根涨跌幅 = {_fmt_pct(close_return)}",
                ],
                [
                    f"Down streak count = {down_streak}, trigger threshold >= 3",
                    f"Current bar return = {_fmt_pct(close_return)}",
                ],
            )
        if event_type == "bearish_streak":
            return (
                [
                    f"连续阴线计数 = {bearish_streak}，触发阈值 >= 3",
                    f"当前收盘 {_fmt_number(close_value)} < 开盘 {_fmt_number(latest.get('open'))}",
                ],
                [
                    f"Bearish candle streak = {bearish_streak}, trigger threshold >= 3",
                    f"Current close {_fmt_number(close_value)} < open {_fmt_number(latest.get('open'))}",
                ],
            )
        if event_type == "volume_climax":
            return (
                [
                    f"成交量 {_fmt_volume(volume, 'zh')} >= 历史 95 分位阈值 {_fmt_volume(volume_p95, 'zh')}",
                    f"当根涨跌 {_fmt_pct(close_return)} <= 历史 15 分位阈值 {_fmt_pct(return_p15)}",
                    "当根必须为下跌 bar",
                ],
                [
                    f"Volume {_fmt_volume(volume, 'en')} >= 95th percentile threshold {_fmt_volume(volume_p95, 'en')}",
                    f"Bar return {_fmt_pct(close_return)} <= 15th percentile threshold {_fmt_pct(return_p15)}",
                    "Current bar must be negative",
                ],
            )
        if event_type == "vix_panic":
            return (
                [
                    f"VIX 当前值 {_fmt_number(vix_close)} >= 历史 85 分位阈值 {_fmt_number(vix_level_p85)}",
                    f"VIX 单期变动 {_fmt_pct(vix_return)} >= 历史 90 分位阈值 {_fmt_pct(vix_return_p90)}",
                ],
                [
                    f"Current VIX {_fmt_number(vix_close)} >= 85th percentile threshold {_fmt_number(vix_level_p85)}",
                    f"VIX change {_fmt_pct(vix_return)} >= 90th percentile threshold {_fmt_pct(vix_return_p90)}",
                ],
            )
        if event_type == "breakdown_low":
            return (
                [
                    f"关键位 = 不含当前 bar 的过去 {key_window}{unit_zh}收盘低点 {_fmt_number(prior_floor)}",
                    f"当前收盘 {_fmt_number(close_value)} <= 关键低点 {_fmt_number(prior_floor)}",
                    f"收盘位置 {close_location * 100:.1f}% <= 30.0%",
                ],
                [
                    f"Key level = prior {key_window} {unit_en_plural} closing floor {_fmt_number(prior_floor)}",
                    f"Current close {_fmt_number(close_value)} <= floor {_fmt_number(prior_floor)}",
                    f"Close location {close_location * 100:.1f}% <= 30.0%",
                ],
            )
        if event_type == "high_volume_breakdown":
            return (
                [
                    f"关键位 = 不含当前 bar 的过去 {key_window}{unit_zh}收盘低点 {_fmt_number(prior_floor)}",
                    f"当前收盘 {_fmt_number(close_value)} <= 关键低点 {_fmt_number(prior_floor)}",
                    f"量能倍数 {_fmt_ratio(volume_ratio)} >= 历史 90 分位阈值 {_fmt_ratio(volume_ratio_p90)}",
                ],
                [
                    f"Key level = prior {key_window} {unit_en_plural} closing floor {_fmt_number(prior_floor)}",
                    f"Current close {_fmt_number(close_value)} <= floor {_fmt_number(prior_floor)}",
                    f"Volume ratio {_fmt_ratio(volume_ratio)} >= 90th percentile threshold {_fmt_ratio(volume_ratio_p90)}",
                ],
            )
        if event_type == "high_vix_breakdown":
            return (
                [
                    f"关键位 = 不含当前 bar 的过去 {key_window}{unit_zh}收盘低点 {_fmt_number(prior_floor)}",
                    f"当前收盘 {_fmt_number(close_value)} <= 关键低点 {_fmt_number(prior_floor)}",
                    f"VIX 当前值 {_fmt_number(vix_close)} >= 历史 85 分位阈值 {_fmt_number(vix_level_p85)}",
                ],
                [
                    f"Key level = prior {key_window} {unit_en_plural} closing floor {_fmt_number(prior_floor)}",
                    f"Current close {_fmt_number(close_value)} <= floor {_fmt_number(prior_floor)}",
                    f"Current VIX {_fmt_number(vix_close)} >= 85th percentile threshold {_fmt_number(vix_level_p85)}",
                ],
            )
        if event_type == "capitulation_selloff":
            return (
                [
                    f"当根涨跌 {_fmt_pct(close_return)} <= 历史 5 分位附近阈值 {_fmt_pct(_quantile(return_values, 0.05))}",
                    f"振幅 {_fmt_pct(range_pct)} >= 历史 95 分位附近阈值 {_fmt_pct(_quantile(range_values, 0.95))}",
                    f"连续下跌计数 = {down_streak}，触发阈值 >= 2",
                ],
                [
                    f"Bar return {_fmt_pct(close_return)} <= ~5th percentile threshold {_fmt_pct(_quantile(return_values, 0.05))}",
                    f"Range {_fmt_pct(range_pct)} >= ~95th percentile threshold {_fmt_pct(_quantile(range_values, 0.95))}",
                    f"Down streak count = {down_streak}, trigger threshold >= 2",
                ],
            )
        if event_type == "failed_breakout":
            return (
                [
                    f"关键位 = 不含当前 bar 的过去 {key_window}{unit_zh}收盘高点 {_fmt_number(prior_ceiling)}",
                    f"当前收盘 {_fmt_number(close_value)} >= 关键高点 {_fmt_number(prior_ceiling)}",
                    f"上影 {_fmt_pct(upper_shadow_pct)} >= 历史 96 分位阈值 {_fmt_pct(upper_shadow_p96)}",
                    f"收盘位置 {close_location * 100:.1f}% <= 45.0%",
                ],
                [
                    f"Key level = prior {key_window} {unit_en_plural} closing ceiling {_fmt_number(prior_ceiling)}",
                    f"Current close {_fmt_number(close_value)} >= ceiling {_fmt_number(prior_ceiling)}",
                    f"Upper shadow {_fmt_pct(upper_shadow_pct)} >= 96th percentile threshold {_fmt_pct(upper_shadow_p96)}",
                    f"Close location {close_location * 100:.1f}% <= 45.0%",
                ],
            )
        if event_type == "panic_reversal":
            return (
                [
                    f"下影 {_fmt_pct(lower_shadow_pct)} >= 历史 98.5 分位阈值 {_fmt_pct(lower_shadow_p985)}",
                    f"连续下跌计数 = {down_streak}，触发阈值 >= 2",
                    f"收盘位置 {close_location * 100:.1f}% >= 60.0%",
                ],
                [
                    f"Lower shadow {_fmt_pct(lower_shadow_pct)} >= 98.5th percentile threshold {_fmt_pct(lower_shadow_p985)}",
                    f"Down streak count = {down_streak}, trigger threshold >= 2",
                    f"Close location {close_location * 100:.1f}% >= 60.0%",
                ],
            )
        if event_type == "outside_reversal_bear":
            return (
                [
                    "当前 bar 为 outside bar，区间同时高于前高且低于前低",
                    f"收盘位置 {close_location * 100:.1f}% <= 25.0%",
                    f"当根涨跌 {_fmt_pct(close_return)} < 0%",
                ],
                [
                    "Current bar is an outside bar: above prior high and below prior low",
                    f"Close location {close_location * 100:.1f}% <= 25.0%",
                    f"Bar return {_fmt_pct(close_return)} < 0%",
                ],
            )
        if event_type == "outside_reversal_bull":
            return (
                [
                    "当前 bar 为 outside bar，区间同时高于前高且低于前低",
                    f"收盘位置 {close_location * 100:.1f}% >= 75.0%",
                    f"当根涨跌 {_fmt_pct(close_return)} > 0%",
                ],
                [
                    "Current bar is an outside bar: above prior high and below prior low",
                    f"Close location {close_location * 100:.1f}% >= 75.0%",
                    f"Bar return {_fmt_pct(close_return)} > 0%",
                ],
            )
        if event_type == "extreme_drop":
            return (
                [
                    f"当根涨跌 {_fmt_pct(close_return)} <= 历史 2 分位阈值 {_fmt_pct(return_p02)}",
                ],
                [
                    f"Bar return {_fmt_pct(close_return)} <= 2nd percentile threshold {_fmt_pct(return_p02)}",
                ],
            )
        if event_type == "extreme_gain":
            return (
                [
                    f"当根涨跌 {_fmt_pct(close_return)} >= 历史 98 分位阈值 {_fmt_pct(return_p98)}",
                ],
                [
                    f"Bar return {_fmt_pct(close_return)} >= 98th percentile threshold {_fmt_pct(return_p98)}",
                ],
            )
        if event_type == "outsized_range":
            return (
                [
                    f"振幅 {_fmt_pct(range_pct)} >= 历史 98 分位阈值 {_fmt_pct(range_p98)}",
                ],
                [
                    f"Range {_fmt_pct(range_pct)} >= 98th percentile threshold {_fmt_pct(range_p98)}",
                ],
            )
        if event_type == "upper_shadow":
            return (
                [
                    f"上影 {_fmt_pct(upper_shadow_pct)} >= 历史 98.5 分位阈值 {_fmt_pct(upper_shadow_p985)}",
                ],
                [
                    f"Upper shadow {_fmt_pct(upper_shadow_pct)} >= 98.5th percentile threshold {_fmt_pct(upper_shadow_p985)}",
                ],
            )
        if event_type == "lower_shadow":
            return (
                [
                    f"下影 {_fmt_pct(lower_shadow_pct)} >= 历史 98.5 分位阈值 {_fmt_pct(lower_shadow_p985)}",
                ],
                [
                    f"Lower shadow {_fmt_pct(lower_shadow_pct)} >= 98.5th percentile threshold {_fmt_pct(lower_shadow_p985)}",
                ],
            )
        if event_type == "gap_down":
            return (
                [
                    "当前开盘价低于前一根最低价",
                    f"开盘跳空幅度 = {_fmt_pct(latest.get('gap_pct'))}",
                ],
                [
                    "Current open is below the prior bar low",
                    f"Gap size = {_fmt_pct(latest.get('gap_pct'))}",
                ],
            )
        if event_type == "gap_up":
            return (
                [
                    "当前开盘价高于前一根最高价",
                    f"开盘跳空幅度 = {_fmt_pct(latest.get('gap_pct'))}",
                ],
                [
                    "Current open is above the prior bar high",
                    f"Gap size = {_fmt_pct(latest.get('gap_pct'))}",
                ],
            )
        return (
            [f"当前 {unit_zh}满足该事件筛选条件"],
            [f"Current {unit_en} satisfies the event filter"],
        )

    def get_series(self, interval: str) -> Dict[str, object]:
        payload = self.get_interval_payload(interval)
        return {
            "interval": interval,
            "bars": self._chart_bars(payload["bars"], interval),
            "full_bar_count": len(payload["bars"]),
            "chart_bar_count": len(self._chart_bars(payload["bars"], interval)),
            "latest_date": payload["bars"][-1]["date"],
            "data_freshness": self.data_freshness(),
        }

    def get_dashboard(self, interval: str) -> Dict[str, object]:
        payload = self.get_interval_payload(interval)
        return {
            "interval": interval,
            "series": {
                "bars": self._chart_bars(payload["bars"], interval),
                "full_bar_count": len(payload["bars"]),
                "chart_bar_count": len(self._chart_bars(payload["bars"], interval)),
                "latest_date": payload["bars"][-1]["date"],
            },
            "events": payload["top_events"],
            "summary": payload["summary"],
            "summary_en": payload["summary_en"],
            "data_freshness": self.data_freshness(),
        }

    def get_events(self, interval: str) -> Dict[str, object]:
        payload = self.get_interval_payload(interval)
        return {
            "interval": interval,
            "events": payload["top_events"],
            "data_freshness": self.data_freshness(),
        }

    def get_event_detail(self, interval: str, event_id: str) -> Dict[str, object]:
        payload = self.get_interval_payload(interval)
        for event in payload["events"]:
            if event["event_id"] == event_id:
                return {
                    "interval": interval,
                    "event": event,
                    "data_freshness": self.data_freshness(),
                }
        raise KeyError(event_id)

    def get_summary(self, interval: str) -> Dict[str, object]:
        payload = self.get_interval_payload(interval)
        return {
            "interval": interval,
            "summary": payload["summary"],
            "summary_en": payload["summary_en"],
            "data_freshness": self.data_freshness(),
        }

    def data_freshness(self) -> Dict[str, str]:
        yahoo_latest = self.yahoo_bars[-1].date.isoformat() if self.yahoo_bars else self.local_latest_date.isoformat()
        return {
            "local_base_end": self.local_latest_date.isoformat(),
            "latest_merged_date": self.latest_merged_date.isoformat(),
            "yahoo_latest": yahoo_latest,
        }
