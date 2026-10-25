# -*- coding: utf-8 -*-
"""Small-cap rotation research: standalone Bloomberg-to-Excel workbook.

QUICK START (Python 3.9+; Bloomberg Desktop logged in):
    python rotation_research_bbg.py
    python rotation_research_bbg.py --backend pdblp
    python rotation_research_bbg.py --end 2026-10-23 --out rotation_output
    python rotation_research_bbg.py --self-test

Dependencies: numpy, pandas, matplotlib, openpyxl; ONE existing Bloomberg
connector: xbbg / pdblp / blpapi. The script never downloads from Yahoo.
Everything required for company use is in THIS file. No local sibling imports.
Default outputs go to ../iwm_results, OUTSIDE the folder containing this script.
Use --out to choose a different location. The copy folder needs only this .py.

Excel is the primary deliverable: raw prices, daily returns, rolling correlations,
100-day moving average, relative strength, drawdowns, dated signal observations,
forward outcomes, event studies, methodology and editable native Excel charts.
The Formula_Audit sheet contains real Excel formulas for independent checking.
HTML and figures use the exact same tables. All content is in English.

All baseline returns are split-adjusted PRICE returns, excluding cash dividends.
US10Y is a yield LEVEL in percent, e.g. 4.5 means 4.5%; changes are in bp.
Correlation uses daily simple RETURNS, never price levels. 3 months = 63 sessions.
No missing prices are forward filled. SPX valid dates define the master calendar.
All rolling windows require complete observations. AAPL is an NDX constituent;
AAPL/NDX is not an independent or ex-Apple comparison.

Signal studies are exploratory, not an optimized trading strategy. Signals are
observed at completed month-end. Entry is next-session CLOSE, and exit is 21/63
sessions after entry. Outcomes are gross of costs and omit dividends. We report
both overlapping month-end observations and a greedy non-overlapping sample
for each rule/horizon. Samples differ across rules; this is not causal evidence.
No flow data, historical constituents or earnings revisions are downloaded:
relative performance alone cannot demonstrate actual net inflows into small caps.

External Yahoo verification and historical research live OUTSIDE the copy folder.
This script recalculates all market results from its own Bloomberg download.
Historical research dates are fixed illustrative cases, not live trading signals.
"""

START_DATE = "1980-01-01"
END_DATE = None                      # None = yesterday in America/New_York
DISPLAY_START = "2023-01-01"
DRAWDOWN_START = "2024-10-18"
BACKEND = "auto"
OUTPUT_DIR = None
FIELD = "PX_LAST"
BBG_HOST = "localhost"
BBG_PORT = 8194
REQUEST_TIMEOUT_SECONDS = 120
CORRELATION_WINDOW = 63
FORWARD_HORIZONS = (21, 63)
SIGNAL_START = "2007-01-01"          # common research era, not a 30-year ETF claim
TICKERS = {
    "AAPL":"AAPL US Equity", "NDX":"NDX Index", "SPX":"SPX Index",
    "IWM":"IWM US Equity", "SPY":"SPY US Equity", "QQQ":"QQQ US Equity",
    "XLV":"XLV US Equity", "XBI":"XBI US Equity", "IBB":"IBB US Equity",
    "RSP":"RSP US Equity", "MTUM":"MTUM US Equity", "SPMO":"SPMO US Equity",
    "KRE":"KRE US Equity", "IJR":"IJR US Equity", "US10Y":"USGG10YR Index",
    "VIX":"VIX Index", "HYG":"HYG US Equity", "LQD":"LQD US Equity",
    "NVDA":"NVDA US Equity", "MSFT":"MSFT US Equity",
    "UNH":"UNH US Equity", "LLY":"LLY US Equity",
}

import argparse
import base64
import html
import importlib.util
import logging
import math
import os
from pathlib import Path
import tempfile
import time
from datetime import datetime, timedelta, timezone
import numpy as np
import pandas as pd

LOG = logging.getLogger("rotation_research")

def normalize_bdh(raw, ticker):
    """Accept legacy wide MultiIndex and modern long-form Bloomberg frames."""
    if hasattr(raw, "to_pandas"): raw = raw.to_pandas()
    if isinstance(raw, pd.Series): return raw.rename(ticker)
    if raw is None or len(raw) == 0: raise ValueError("Empty Bloomberg response for " + ticker)
    raw = raw.copy()
    if isinstance(raw.columns, pd.MultiIndex):
        match = [c for c in raw.columns if ticker.upper() in [str(x).upper() for x in c]
                 and FIELD.upper() in [str(x).upper() for x in c]]
        if len(match) != 1: raise ValueError("Cannot identify ticker/field in Bloomberg columns: " + str(raw.columns))
        return raw[match[0]].rename(ticker)
    lower = {str(c).lower(): c for c in raw.columns}
    date_col = next((lower[k] for k in ["date", "time", "datetime"] if k in lower), None)
    if date_col is not None:
        security = next((lower[k] for k in ["ticker", "security"] if k in lower), None)
        if security is not None: raw = raw[raw[security].astype(str).str.upper() == ticker.upper()]
        if "field" in lower: raw = raw[raw[lower["field"]].astype(str).str.upper() == FIELD.upper()]
        value_col = lower.get("value", lower.get(FIELD.lower()))
        if value_col is None: raise ValueError("Cannot identify Bloomberg value column")
        result = pd.Series(raw[value_col].to_numpy(), index=pd.to_datetime(raw[date_col]), name=ticker)
    elif FIELD.lower() in lower:
        result = raw[lower[FIELD.lower()]].rename(ticker)
    elif ticker.lower() in lower:
        result = raw[lower[ticker.lower()]].rename(ticker)
    else:
        raise ValueError("Unsupported Bloomberg response layout: " + str(raw.columns))
    if result.empty: raise ValueError("No values for " + ticker + " / " + FIELD)
    return result


def native_bdh(ticker, start, end):
    """Desktop Bloomberg reference-data request; handles partial responses/errors."""
    import blpapi
    options = blpapi.SessionOptions()
    options.setServerHost(BBG_HOST)
    options.setServerPort(BBG_PORT)
    session = blpapi.Session(options)
    try:
        if not session.start(): raise RuntimeError("Bloomberg session could not start; check Terminal/API connection")
        if not session.openService("//blp/refdata"): raise RuntimeError("Cannot open //blp/refdata")
        request = session.getService("//blp/refdata").createRequest("HistoricalDataRequest")
        request.getElement("securities").appendValue(ticker)
        request.getElement("fields").appendValue(FIELD)
        request.set("startDate", start.strftime("%Y%m%d"))
        request.set("endDate", end.strftime("%Y%m%d"))
        request.set("periodicitySelection", "DAILY")
        request.set("nonTradingDayFillOption", "ACTIVE_DAYS_ONLY")
        request.set("adjustmentNormal", False)
        request.set("adjustmentAbnormal", False)
        request.set("adjustmentSplit", True)
        request.set("adjustmentFollowDPDF", False)
        session.sendRequest(request)
        deadline = time.monotonic() + REQUEST_TIMEOUT_SECONDS
        rows = []
        while time.monotonic() < deadline:
            event = session.nextEvent(1000)
            for msg in event:
                if msg.hasElement("responseError"): raise RuntimeError(str(msg.getElement("responseError")))
                if str(msg.messageType()) in ["RequestFailure", "SessionTerminated"]:
                    raise RuntimeError(str(msg))
                if not msg.hasElement("securityData"): continue
                sec = msg.getElement("securityData")
                if sec.hasElement("securityError"): raise RuntimeError(str(sec.getElement("securityError")))
                if sec.hasElement("fieldExceptions") and sec.getElement("fieldExceptions").numValues():
                    raise RuntimeError(str(sec.getElement("fieldExceptions")))
                data = sec.getElement("fieldData")
                for i in range(data.numValues()):
                    item = data.getValueAsElement(i)
                    if item.hasElement(FIELD) and not item.getElement(FIELD).isNull():
                        rows.append((pd.Timestamp(item.getElementAsDatetime("date")), item.getElementAsFloat(FIELD)))
            if event.eventType() == blpapi.Event.RESPONSE:
                if not rows: raise ValueError("No Bloomberg observations for " + ticker)
                return pd.Series([v for _, v in rows], index=[t for t, _ in rows], name=ticker)
        raise TimeoutError("Bloomberg request timed out: " + ticker)
    finally:
        session.stop()


def pull_bloomberg(config, backend="auto"):
    """Replace THIS function if your firm has its own Bloomberg data loader."""
    if backend == "auto":
        backend = next((name for name in ["xbbg", "pdblp", "blpapi"]
                        if importlib.util.find_spec(name) is not None), None)
    if backend is None:
        raise RuntimeError("No Bloomberg package found. Use your firm's xbbg/pdblp/blpapi environment, or --csv daily.csv")
    config["source"] = "Bloomberg / " + backend
    LOG.info("Backend: %s; request %s to %s", backend, config["start"].date(), config["end"].date())
    collected = {}
    con = None
    try:
        if backend == "pdblp":
            import pdblp
            con = pdblp.BCon(debug=False, port=BBG_PORT, timeout=REQUEST_TIMEOUT_SECONDS * 1000)
            con.start()
        for key, ticker in config["tickers"].items():
            try:
                LOG.info("Pull %s: %s / %s", key, ticker, FIELD)
                if backend == "xbbg":
                    from xbbg import blp
                    raw = blp.bdh(tickers=ticker, flds=FIELD,
                                  start_date=config["start"].strftime("%Y%m%d"),
                                  end_date=config["end"].strftime("%Y%m%d"), Per="D",
                                  adjustmentNormal=False, adjustmentAbnormal=False,
                                  adjustmentSplit=True, adjustmentFollowDPDF=False)
                    series = normalize_bdh(raw, ticker)
                elif backend == "pdblp":
                    raw = con.bdh([ticker], [FIELD], config["start"].strftime("%Y%m%d"),
                                  config["end"].strftime("%Y%m%d"),
                                  elms=[("periodicitySelection", "DAILY"),
                                        ("adjustmentNormal", False), ("adjustmentAbnormal", False),
                                        ("adjustmentSplit", True), ("adjustmentFollowDPDF", False)])
                    series = normalize_bdh(raw, ticker)
                else:
                    series = native_bdh(ticker, config["start"], config["end"])
                series.index = pd.to_datetime(series.index)
                collected[key] = series
            except Exception as exc:
                raise RuntimeError("Required series pull failed: " + ticker + " / " + str(exc)) from exc
    finally:
        if con is not None: con.stop()
    return pd.DataFrame(collected)


def clean_data(raw, end):
    p = raw.copy()
    if not isinstance(p.index, pd.DatetimeIndex):
        raise ValueError('Data must have an explicit DatetimeIndex')
    if p.index.tz is not None: p.index = p.index.tz_localize(None)
    p.index = p.index.normalize()
    if p.index.has_duplicates or p.index.isna().any(): raise ValueError('Invalid/duplicate dates')
    if p.columns.has_duplicates: raise ValueError('Duplicate columns')
    missing = set(TICKERS) - set(p.columns)
    if missing: raise ValueError('Missing required series: ' + ', '.join(sorted(missing)))
    p = p[list(TICKERS)].apply(pd.to_numeric, errors='raise').astype(float).sort_index().loc[:end]
    if np.isinf(p.to_numpy()).any() or (p.drop(columns=['US10Y']) <= 0).any().any():
        raise ValueError('Nonfinite or nonpositive price data')
    coverage = []
    for c in p:
        s = p[c].dropna()
        if len(s) < 150: raise ValueError('Too little history: ' + c)
        coverage.append(dict(series=c, ticker=TICKERS[c], first=s.index[0], last=s.index[-1],
                             observations=len(s), stale_calendar_days=(pd.Timestamp(end)-s.index[-1]).days))
    # Remove holidays on which a non-equity series alone has a print.
    p = p.reindex(p.SPX.dropna().index)
    p.index.name = 'Date'
    cov = pd.DataFrame(coverage)
    cov['missing_inside_SPX_calendar'] = [int(p.loc[row['first']:row['last'], row['series']].isna().sum()) for row in coverage]
    return p, cov


def relative_path(p, a, b):
    q = p[a] / p[b]
    valid = q.dropna()
    return q / valid.iloc[0] * 100 if len(valid) else q


def indicators(p):
    returns = p.pct_change(fill_method=None)
    returns['US10Y'] = p.US10Y.diff() * 100  # bp, not a bond return
    correlations = pd.DataFrame(index=p.index)
    pairs = [('AAPL','NDX'),('AAPL','XLV'),('XLV','NDX'),('XBI','NDX'),
             ('IWM','NDX'),('IWM','SPY'),('NVDA','NDX'),('MSFT','NDX'),('UNH','NDX'),('LLY','NDX')]
    for a,b in pairs:
        correlations[a+'_'+b+'_63d'] = returns[a].rolling(CORRELATION_WINDOW, min_periods=CORRELATION_WINDOW).corr(returns[b])
    correlations['AAPL_NDX_21d'] = returns.AAPL.rolling(21).corr(returns.NDX)
    correlations['AAPL_NDX_126d'] = returns.AAPL.rolling(126).corr(returns.NDX)
    correlations['IWM_US10Y_change_63d'] = returns.IWM.rolling(63).corr(returns.US10Y)
    sma = pd.DataFrame({'SPX':p.SPX, 'SPX_SMA100':p.SPX.rolling(100).mean()})
    sma['distance_to_SMA100'] = sma.SPX / sma.SPX_SMA100 - 1
    sma['cross_below'] = (sma.distance_to_SMA100 < 0) & (sma.distance_to_SMA100.shift(1) >= 0)
    rel = pd.DataFrame(index=p.index)
    for a,b in [('IWM','SPY'),('RSP','SPY'),('XLV','QQQ'),('XBI','IBB'),('KRE','SPY'),('IJR','IWM'),('MTUM','SPY')]:
        rel[a+'_'+b+'_rebased'] = relative_path(p,a,b)
    ddp = p.loc[DRAWDOWN_START:, ['MTUM','SPMO','SPX']]
    dd = ddp / ddp.cummax() - 1
    dd.columns = [c+'_drawdown' for c in dd]
    sig = pd.DataFrame(index=p.index)
    sig['AAPL_NDX_corr63'] = correlations.AAPL_NDX_63d
    sig['AAPL_XLV_corr63'] = correlations.AAPL_XLV_63d
    sig['XLV_NDX_corr63'] = correlations.XLV_NDX_63d
    trailing = p.pct_change(21, fill_method=None)
    for a,b in [('IWM','SPY'),('RSP','SPY'),('XLV','QQQ'),('XBI','IBB'),('KRE','SPY')]:
        sig[a+'_minus_'+b+'_21d'] = trailing[a] - trailing[b]
    for a in ['IWM','SPY','AAPL','XLV','XBI','KRE','MTUM','NVDA']:
        sig[a+'_return_21d'] = trailing[a]
    sig['US10Y_change_21d_bp'] = (p.US10Y - p.US10Y.shift(21))*100
    sig['SPX_distance_SMA100'] = sma.distance_to_SMA100
    sig['VIX_level'] = p.VIX
    sig['US10Y_level_pct'] = p.US10Y
    sig['HYG_return_21d'] = trailing.HYG  # NOT a credit-spread measurement
    return returns, correlations, sma, rel, dd, sig


def rule_masks(sig):
    def flag(cols, test):
        available = sig[cols].notna().all(axis=1)
        return pd.Series(np.where(available, test.astype(float), np.nan), index=sig.index)
    rules = pd.DataFrame(index=sig.index)
    rules['Baseline'] = 1.
    rules['AAPL_decoupling'] = flag(['AAPL_NDX_corr63'], sig.AAPL_NDX_corr63 < .25)
    rules['AAPL_negative'] = flag(['AAPL_NDX_corr63'], sig.AAPL_NDX_corr63 < 0)
    rules['Healthcare_leadership'] = flag(['XLV_minus_QQQ_21d'], sig.XLV_minus_QQQ_21d > 0)
    rules['Breadth_improvement'] = flag(['RSP_minus_SPY_21d'], sig.RSP_minus_SPY_21d > 0)
    rules['Yield_relief'] = flag(['US10Y_change_21d_bp'], sig.US10Y_change_21d_bp < 0)
    rules['Breadth_plus_yield'] = flag(['RSP_minus_SPY_21d','US10Y_change_21d_bp'],
              (sig.RSP_minus_SPY_21d > 0) & (sig.US10Y_change_21d_bp < 0))
    rules['Rotation_confirmation'] = flag(['RSP_minus_SPY_21d','US10Y_change_21d_bp','KRE_return_21d','SPX_distance_SMA100'],
              (sig.RSP_minus_SPY_21d > 0) & (sig.US10Y_change_21d_bp < 0) &
              (sig.KRE_return_21d > 0) & (sig.SPX_distance_SMA100 > 0))
    return rules


def conditional_study(p, sig, asof):
    rules = rule_masks(sig)
    # Exclude the current unfinished month even if its last available price is stale.
    cutoff = min(p.index[-1], pd.Timestamp(asof).normalize()-pd.Timedelta(days=1))
    period = cutoff.to_period('M')
    expected_final_weekday = pd.offsets.BMonthEnd().rollback(period.end_time.normalize())
    # Conservative month completeness: without an exchange calendar, a month
    # ending on an exchange-only holiday may be excluded, never assumed complete.
    complete = p.index.to_period('M') < period
    if cutoff >= expected_final_weekday:
        complete = p.index.to_period('M') <= period
    ends = pd.Series(p.index[complete], index=p.index[complete]).groupby(p.index[complete].to_period('M')).last()
    dates = pd.DatetimeIndex(ends.values)
    dates = dates[dates >= pd.Timestamp(SIGNAL_START)]
    monthly = sig.reindex(dates).join(rules.reindex(dates).add_prefix('rule_'))
    records = []
    for d in dates:
        pos = p.index.get_loc(d)
        if pos+1 >= len(p): continue
        for h in FORWARD_HORIZONS:
            entry, exitpos = pos+1, pos+1+h
            valid_end = exitpos < len(p)
            record = dict(signal_date=d, entry_date=p.index[entry], horizon=h,
                          exit_date=p.index[exitpos] if valid_end else pd.NaT,
                          IWM_forward=np.nan, SPY_forward=np.nan, excess_forward=np.nan)
            if valid_end:
                path = p.iloc[entry:exitpos+1][['IWM','SPY']]
                if path.notna().all().all():
                    f = path.iloc[-1]/path.iloc[0]-1
                    record.update(IWM_forward=f.IWM,SPY_forward=f.SPY,excess_forward=f.IWM-f.SPY)
            record.update(rules.loc[d].to_dict())
            records.append(record)
    forward = pd.DataFrame(records)
    if forward.empty: raise ValueError('No complete month-end research observations')
    stats, selected = [], []
    for h in FORWARD_HORIZONS:
        for rule in rules:
            eligible = forward[(forward.horizon==h) & forward.excess_forward.notna() & forward[rule].notna()]
            matched = eligible[eligible[rule]==1]
            for mode in ['All_month_ends','Non_overlapping']:
                sample = matched
                if mode == 'Non_overlapping':
                    keep, last_exit = [], pd.Timestamp.min
                    for ix, row in matched.iterrows():
                        if row.entry_date > last_exit:
                            keep.append(ix); last_exit = row.exit_date
                    sample = matched.loc[keep]
                    for _, row in sample.iterrows(): selected.append(dict(rule=rule, **row.to_dict()))
                n = len(sample)
                stats.append(dict(rule=rule, horizon=h, sample_mode=mode, n=n,
                    eligible_months=len(eligible), mean_IWM=sample.IWM_forward.mean(),
                    median_IWM=sample.IWM_forward.median(), IWM_win_rate=(sample.IWM_forward>0).mean() if n else np.nan,
                    mean_excess=sample.excess_forward.mean(), median_excess=sample.excess_forward.median(),
                    excess_win_rate=(sample.excess_forward>0).mean() if n else np.nan,
                    eligible_baseline_mean_excess=eligible.excess_forward.mean(),
                    mean_excess_lift=sample.excess_forward.mean()-eligible.excess_forward.mean(),
                    min_excess=sample.excess_forward.min(),max_excess=sample.excess_forward.max(),
                    interpretation='Exploratory; no causal inference; low power' if n<20 else 'Exploratory; overlapping observations are dependent'))
    return monthly, forward, pd.DataFrame(stats), pd.DataFrame(selected)


def event_study(p, sig):
    windows = [
        ('2023_Nov_Dec','2023-10-31','2023-12-31','Fixed calendar comparison'),
        ('2024_Nov','2024-10-31','2024-11-30','Fixed calendar comparison'),
        ('2024_Dec','2024-11-30','2024-12-31','Fixed calendar comparison'),
        ('2025_Oct','2025-09-30','2025-10-31','Pre-November calendar window'),
        ('2025_Nov_selloff','2025-10-31','2025-11-20','Retrospective trough endpoint; NOT tradable timing'),
        ('2025_Rebound','2025-11-20','2025-12-11','Retrospective trough-to-rally window; NOT a backtest'),
        ('2025_Nov','2025-10-31','2025-11-30','Fixed calendar comparison'),
        ('2025_Dec','2025-11-30','2025-12-31','Fixed calendar comparison'),
        ('2025_Nov_Dec','2025-10-31','2025-12-31','Fixed calendar comparison'),
    ]
    rows=[]
    for label,a,b,kind in windows:
        if p.index[-1] < pd.Timestamp(b) - pd.Timedelta(days=3): continue
        prior=p.loc[:a];finish=p.loc[:b]
        if prior.empty or finish.empty: continue
        x,y=prior.iloc[-1],finish.iloc[-1]
        for key in p:
            if key in ['US10Y','VIX']: continue
            path=p.loc[prior.index[-1]:finish.index[-1],key]
            if path.isna().any(): value=dd=np.nan
            else:
                value=y[key]/x[key]-1
                dd=(path/path.cummax()-1).min()
            rows.append(dict(window=label, description=kind,series=key,start=prior.index[-1],end=finish.index[-1],
                             price_return=value,max_drawdown=dd,US10Y_change_bp=(y.US10Y-x.US10Y)*100))
    checkpoints=['2023-10-31','2024-10-31','2025-09-30','2025-10-31','2025-11-20','2025-11-28','2025-12-31']
    checks=[]
    for target in checkpoints+[str(p.index[-1].date())]:
        if pd.Timestamp(target)>p.index[-1]: continue
        z=sig.loc[:target]
        if z.empty: continue
        d=z.index[-1]
        for key in sig:
            history=sig.loc[sig.index<d,key].dropna()
            value=sig.loc[d,key]
            percentile=float((history<=value).mean()) if len(history)>=252 and pd.notna(value) else np.nan
            checks.append(dict(date=d,metric=key,value=value,prior_history_percentile=percentile,
                               prior_observations=len(history),note='Historical checkpoint; not an ex-ante selected trade date'))
    return pd.DataFrame(rows,columns=['window','description','series','start','end','price_return','max_drawdown','US10Y_change_bp']),pd.DataFrame(checks)



def setup_comparison(p, t):
    """Equal-length trailing windows and explicit retrospective episode windows."""
    cases=[('2025_Oct_tariff_shock','2025-10-10'),('2025_Oct_end','2025-10-31'),
           ('2025_Nov_selloff_low','2025-11-20'),('2025_Nov_rebound','2025-11-28'),
           ('2025_Dec_rally','2025-12-11'),('Current',str(p.index[-1].date()))]
    rows=[]
    keys=['QQQ','NVDA','MTUM','IWM','SPY','RSP','XLV','XBI','KRE','AAPL']
    for label,date in cases:
        if pd.Timestamp(date)>p.index[-1]:continue
        pos=p.index.searchsorted(pd.Timestamp(date),side='right')-1
        if pos<0:continue
        for n in [1,5,10,21,63]:
            if pos<n:continue
            start,end=p.index[pos-n],p.index[pos]
            path=p.loc[start:end,keys]
            ret=(path.iloc[-1]/path.iloc[0]-1).where(path.notna().all())
            row=dict(case=label,date=end,lookback_sessions=n,start=start)
            row.update({k+'_return':ret[k] for k in keys})
            row.update(IWM_minus_SPY=ret.IWM-ret.SPY,RSP_minus_SPY=ret.RSP-ret.SPY,
                       XLV_minus_QQQ=ret.XLV-ret.QQQ,
                       US10Y_change_bp=(p.US10Y.iloc[pos]-p.US10Y.iloc[pos-n])*100,
                       US10Y_level_pct=p.US10Y.iloc[pos],VIX=p.VIX.iloc[pos],
                       AAPL_NDX_corr63=t['Correlations'].AAPL_NDX_63d.iloc[pos],
                       SPX_distance_SMA100=t['SPX_100DMA'].distance_to_SMA100.iloc[pos])
            rows.append(row)
    setup=pd.DataFrame(rows)
    windows=[('2025_Oct_shock','2025-10-09','2025-10-10'),
       ('2025_Oct_rebound','2025-10-10','2025-10-28'),
       ('2025_Nov_derisk','2025-10-28','2025-11-20'),
       ('2025_Nov_first_rebound','2025-11-20','2025-11-28'),
       ('2025_Nov_Dec_rally','2025-11-20','2025-12-11'),
       ('2025_Dec_fade','2025-12-11','2025-12-31'),
       ('2025_Nov_Dec_total','2025-10-31','2025-12-31'),
       ('2026_Oct_recent_derisk','2026-10-07','2026-10-23'),
       ('2026_Oct_22_shock','2026-10-21','2026-10-22'),
       ('2026_Oct_23_followup','2026-10-22','2026-10-23')]
    episodes=[]
    for label,a,b in windows:
        if pd.Timestamp(b)>p.index[-1] or pd.Timestamp(a)<p.index[0]:continue
        z=p.loc[:a];q=p.loc[:b];x,y=z.iloc[-1],q.iloc[-1]
        ret=y[keys]/x[keys]-1
        rec=dict(episode=label,start=z.index[-1],end=q.index[-1],
                 US10Y_start_pct=x.US10Y,US10Y_end_pct=y.US10Y,US10Y_change_bp=(y.US10Y-x.US10Y)*100,
                 IWM_minus_SPY=ret.IWM-ret.SPY,IWM_minus_QQQ=ret.IWM-ret.QQQ)
        rec.update({k+'_return':ret[k] for k in keys});episodes.append(rec)
    # Relative event time: display historical outcomes, truncate the current case
    # at observed prices. Never extend the current trajectory with an old rally.
    paths=[]
    for label,date in [('Oct_2025','2025-10-10'),('Nov_2025','2025-11-20'),('Oct_2026','2026-10-22')]:
        d=pd.Timestamp(date)
        if d not in p.index:continue
        pos=p.index.get_loc(d);base=p.iloc[pos]
        for j in range(max(0,pos-21),min(len(p),pos+22)):
            r=p.iloc[j]/base-1
            paths.append(dict(case=label,anchor=d,date=p.index[j],session=j-pos,IWM_return=r.IWM,
                              QQQ_return=r.QQQ,IWM_minus_SPY=r.IWM-r.SPY,
                              US10Y_change_bp=(p.US10Y.iloc[j]-base.US10Y)*100))
    return setup,pd.DataFrame(episodes),pd.DataFrame(paths)


def tech_shock_study(p):
    """Predefined daily tech/rates conditions, not retrospective 'AI' labels.

    Since 2013; QQQ <= -2% close-to-close; relief means 10Y change <= -5 bp.
    Entries next close; 5/21/63-session exits. A single tech shock does not
    identify its economic cause. Return comparisons are arithmetic differences.
    """
    r=p.pct_change(fill_method=None);dy=p.US10Y.diff()*100
    valid=p[['QQQ','IWM','SPY','US10Y']].notna().all(axis=1)&r.QQQ.notna()&dy.notna()
    valid &= p.index>=pd.Timestamp('2013-01-01')
    rules={'All_sessions':valid,'Tech_selloff':valid&(r.QQQ<=-.02),
           'Tech_selloff_yield_relief':valid&(r.QQQ<=-.02)&(dy<=-5),
           'Tech_relief_IWM_resilience':valid&(r.QQQ<=-.02)&(dy<=-5)&(r.IWM>r.SPY)}
    stats=[];outcomes=[];selected=[]
    for h in [5,21,63]:
        forward=pd.DataFrame(index=p.index)
        for k in ['IWM','SPY']:
            f=p[k].shift(-(h+1))/p[k].shift(-1)-1
            complete=p[k].notna().rolling(h+1).sum().shift(-(h+1))==(h+1)
            forward[k]=f.where(complete)
        forward['excess']=forward.IWM-forward.SPY
        eligible=forward.excess.notna()
        for rule,mask in rules.items():
            dates=p.index[mask]
            sample=forward.loc[dates].dropna()
            if rule!='All_sessions':
                for d in dates:
                    i=p.index.get_loc(d);e=i+1+h
                    outcomes.append(dict(rule=rule,horizon=h,signal_date=d,
                        QQQ_day_return=r.loc[d,'QQQ'],US10Y_day_change_bp=dy.loc[d],
                        IWM_day_minus_SPY=r.loc[d,'IWM']-r.loc[d,'SPY'],
                        entry_date=p.index[i+1] if i+1<len(p) else pd.NaT,
                        exit_date=p.index[e] if e<len(p) else pd.NaT,
                        IWM_forward=forward.loc[d,'IWM'],SPY_forward=forward.loc[d,'SPY'],excess_forward=forward.loc[d,'excess']))
            keep=[];last_exit=-1
            for d in sample.index:
                i=p.index.get_loc(d)
                if i+1>last_exit:keep.append(d);last_exit=i+1+h
            for mode,subset in [('All_signals',sample),('Non_overlapping',sample.loc[keep])]:
                n=len(subset)
                stats.append(dict(rule=rule,horizon=h,sample_mode=mode,n=n,
                    mean_IWM=subset.IWM.mean(),median_IWM=subset.IWM.median(),
                    IWM_win_rate=(subset.IWM>0).mean() if n else np.nan,
                    mean_excess=subset.excess.mean(),excess_win_rate=(subset.excess>0).mean() if n else np.nan,
                    mean_excess_before_2020=subset.loc[:'2019-12-31','excess'].mean(),
                    mean_excess_since_2020=subset.loc['2020-01-01':,'excess'].mean()))
                if mode=='Non_overlapping':
                    for d in subset.index:
                        i=p.index.get_loc(d)
                        selected.append(dict(rule=rule,horizon=h,signal_date=d,entry_date=p.index[i+1],exit_date=p.index[i+1+h],
                                             IWM_forward=subset.loc[d,'IWM'],excess_forward=subset.loc[d,'excess']))
    return pd.DataFrame(stats),pd.DataFrame(outcomes),pd.DataFrame(selected)


def add_setup_tables(t):
    setup,episodes,paths=setup_comparison(t['Raw_Prices'],t)
    stats,events,nonoverlap=tech_shock_study(t['Raw_Prices'])
    t.update(Setup_Comparison=setup,Episode_Sequence=episodes,Event_Aligned_Paths=paths,
             Tech_Shock_Stats=stats,Tech_Shock_Outcomes=events,Tech_Shock_NonOverlap=nonoverlap)
    for metric in ['IWM_return','IWM_minus_SPY','US10Y_change_bp']:
        name={'IWM_return':'Chart_Event_IWM','IWM_minus_SPY':'Chart_Event_Excess','US10Y_change_bp':'Chart_Event_Yield'}[metric]
        if not paths.empty:t[name]=paths.pivot(index='session',columns='case',values=metric).reset_index()
    return t


def setup_findings(t):
    s=t['Setup_Comparison'];e=t['Episode_Sequence'].set_index('episode')
    now=s[s['case']=='Current'].set_index('lookback_sessions')
    lines=['The tactical rotation case has improved at short horizons. This is a conditional early-rotation thesis, '
           'not proof of durable small-cap leadership or measured fund inflows. QQQ is a technology/growth proxy, '
           'not a pure AI portfolio; headlines identify event context, while prices quantify the pattern.']
    for n in [5,10,21]:
        if n in now.index:
            z=now.loc[n]
            lines.append('Latest %d sessions (%s to %s): QQQ %+.2f%%, IWM %+.2f%%, IWM minus SPY %+.2f pp, '
                         'RSP minus SPY %+.2f pp, KRE %+.2f%% and 10-year yield %+.1f bp.' %
                         (n,str(z.start.date()),str(z.date.date()),100*z.QQQ_return,100*z.IWM_return,
                          100*z.IWM_minus_SPY,100*z.RSP_minus_SPY,100*z.KRE_return,z.US10Y_change_bp))
    if 5 in now.index and 21 in now.index:
        if now.loc[5,'IWM_minus_SPY']>0 and now.loc[21,'IWM_minus_SPY']<0:
            lines.append('Time-horizon reconciliation: short-window relative strength is improving while the 21-session '
                         'record still lags. A 21-session-only conclusion misses the emerging turn; a one-day-only '
                         'conclusion would overstate its durability.')
    for key in e.index:
        z=e.loc[key]
        lines.append('%s (%s to %s): QQQ %+.2f%%, IWM %+.2f%%, IWM minus SPY %+.2f pp, yield %+.1f bp.' %
                     (key,str(z.start.date()),str(z.end.date()),100*z.QQQ_return,100*z.IWM_return,
                      100*z.IWM_minus_SPY,z.US10Y_change_bp))
    lines.extend([
      'Historical interpretation: October 10, 2025 was primarily a tariff-risk shock, although AI valuations were '
      'also a concern. Its subsequent October rebound favored QQQ over IWM. The cleaner small-cap handoff '
      'was the late-November rebound after the AI/momentum sell-off. October as a whole was still tech-led.',
      'The November 2025 sequence was not a continuous fall in long yields followed mechanically by an IWM rally. '
      'Long yields rose over the late-October-to-November-20 sell-off, fell during the first rebound week, '
      'and were slightly above the November-20 level by December 11. Rates relief was part of the setup, not '
      'a sufficient explanation for the whole rally.',
      'Similarities: technology/momentum weakness, relative healthcare resilience, Treasury buying and room for '
      'broader participation. Current RSP and regional-bank strength make the short-horizon rotation case more '
      'credible than Apple decorrelation alone.',
      'Differences: current long yields remain substantially above their late-2025 levels. Recent oil and inflation '
      'pressure also differ from the falling-oil backdrop documented in November 2025. A safe-haven bond rally can '
      'coincide with growth stress; it is not automatically benign easing for small-cap earnings.',
      'Trade argument: easing long yields may remove a constraint on small caps while crowded technology '
      'positions unwind. Look for IWM/SPY strength to persist across 5/10/21 sessions, with KRE/RSP and biotech '
      'participation. Fading relative strength or renewed yield pressure would weaken this thesis. This is '
      'a price-based framework; earnings, credit spreads and net fund flows still need separate confirmation.',
    ])
    z=t['Tech_Shock_Stats']
    for rule in ['Tech_selloff','Tech_selloff_yield_relief','Tech_relief_IWM_resilience']:
        row=z[(z.rule==rule)&(z.horizon==21)&(z.sample_mode=='Non_overlapping')].iloc[0]
        if row.n:
            lines.append('Historical screen %s: %d non-overlapping episodes, next-21-session mean IWM return '
                         '%+.2f%%, mean excess over SPY %+.2f pp, relative win rate %.1f%%. Signals are observed '
                         'at close and entered at the next close. These are descriptive, regime-dependent results, '
                         'not an optimized strategy or causal estimate.' %
                         (rule,row.n,100*row.mean_IWM,100*row.mean_excess,100*row.excess_win_rate))
    lines.append('Historical rebound endpoints and event anchors are retrospective illustrations. The event-aligned '
                 'chart shows the future of old episodes, never an assumed future for today. Fixed thresholds in '
                 'the daily study are QQQ <= -2% and 10-year yield change <= -5 bp; the resilience subset additionally '
                 'requires IWM to outperform SPY that day. No claim is made that all such events were AI-driven.')
    return lines


def setup_report(t,config,out):
    import matplotlib.pyplot as plt
    from matplotlib.ticker import PercentFormatter
    findings=setup_findings(t)
    fig,axes=plt.subplots(3,1,figsize=(12,10),sharex=True)
    paths=t['Event_Aligned_Paths']
    for label,color in [('Oct_2025','#8291a6'),('Nov_2025','#080878'),('Oct_2026','#168b75')]:
        q=paths[paths['case']==label] if not paths.empty else paths
        if q.empty:continue
        for ax,metric in zip(axes,['IWM_return','IWM_minus_SPY','US10Y_change_bp']):
            ax.plot(q.session,q[metric],color=color,lw=2,label=label.replace('_',' '))
            ax.axvline(0,color='#aaa',ls='--',lw=.7);ax.axhline(0,color='#ccc',lw=.7)
    axes[0].set_ylabel('IWM price return');axes[0].yaxis.set_major_formatter(PercentFormatter(1))
    axes[1].set_ylabel('IWM minus SPY');axes[1].yaxis.set_major_formatter(PercentFormatter(1))
    axes[2].set_ylabel('10Y change (bp)');axes[2].set_xlabel('Trading sessions relative to event anchor')
    axes[0].legend(loc='upper left',frameon=False,ncol=3)
    fig.suptitle('Compare the setup, not just the subsequent rally',x=.08,ha='left',weight='bold',fontsize=17)
    fig.text(.08,.02,'Anchors: Oct 10 2025 / Nov 20 2025 / Oct 22 2026. Current path stops at available observations.\n'
             'Historical paths show up to 21 subsequent sessions; retrospective illustrations, not forecasts. Source: '+config['source'],fontsize=8)
    fig.subplots_adjust(left=.09,right=.96,top=.92,bottom=.11,hspace=.20)
    for ext in ['png','svg']:fig.savefig(out/('setup_event_comparison.'+ext),dpi=180,facecolor='white')
    plt.close(fig)
    compare=t['Setup_Comparison'].query('lookback_sessions == 10')[
        ['case','date','QQQ_return','IWM_return','IWM_minus_SPY','RSP_minus_SPY','KRE_return','XLV_return','US10Y_change_bp','US10Y_level_pct']].copy()
    for c in compare:
        if c.endswith('_return') or c.endswith('_SPY'):compare[c]=compare[c].map(lambda v:'%+.2f%%'%(100*v))
        elif c=='US10Y_change_bp':compare[c]=compare[c].map(lambda v:'%+.1f bp'%v)
        elif c=='US10Y_level_pct':compare[c]=compare[c].map(lambda v:'%.3f%%'%v)
    compare=compare.rename(columns={'IWM_minus_SPY':'IWM minus SPY (pp)','RSP_minus_SPY':'RSP minus SPY (pp)'})
    for c in ['IWM minus SPY (pp)','RSP minus SPY (pp)']:compare[c]=compare[c].str.replace('%','',regex=False)
    enc=base64.b64encode((out/'setup_event_comparison.svg').read_bytes()).decode('ascii')
    parts=['<!doctype html><html lang="en"><head><meta charset="utf-8"><title>IWM: 2025 setup vs now</title>',
       '<style>body{margin:0;background:#eef1f5;color:#17233d;font:16px/1.65 Arial}main{max-width:1240px;margin:auto;padding:40px;background:white}h1{font-size:40px;line-height:1.15}h2{margin-top:32px}a{color:#164aa0}table{border-collapse:collapse;font-size:12px;width:100%}td,th{padding:9px;border-bottom:1px solid #ddd;text-align:left}th{background:#e8edf6}img{width:100%}.meta{color:#67778b}.scroll{overflow-x:auto}</style></head><body><main>',
       '<p class="meta">IWM / TACTICAL ROTATION / AS OF '+str(t['Raw_Prices'].index[-1].date())+'</p>',
       '<h1>Is the late-2025 small-cap setup returning?</h1>',
       '<p><b>Short-horizon rotation is emerging. The historical analogy is partial, and October and November 2025 were different episodes.</b></p>',
       '<p><a href="rotation_research.xlsx">Detailed Excel workbook and editable charts</a> · <a href="full_report.html">Full correlation and historical research</a></p>',
       '<h2>Equal-length comparison: trailing 10 trading sessions</h2><p>Returns exclude cash dividends. Relative returns are percentage-point differences. All metrics in each row were observable at that row\'s date.</p>',
       '<div class="scroll">'+compare.to_html(index=False,escape=True)+'</div>',
       '<h2>Evidence and interpretation</h2>']
    parts+=['<p>'+html.escape(x)+'</p>' for x in findings]
    parts+=['<img alt="Historical and current event-aligned paths" src="data:image/svg+xml;base64,'+enc+'">',
            '<h2>Data and sources</h2><p>'+html.escape(config['source'])+'. Yahoo ^TNX and Bloomberg USGG10YR are different yield series/observation conventions: small level and basis-point differences are expected. '
            'Co-movement and event ordering do not prove that the AI sell-off caused Treasury buying or IWM gains.</p><ul>']
    for title,url,note in SOURCES:parts.append('<li><a href="'+html.escape(url)+'">'+html.escape(title)+'</a>: '+html.escape(note)+'</li>')
    parts+=['</ul></main></body></html>']
    (out/'report.html').write_text('\n'.join(parts),encoding='utf-8')
    (out/'setup_summary.txt').write_text('\n\n'.join(findings),encoding='utf-8')


def analyze(raw, config):
    p,coverage=clean_data(raw, config['end'])
    returns,corr,sma,rel,dd,sig=indicators(p)
    monthly,forward,stats,selected=conditional_study(p,sig,config['asof'])
    events,checkpoints=event_study(p,sig)
    latest=sig.iloc[-1].rename('value').to_frame().reset_index().rename(columns={'index':'metric'})
    latest['as_of']=sig.index[-1]
    c=corr.AAPL_NDX_63d.dropna()
    corr_audit=pd.DataFrame([dict(series=key,first_valid=s.dropna().index[0] if s.notna().any() else pd.NaT,
        last_valid=s.dropna().index[-1] if s.notna().any() else pd.NaT,
        latest=s.iloc[-1],minimum=s.min(),minimum_date=s.idxmin() if s.notna().any() else pd.NaT,
        negative_observations=int((s<0).sum()),first_negative=s[s<0].index[0] if (s<0).any() else pd.NaT)
        for key,s in corr.items()])
    rebased=p.loc[DISPLAY_START:,['IWM','SPY','QQQ','AAPL','XLV','XBI','KRE','RSP']].copy()
    for key in rebased:
        valid=rebased[key].dropna()
        if len(valid): rebased[key]=rebased[key]/valid.iloc[0]*100
    return add_setup_tables(dict(Raw_Prices=p,Daily_Returns=returns,Correlations=corr,SPX_100DMA=sma,
        Relative_Strength=rel,Momentum_Drawdown=dd,Daily_Signals=sig,Monthly_Signals=monthly,
        Forward_Outcomes=forward,Backtest_Statistics=stats,NonOverlap_Events=selected,
        Event_Windows=events,Checkpoints=checkpoints,Latest=latest,Coverage=coverage,
        Correlation_Audit=corr_audit,Rebased_Prices=rebased))

SOURCES = [
 ('First Trust October 10, 2025 review','https://www.ftportfolios.com/Commentary/Insights/2025/10/13/week-ended-october-10,-2025',
  'Tariff threats drove broad risk aversion and Treasury buying; this was not solely an AI event.'),
 ('Nasdaq October 2025 review','https://www.nasdaq.com/articles/october-2025-review-and-outlook',
  'Published 2025-11-03: October as a whole remained led by large-cap growth and technology.'),
 ('AP October 22, 2026 market report','https://apnews.com/article/33767691a4a0730b96d6f0c5100ea9b2',
  'Contemporaneous account of AI-stock declines, safe-haven Treasury buying and rising oil prices.'),
 ('Nasdaq November 2025 review','https://www.nasdaq.com/articles/november-2025-review-and-outlook',
  'Published 2025-12-01: contemporaneous account of technology weakness and rotation toward healthcare and regional banks.'),
 ('Janus Henderson December 2025 review','https://www.janushenderson.com/en-us/institutional/article/market-moves-themes-that-mattered-december-2025/',
  'Published 2026-01-05: December small-cap decline and higher long yields limit a simple year-end rally narrative.'),
 ('Federal Reserve December 2025 decision','https://www.federalreserve.gov/newsevents/pressreleases/monetary20251210a.htm',
  '2025-12-10: policy target cut by 25 bp to 3.50%-3.75%. A policy cut does not guarantee lower long yields.'),
 ('Nasdaq-100 launch history','https://www.nasdaq.com/articles/nasdaq-100-indexr-celebrates-40-years-innovation',
  'NDX launched in January 1985. Do not claim an observed NDX correlation history starting in 1981.'),
 ('MTUM fund','https://www.ishares.com/us/products/251614/ishares-msci-usa-momentum-factor-etf',
  'US momentum ETF; not the proprietary GS high-beta momentum basket.'),
 ('XLV fund','https://www.ssga.com/us/en/institutional/etfs/state-street-health-care-select-sector-spdr-etf-xlv',
  'Broad large-cap healthcare exposure; healthcare leadership need not imply small-cap risk appetite.'),
 ('XBI fund','https://www.ssga.com/us/en/individual/etfs/state-street-spdr-sp-biotech-etf-xbi',
  'Biotechnology exposure across capitalization sizes; a diagnostic, not a pure small-cap index.'),
 ('RSP fund','https://www.invesco.com/us/financial-products/etfs/product-detail?productId=RSP',
  'Equal-weight S&P 500 is a participation proxy, not a small-cap or fund-flow measure.'),
]
ROLES = {
 'IWM':('Core small-cap expression','Direct Russell 2000 exposure; test absolute returns and relative performance.'),
 'IJR':('Alternative small-cap expression','Compare another small-cap index methodology; not interchangeable with IWM.'),
 'AAPL':('Leadership diagnostic','Decoupling from NDX identifies dispersion; not proof of capital moving into IWM.'),
 'XLV':('Healthcare leadership','Can be defensive rotation; compare with XBI before calling it broad risk-on.'),
 'XBI':('Biotech risk appetite','Compare XBI/IBB and absolute returns; use as a satellite diagnostic.'),
 'KRE':('Domestic cyclicality diagnostic','Regional-bank strength can corroborate broader participation; sector-specific risks remain.'),
 'RSP':('Breadth diagnostic','RSP/SPY strength measures equal-weight relative performance, not constituent-level breadth.'),
 'MTUM':('Momentum unwind diagnostic','ETF factor exposure differs from a high-beta momentum basket.'),
 'SPMO':('Alternative momentum diagnostic','Useful robustness check against MTUM-specific holdings and rebalances.'),
 'NVDA':('AI leadership diagnostic','Compare with AAPL to distinguish company-specific resilience from broad tech strength.'),
 'MSFT':('Mega-cap leadership diagnostic','A second mega-cap comparison reduces reliance on a single AAPL anecdote.'),
 'LLY':('Healthcare concentration diagnostic','A large single-stock move can lift XLV without broad sector risk appetite.'),
 'UNH':('Healthcare dispersion diagnostic','Compare healthcare leaders and laggards rather than treating the sector as uniform.'),
 'HYG':('Imperfect credit-risk proxy','Price is affected by rates and distributions; it is NOT a credit-spread series.'),
 'US10Y':('Discount-rate regime','A lower yield helps only conditionally; recession-driven yield declines can hurt small caps.'),
}


def narrative(t, config):
    latest=t['Daily_Signals'].iloc[-1]
    corr=t['Correlations'].AAPL_NDX_63d.dropna()
    neg=corr[corr<0]
    date=str(t['Raw_Prices'].index[-1].date())
    audit=(('First negative observation in this downloaded 63-session PRICE-return sample: ' + str(neg.index[0].date()) + '.')
           if len(neg) else 'No negative observation in this downloaded 63-session sample.')
    lines=[
      'As of '+date+'. Source: '+config['source']+'. All baseline returns exclude cash dividends.',
      'Working thesis: easing long yields may support a broadening rally if earnings and credit conditions remain resilient. '
      'The evidence below measures price rotation; it does not establish actual net fund inflows or causality.',
      'AAPL / NDX: latest 63-session daily-return correlation %.3f; sample starts %s. %s '
      'The 21- and 126-session sensitivity series are also supplied. AAPL is included in NDX; this is not an ex-Apple index.'
      % (latest.AAPL_NDX_corr63, str(corr.index[0].date()), audit),
      'Current 21-session relative returns: IWM minus SPY %+.2f pp; RSP minus SPY %+.2f pp; '
      'XLV minus QQQ %+.2f pp; XBI minus IBB %+.2f pp. These are arithmetic return differences, not long-short portfolio returns.'
      % tuple(100*latest[x] for x in ['IWM_minus_SPY_21d','RSP_minus_SPY_21d','XLV_minus_QQQ_21d','XBI_minus_IBB_21d']),
      'US 10-year yield: %.3f%%, 21-session change %+.1f bp. SPX stands %+.2f%% above/below its 100-session moving average. '
      'A touch of the moving average is an observation, not an independently validated support signal.'
      % (latest.US10Y_level_pct,latest.US10Y_change_21d_bp,100*latest.SPX_distance_SMA100),
      'Trade framing: IWM expresses a small-cap rotation thesis; IJR is an alternative small-cap exposure. '
      'Use RSP, KRE and XBI/IBB as confirmation checks. AAPL, XLV, LLY and UNH help identify whether leadership is '
      'defensive or concentrated in a few names. They are research comparisons, not a recommended stock basket.',
    ]
    checkpoints=t['Checkpoints']
    pre=checkpoints[checkpoints.date==pd.Timestamp('2025-10-31')].set_index('metric')
    if len(pre):
        lines.append('Before the November 2025 sell-off, the October 31 checkpoint showed AAPL/NDX correlation '
            '%.3f (still positive), AAPL/XLV %.3f and XLV/NDX %.3f. XLV/NDX was at approximately the %.1fth '
            'percentile of its prior history. Healthcare decorrelation was more unusual than Apple/NDX decorrelation.' %
            (pre.loc['AAPL_NDX_corr63','value'],pre.loc['AAPL_XLV_corr63','value'],pre.loc['XLV_NDX_corr63','value'],
             100*pre.loc['XLV_NDX_corr63','prior_history_percentile']))
        lines.append('That same October 31 checkpoint did NOT yet show broad small-cap confirmation: '
            'RSP minus SPY was %+.2f pp over 21 sessions, KRE returned %+.2f%%, while XBI minus IBB was %+.2f pp. '
            'This was selective dispersion, not a clean all-market rotation signal.' %
            (100*pre.loc['RSP_minus_SPY_21d','value'],100*pre.loc['KRE_return_21d','value'],100*pre.loc['XBI_minus_IBB_21d','value']))
    sensitivity=t['Correlations'].iloc[-1]
    lines.append('Correlation-window sensitivity today: AAPL/NDX is %.3f over 21 sessions, %.3f over 63, '
        'and %.3f over 126. A claim such as first-ever negative correlation is window-, return-basis- and '
        'sample-dependent; it must not be generalized across definitions.' %
        (sensitivity.AAPL_NDX_21d,sensitivity.AAPL_NDX_63d,sensitivity.AAPL_NDX_126d))
    events=t['Event_Windows']
    nov=events[events.window=='2025_Nov'].set_index('series')
    if len(nov):
        lines.append('Other useful comparisons in November 2025: LLY %+.2f%% versus UNH %+.2f%%; '
            'NVDA %+.2f%% versus AAPL %+.2f%%. These contrasts help distinguish sector concentration and '
            'company-specific resilience from a broad small-cap rally; no holdings-weight attribution is assumed.' %
            tuple(100*nov.loc[k,'price_return'] for k in ['LLY','UNH','NVDA','AAPL']))
    for window in ['2023_Nov_Dec','2024_Nov','2024_Dec','2025_Nov_selloff','2025_Rebound','2025_Nov_Dec']:
        z=events[(events.window==window)&events.series.isin(['IWM','SPY','XLV','AAPL'])].set_index('series')
        if 'IWM' in z.index and 'SPY' in z.index:
            lines.append('%s: IWM %+.2f%%; SPY %+.2f%%; IWM minus SPY %+.2f pp; 10-year yield change %+.1f bp. %s.' %
                (window,100*z.loc['IWM','price_return'],100*z.loc['SPY','price_return'],
                 100*(z.loc['IWM','price_return']-z.loc['SPY','price_return']),z.loc['IWM','US10Y_change_bp'],z.loc['IWM','description']))
    stats=t['Backtest_Statistics']
    for rule in ['AAPL_negative','Healthcare_leadership','Breadth_plus_yield','Rotation_confirmation']:
        z=stats[(stats.rule==rule)&(stats.horizon==63)&(stats.sample_mode=='Non_overlapping')].iloc[0]
        if not z.n:
            lines.append(rule+': no matured non-overlapping 63-session outcome in the research era. No historical hit rate can be inferred.')
        else:
            lines.append('%s: %d non-overlapping historical episodes, mean next-63-session IWM return %+.2f%%, '
                         'mean IWM-minus-SPY %+.2f pp, relative win rate %.1f%%. Exploratory and gross of costs; '
                         'rules use different episode dates and have not been optimized or validated out of sample.' %
                         (rule,z.n,100*z.mean_IWM,100*z.mean_excess,100*z.excess_win_rate))
    if latest.IWM_minus_SPY_21d<=0:
        lines.append('Longer-window limitation: IWM has not outperformed SPY over the last 21 sessions. See Setup_Comparison for the potentially different 5/10-session direction. '
                     'Healthcare/Apple resilience alone is insufficient to claim an established small-cap rotation.')
    if latest.XBI_return_21d<0 and latest.XLV_return_21d>0:
        lines.append('Healthcare distinction: XLV is rising while XBI is falling over 21 sessions. '
                     'This looks more selective than a broad small-cap/biotech risk-on rally.')
    lines += [
      'Falsification checks: renewed long-yield increases; IWM/SPY relative trend deterioration; weakness spreading to '
      'regional banks and biotech; or sustained SPX weakness below its moving average. Confirm credit spreads and '
      'small-cap earnings revisions separately before translating a price pattern into a fundamental allocation.',
      'Research limitations: today\'s chosen instruments create selection bias; this is not a constituent-level breadth '
      'or flow backtest. Fixed thresholds were chosen for interpretability, not fit. Conditional averages are not '
      'statistical proof; overlapping outcomes are dependent and non-overlap samples remain regime-sensitive.',
    ]
    return lines


def unit_for(metric):
    if 'corr' in metric.lower(): return 'Correlation'
    if metric.endswith('_bp'): return 'Basis points'
    if metric=='US10Y_level_pct': return 'Yield percent (4.5 = 4.5%)'
    if metric=='VIX_level': return 'Index level'
    return 'Decimal return/spread (0.01 = 1% / 1 pp)'


def workbook(t, config, path, notes):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.chart import LineChart, Reference
    from openpyxl.chart.axis import DateAxis
    from openpyxl.workbook.properties import CalcProperties
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.utils import get_column_letter
    wb=openpyxl.Workbook();wb.remove(wb.active)
    wb.calculation=CalcProperties(calcId=191029, fullCalcOnLoad=True)
    universe=pd.DataFrame([dict(series=k,ticker=v,role=ROLES.get(k,('Benchmark/comparison',''))[0],
                         use_and_limit=ROLES.get(k,('','Context/benchmark series'))[1]) for k,v in TICKERS.items()])
    instructions=[
      ('How to use','Open Charts for editable Excel line charts. Each chart points to numeric worksheet cells; change the date range or style directly in Excel.'),
      ('How to replot','Select Date plus relevant series on a Chart_* sheet, then Insert > Line. Use a percentage axis for drawdowns/returns and a numeric axis for correlations.'),
      ('Raw data','Raw_Prices stores full downloaded observations on the SPX trading calendar. Daily_Returns stores simple returns; US10Y is daily change in bp.'),
      ('Cell values','All calculations are stored as numeric values for immediate plotting. Formula_Audit independently recalculates key indicators with native Excel formulas.'),
      ('Formula recalculation','Excel recalculates formulas on open. Formula_Audit has blank warm-up rows; cached values may be absent until Excel opens the workbook.'),
      ('Return basis','Split-adjusted prices, excluding cash dividends. Decimal 0.01 means 1% return or 1 percentage point excess return. This is not total return.'),
      ('Correlation','Pearson correlation of simple daily returns, 63 complete SPX sessions; 21/126-session alternatives included. No level correlation, filling or fabricated pre-inception history.'),
      ('Moving average','SPX 100-session arithmetic mean including today, with pre-display warm-up history. Cross-below markers use only observations available that day.'),
      ('Drawdown','Current price / cumulative peak since '+DRAWDOWN_START+' minus 1. This recovers to zero; cumulative worst drawdown is a different statistic.'),
      ('Signal execution','Completed month-end signal; next-session close entry; exit 21/63 sessions after entry. Unmatured outcomes stay blank, never counted as losses or wins.'),
      ('Rule: AAPL_decoupling','AAPL/NDX 63-session return correlation < 0.25. AAPL_negative uses < 0.'),
      ('Rule: Healthcare_leadership','XLV minus QQQ trailing 21-session return > 0.'),
      ('Rule: Breadth_improvement','RSP minus SPY trailing 21-session return > 0.'),
      ('Rule: Yield_relief','10-year yield change over 21 sessions < 0 bp.'),
      ('Rule: Breadth_plus_yield','Both Breadth_improvement and Yield_relief.'),
      ('Rule: Rotation_confirmation','Breadth_plus_yield, KRE positive trailing 21-session return, and SPX above its 100-session average. A descriptive screen, not a validated buy rule.'),
      ('Non-overlap','Within each rule/horizon, keep the earliest eligible signal, then only entries strictly after the previous exit. No forward outcomes are used to select dates.'),
      ('Baseline comparison','eligible_baseline_mean_excess uses all eligible month-ends for that rule. mean_excess_lift is descriptive; non-overlap signal dates differ from baseline dates.'),
      ('Daily tech screen','Since 2013: QQQ daily return <= -2%; yield relief <= -5 bp; resilience additionally requires IWM to beat SPY on the signal day. Next-close entry, 5/21/63-session holding periods. Historical events are not all AI-driven.'),
      ('Setup comparison','Current and historical dates use identical 1/5/10/21/63-session lookbacks. Event windows are retrospective illustrations; no future outcomes are used to construct the setup.'),
      ('Historical windows','November 20, 2025 and December 11, 2025 are retrospective illustration dates. The month-end conditional study is separate and does not use these dates as trade signals.'),
      ('Month completeness','A month is included only after its last weekday is observed. An exchange holiday on the last weekday may conservatively exclude that month; inspect Monthly_Signals.'),
      ('Missing data','Never forward filled. Coverage reports missing sessions and stale endpoints. Missing values invalidate full rolling windows and forward paths.'),
      ('News context','Historical commentary was reviewed on 2026-10-25. Price-based tables recalculate on each run; news context is not automatically refreshed.'),
      ('Source',config['source']),('Run as of',str(config['asof'].date())),
    ]
    readme=pd.DataFrame(instructions+ [('Finding %02d'%(i+1),n) for i,n in enumerate(notes)], columns=['Topic','Detail'])
    frames={'ReadMe':readme,'Universe':universe,'Coverage':t['Coverage']}
    frames.update(t)
    frames['Sources']=pd.DataFrame(SOURCES,columns=['Source','URL','Purpose'])
    # Dedicated compact chart sheets avoid forcing a forty-year x axis on recent plots.
    frames['Chart_Corr_Long']=t['Correlations'][['AAPL_NDX_63d']].dropna()
    frames['Chart_Corr_Recent']=t['Correlations'].loc[DISPLAY_START:,['AAPL_NDX_63d','AAPL_XLV_63d','XLV_NDX_63d']]
    frames['Chart_SPX_DMA']=t['SPX_100DMA'].loc[DISPLAY_START:,['SPX','SPX_SMA100']]
    frames['Chart_Relative']=t['Relative_Strength'].loc[DISPLAY_START:].copy()
    for c in frames['Chart_Relative']:
        z=frames['Chart_Relative'][c].dropna()
        if len(z):frames['Chart_Relative'][c]=frames['Chart_Relative'][c]/z.iloc[0]*100
    frames['Chart_2025']=t['Raw_Prices'].loc['2025-09-30':'2025-12-31',['IWM','SPY','AAPL','XLV','XBI','KRE','QQQ']].copy()
    for c in frames['Chart_2025']:
        z=frames['Chart_2025'][c].dropna()
        if len(z):frames['Chart_2025'][c]=frames['Chart_2025'][c]/z.iloc[0]*100
    for name,frame in frames.items():
        frame=frame.copy()
        if isinstance(frame.index,pd.DatetimeIndex):frame=frame.rename_axis('Date').reset_index()
        if name in ['Latest','Checkpoints']: frame['unit']=[unit_for(x) for x in frame.metric]
        ws=wb.create_sheet(name[:31]);ws.append(list(frame.columns))
        for row in frame.itertuples(index=False,name=None):
            clean=[]
            for val in row:
                if pd.isna(val): val=None
                elif isinstance(val,(pd.Timestamp,np.datetime64)):val=pd.Timestamp(val).to_pydatetime()
                elif isinstance(val,np.generic):val=val.item()
                clean.append(val)
            ws.append(clean)
        ws.freeze_panes='B2';ws.auto_filter.ref=ws.dimensions
        for cell in ws[1]:cell.font=Font(color='FFFFFF',bold=True);cell.fill=PatternFill('solid',fgColor='142652');cell.alignment=Alignment(wrap_text=True)
        ws.row_dimensions[1].height=32
        for j,c in enumerate(frame.columns,1):
            ws.column_dimensions[get_column_letter(j)].width=18 if j>1 else 22
            percent=(name=='Daily_Returns' and c not in ['Date','US10Y']) or 'drawdown' in c.lower() or c in ['price_return','distance_to_SMA100','SPX_distance_SMA100']
            percent=percent or ('21d' in c and not c.endswith('_bp') and name not in ['Correlations','Chart_Corr_Long','Chart_Corr_Recent']) or c in ['IWM_forward','SPY_forward','excess_forward']
            percent=percent or (name in ['Backtest_Statistics','Tech_Shock_Stats'] and any(c.startswith(x) for x in ['mean_','median_','min_excess','max_excess','IWM_win','excess_win','eligible_baseline']))
            percent = percent or c.endswith('_return') or c in ['IWM_minus_SPY','IWM_minus_QQQ','RSP_minus_SPY','XLV_minus_QQQ','IWM_day_minus_SPY']
            percent = percent or name in ['Chart_Event_IWM','Chart_Event_Excess'] and c!='session'
            for colcell in ws.iter_cols(min_col=j,max_col=j,min_row=2):
                for cell in colcell:
                    if isinstance(cell.value,datetime):cell.number_format='yyyy-mm-dd'
                    elif isinstance(cell.value,(float,int)) and not isinstance(cell.value,bool):
                        cell.number_format='0.00%;[Red](0.00%)' if percent else ('0.000' if 'corr' in c.lower() or name.startswith('Chart_Corr') or name=='Correlations' else '0.00')
        if name=='ReadMe':
            ws.column_dimensions['A'].width=28;ws.column_dimensions['B'].width=125
            for row in ws.iter_rows(min_row=2):
                row[1].alignment=Alignment(wrap_text=True,vertical='top');ws.row_dimensions[row[0].row].height=48
        if name in ['Universe','Sources']:
            ws.column_dimensions['C'].width=45
            if name=='Sources': ws.column_dimensions['B'].width=75
        if name in ['Backtest_Statistics','Tech_Shock_Stats']:
            ws.conditional_formatting.add('J2:J'+str(ws.max_row),ColorScaleRule(start_type='min',start_color='F2B7AE',mid_type='percentile',mid_value=50,mid_color='FFFFFF',end_type='max',end_color='A9D8C4'))
    audit=wb.create_sheet('Formula_Audit')
    audit.append(['Date','AAPL','NDX','SPX','IWM','SPY','US10Y','AAPL_return','NDX_return','Corr63_Excel','SPX_SMA100_Excel','US10Y_change_bp','Corr63_Python','SMA100_Python','Corr_difference','SMA_difference'])
    raw=t['Raw_Prices'].tail(420)
    for i,(date,row) in enumerate(raw.iterrows(),2):
        audit.append([date.to_pydatetime()]+[None if pd.isna(row[c]) else float(row[c]) for c in ['AAPL','NDX','SPX','IWM','SPY','US10Y']])
        audit.cell(i,1).number_format='yyyy-mm-dd'
        if i>=3:
            for col,source in [(8,'B'),(9,'C')]:audit.cell(i,col,f'=IF(AND(ISNUMBER({source}{i}),ISNUMBER({source}{i-1}),{source}{i-1}>0),{source}{i}/{source}{i-1}-1,"")')
            audit.cell(i,12,f'=IF(COUNT(G{i-1}:G{i})=2,(G{i}-G{i-1})*100,"")')
        if i>=65:
            audit.cell(i,10,f'=IF(AND(COUNT(H{i-62}:H{i})=63,COUNT(I{i-62}:I{i})=63),IFERROR(CORREL(H{i-62}:H{i},I{i-62}:I{i}),""),"")')
        if i>=101:audit.cell(i,11,f'=IF(COUNT(D{i-99}:D{i})=100,AVERAGE(D{i-99}:D{i}),"")')
        for col,value in [(13,t['Correlations'].loc[date,'AAPL_NDX_63d']),(14,t['SPX_100DMA'].loc[date,'SPX_SMA100'])]:
            if pd.notna(value):audit.cell(i,col,float(value))
        for col,a,b in [(15,'J','M'),(16,'K','N')]:audit.cell(i,col,f'=IF(AND(ISNUMBER({a}{i}),ISNUMBER({b}{i})),{a}{i}-{b}{i},"")')
        for j in [8,9]:audit.cell(i,j).number_format='0.00%'
    audit.freeze_panes='H2';audit.auto_filter.ref=audit.dimensions
    for col in range(1,17):audit.column_dimensions[get_column_letter(col)].width=22
    for cell in audit[1]:cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='142652')
    charts=wb.create_sheet('Charts',3)
    charts['A1']='Editable charts linked to workbook cells. All returns are price returns, excluding cash dividends.'
    charts.column_dimensions['A'].width=22
    chartspec=[('Chart_SPX_DMA',['SPX','SPX_SMA100'],'S&P 500 and 100-session moving average','Index level'),
      ('Chart_Corr_Long',['AAPL_NDX_63d'],'AAPL / NDX: 63-session daily-return correlation','Correlation'),
      ('Chart_Corr_Recent',['AAPL_NDX_63d','AAPL_XLV_63d','XLV_NDX_63d'],'Apple and healthcare: recent correlations','Correlation'),
      ('Chart_Relative',['IWM_SPY_rebased','RSP_SPY_rebased','KRE_SPY_rebased'],'Small-cap and breadth relative strength','Rebased = 100'),
      ('Momentum_Drawdown',['MTUM_drawdown','SPMO_drawdown','SPX_drawdown'],'Momentum ETFs vs SPX: underwater drawdown','Drawdown'),
      ('Chart_2025',['IWM','SPY','QQQ','AAPL','XLV','XBI','KRE'],'Q4 2025: distinguish sell-off from rebound','Rebased = 100')]
    colors=['080878','2020FF','159B8B','D28B3B','9755A6','6A7A8C','C64E55']
    chart_count=0
    for k,(sheet,cols,title,unit) in enumerate(chartspec):
        if wb[sheet].max_row < 3: continue
        chart_count+=1
        ws=wb[sheet];headers=[c.value for c in ws[1]]
        chart=LineChart();chart.title=title;chart.style=13;chart.width=28;chart.height=13
        chart.x_axis=DateAxis(axId=10,crossAx=100,axPos="b");chart.x_axis.number_format='mmm-yy';chart.x_axis.title='Date'
        chart.y_axis.title=unit;chart.legend.position='b';chart.display_blanks='gap'
        if unit=='Drawdown':chart.y_axis.numFmt='0%';chart.y_axis.scaling.max=0
        if unit=='Correlation':chart.y_axis.scaling.min=-1;chart.y_axis.scaling.max=1
        for j,c in enumerate(cols):
            col=headers.index(c)+1
            chart.add_data(Reference(ws,min_col=col,min_row=1,max_row=ws.max_row),titles_from_data=True)
            chart.series[-1].graphicalProperties.line.solidFill=colors[j%len(colors)]
            chart.series[-1].graphicalProperties.line.width=24000
        chart.set_categories(Reference(ws,min_col=1,min_row=2,max_row=ws.max_row))
        charts.add_chart(chart,'A'+str(3+k*27))
    for sheet,title,unit in [('Chart_Event_IWM','Event-aligned IWM performance','Return'),
        ('Chart_Event_Excess','Event-aligned IWM minus SPY','Return difference'),
        ('Chart_Event_Yield','Event-aligned Treasury yield change','Basis points')]:
        if sheet not in wb or wb[sheet].max_row<3:continue
        ws=wb[sheet];chart=LineChart();chart.title=title;chart.width=28;chart.height=13
        chart.x_axis.title='Trading sessions relative to anchor';chart.y_axis.title=unit
        chart.display_blanks='gap';chart.legend.position='b'
        if unit!='Basis points':chart.y_axis.numFmt='0%'
        chart.add_data(Reference(ws,min_col=2,max_col=ws.max_column,min_row=1,max_row=ws.max_row),titles_from_data=True)
        chart.set_categories(Reference(ws,min_col=1,min_row=2,max_row=ws.max_row))
        charts.add_chart(chart,'A'+str(3+chart_count*27));chart_count+=1
    wb.save(path)
    # Ensure the saved archive can be reopened and that chart links survived.
    check=openpyxl.load_workbook(path,read_only=False,data_only=False)
    if len(check['Charts']._charts)!=chart_count:raise AssertionError('Excel charts missing after save')
    if not str(check['Formula_Audit']['J65'].value).startswith('='):raise AssertionError('Audit formulas missing')
    check.close()


def plot_report(t, config, out):
    os.environ.setdefault('MPLCONFIGDIR',str(Path(tempfile.gettempdir())/'rotation_mpl'))
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from matplotlib.ticker import PercentFormatter
    plt.rcParams.update({'font.family':'DejaVu Sans','font.size':10,'svg.fonttype':'path'})
    pictures=[]
    def chart(name,frame,title,unit,percent=False,zero=False,markers=None):
        fig,ax=plt.subplots(figsize=(12,5.4))
        colors=['#080878','#2020ff','#109b85','#cc8c2f','#a05bba','#c65454','#6c8191']
        for i,col in enumerate(frame):
            s=frame[col]
            ax.plot(s.index,s,lw=1.6,color=colors[i%len(colors)],label=col.replace('_',' '))
        if markers is not None:
            ix=markers[markers].index.intersection(frame.index)
            ax.scatter(ix,frame.loc[ix].iloc[:,0],s=50,facecolors='none',edgecolors='#080878',zorder=4)
        if zero:ax.axhline(0,color='#8490ab',lw=.7,ls='--')
        ax.yaxis.tick_right();ax.yaxis.set_label_position('right');ax.set_ylabel(unit)
        if percent:ax.yaxis.set_major_formatter(PercentFormatter(1))
        ax.set_title(title,loc='left',pad=16,weight='bold')
        loc=mdates.AutoDateLocator(minticks=4,maxticks=8);ax.xaxis.set_major_locator(loc)
        ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(loc))
        ax.grid(axis='y',alpha=.15);ax.legend(loc='upper left',bbox_to_anchor=(0,-.12),ncol=min(4,len(frame)),frameon=False,fontsize=8)
        fig.subplots_adjust(left=.035,right=.91,top=.87,bottom=.25)
        fig.text(.035,.02,'Source: '+config['source']+' | Price returns; cash dividends excluded. Latest: '+str(t['Raw_Prices'].index[-1].date()),fontsize=8,color='#666')
        fig.savefig(out/(name+'.png'),dpi=170,facecolor='white')
        fig.savefig(out/(name+'.svg'),facecolor='white')
        plt.close(fig);pictures.append((title,out/(name+'.svg')))
    chart('01_spx_100dma',t['SPX_100DMA'].loc[DISPLAY_START:,['SPX','SPX_SMA100']],
          'S&P 500 and its 100-session moving average','Index level',markers=t['SPX_100DMA'].cross_below)
    chart('02_aapl_ndx_correlation',t['Correlations'][['AAPL_NDX_63d']].dropna(),
          'Apple / Nasdaq-100: 63-session realized return correlation','Correlation',zero=True)
    chart('03_healthcare_correlation',t['Correlations'].loc[DISPLAY_START:,['AAPL_NDX_63d','AAPL_XLV_63d','XLV_NDX_63d']],
          'Apple and healthcare: leadership dispersion','Correlation',zero=True)
    rel=t['Relative_Strength'].loc[DISPLAY_START:,['IWM_SPY_rebased','RSP_SPY_rebased','KRE_SPY_rebased']].copy()
    for c in rel:
        z=rel[c].dropna()
        if len(z):rel[c]=rel[c]/z.iloc[0]*100
    chart('04_relative_strength',rel,'Does the rotation reach small caps?','Relative price, start = 100')
    chart('05_momentum_drawdown',t['Momentum_Drawdown'],'Momentum ETF drawdowns vs S&P 500','Drawdown',percent=True)
    ep=t['Raw_Prices'].loc['2025-09-30':'2025-12-31',['IWM','SPY','QQQ','AAPL','XLV','XBI','KRE']].copy()
    if len(ep)>1:
        for c in ep:ep[c]=ep[c]/ep[c].dropna().iloc[0]*100
        chart('06_2025_rotation',ep,'Q4 2025: sell-off, rebound, then consolidation','Price, start = 100')
    return pictures


def save_report(t, config, out):
    out=Path(out).expanduser().resolve()
    if out.exists() and any(out.iterdir()):out=out/('run_'+pd.Timestamp.now().strftime('%Y%m%d_%H%M%S_%f'))
    out.mkdir(parents=True,exist_ok=True)
    notes=setup_findings(t)+narrative(t,config)
    stale=t['Coverage'].query('stale_calendar_days > 7')
    if len(stale):notes.insert(0,'DATA WARNING: stale series: '+', '.join(stale.series)+'. Inspect Coverage before using current conclusions.')
    LOG.info('Writing detailed Excel workbook and native charts')
    workbook(t,config,out/'rotation_research.xlsx',notes)
    for name,frame in t.items():frame.to_csv(out/(name+'.csv'),index=isinstance(frame.index,pd.DatetimeIndex),index_label='Date')
    pictures=plot_report(t,config,out)
    evidence=t['Backtest_Statistics'].query('horizon == 63 and sample_mode == "Non_overlapping"')[['rule','n','mean_IWM','mean_excess','excess_win_rate']].copy()
    for c in ['mean_IWM','mean_excess','excess_win_rate']:evidence[c]=evidence[c].map(lambda v:'n/a' if pd.isna(v) else '%.2f%%'%(100*v))
    parts=['<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Small-cap rotation research</title>',
      '<style>body{font:16px/1.6 Arial,sans-serif;color:#17233d;background:#f7f8fb;margin:0}main{max-width:1180px;margin:0 auto;padding:44px 28px;background:white}h1{font-size:38px;line-height:1.1}h2{margin-top:38px;color:#142652}p{max-width:1050px}a{color:#2020bb}img{width:100%;margin-top:15px}table{border-collapse:collapse;font-size:13px;width:100%}td,th{text-align:left;padding:10px;border-bottom:1px solid #ddd}th{background:#eef1f8}.meta{color:#62718b}.lead{font-size:20px;font-weight:bold}li{margin:9px 0}</style></head><body><main>',
      '<p class="meta">EQUITY STRATEGY / REPRODUCIBLE RESEARCH</p><h1>Rotation is visible.<br>Small-cap confirmation requires more.</h1>',
      '<p class="lead">Test the long-small-cap argument against Apple, healthcare, market breadth and long yields.</p>',
      '<p><a href="rotation_research.xlsx">Open the detailed Excel workbook with editable charts</a></p>',
      '<h2>Findings and implications</h2>']
    parts+=['<p>'+html.escape(n)+'</p>' for n in notes]
    parts+=['<h2>Historical conditional outcomes</h2><p>Next 63 trading sessions, entry at the next-session close; non-overlapping episodes within each rule. '
            'Excess return = IWM return minus SPY return. Low sample sizes and different event dates limit comparisons.</p>',evidence.to_html(index=False,escape=True)]
    for title,path in pictures:
        enc=base64.b64encode(path.read_bytes()).decode('ascii')
        parts.append('<h2>'+html.escape(title)+'</h2><img alt="'+html.escape(title)+'" src="data:image/svg+xml;base64,'+enc+'">')
    parts+=['<h2>Sources and historical context</h2><ul>']
    for label,url,purpose in SOURCES:parts.append('<li><a href="'+html.escape(url)+'">'+html.escape(label)+'</a>: '+html.escape(purpose)+'</li>')
    parts+=['</ul><h2>Methodology</h2><p>'+html.escape(__doc__).replace('\n','<br>')+'</p></main></body></html>']
    (out/'full_report.html').write_text('\n'.join(parts),encoding='utf-8')
    setup_report(t,config,out)
    (out/'research_note.txt').write_text('\n\n'.join(notes)+'\n\n'+__doc__,encoding='utf-8')
    return out


def self_test():
    idx=pd.bdate_range('2020-01-01',periods=1500)
    rng=np.random.default_rng(20261025)
    p=pd.DataFrame({key:100*np.exp(np.cumsum(rng.normal(.0002,.008,len(idx)))) for key in TICKERS},index=idx)
    p['US10Y']=4+np.cumsum(rng.normal(0,.01,len(idx)))
    cfg=dict(end=idx[-1],asof=idx[-1]+pd.Timedelta(days=1),source='Synthetic test')
    t=analyze(p,cfg)
    last=t['Correlations'].AAPL_NDX_63d.iloc[-1]
    r=p[['AAPL','NDX']].pct_change(fill_method=None).iloc[-63:]
    np.testing.assert_allclose(last,np.corrcoef(r.AAPL,r.NDX)[0,1],atol=1e-12)
    np.testing.assert_allclose(t['SPX_100DMA'].SPX_SMA100.iloc[-1],p.SPX.iloc[-100:].mean())
    np.testing.assert_allclose(t['Daily_Returns'].US10Y.iloc[-1],(p.US10Y.iloc[-1]-p.US10Y.iloc[-2])*100)
    forward=t['Forward_Outcomes'].dropna(subset=['IWM_forward'])
    for _,row in forward.iterrows():
        sigpos=p.index.get_loc(row.signal_date);entrypos=p.index.get_loc(row.entry_date);exitpos=p.index.get_loc(row.exit_date)
        assert entrypos==sigpos+1 and exitpos==entrypos+row.horizon
        np.testing.assert_allclose(row.IWM_forward,p.IWM.iloc[exitpos]/p.IWM.iloc[entrypos]-1)
    for _,group in t['NonOverlap_Events'].groupby(['rule','horizon']):
        dates=group.sort_values('entry_date')
        assert (dates.entry_date.iloc[1:].to_numpy()>dates.exit_date.iloc[:-1].to_numpy()).all()
    # Changes to future prices must not alter past indicators or signals.
    revised=p.copy();revised.iloc[-100:]*=1.2
    ta=indicators(p);tb=indicators(revised)
    pd.testing.assert_frame_equal(ta[-1].iloc[:-100],tb[-1].iloc[:-100])
    bad=p.copy();bad.loc[idx[-10],'AAPL']=np.nan
    assert pd.isna(indicators(bad)[1].AAPL_NDX_63d.iloc[-1])
    # Index-only holiday prints do not create artificial equity zero-return sessions.
    holiday=idx[-1]+pd.Timedelta(days=1)
    raw=pd.concat([p,pd.DataFrame({'VIX':[20.]},index=[holiday])])
    cleaned,_=clean_data(raw,holiday)
    pd.testing.assert_frame_equal(cleaned.rename_axis(None),p.rename_axis(None),check_freq=False)
    for invalid in [pd.concat([p,p.iloc[:1]]),p.assign(IWM=0),p.assign(AAPL=np.inf)]:
        try:clean_data(invalid,idx[-1])
        except ValueError:pass
        else:raise AssertionError('Bad data accepted')
    historical=p.loc[:'2024-11-20']
    hist_sig=indicators(historical)[-1]
    hist_months=conditional_study(historical,hist_sig,pd.Timestamp('2026-10-25'))[0]
    assert hist_months.index[-1].month==10  # an old cutoff must not masquerade as a full month
    shock=p.copy()
    shock['QQQ']=100+np.arange(len(idx))*.01
    shock['US10Y']=4+np.arange(len(idx))*.0001
    for loc in [-100,-20,-2]:
        j=len(idx)+loc
        shock.loc[idx[j],'QQQ']=shock.QQQ.iloc[j-1]*.97
        shock.loc[idx[j],'US10Y']=shock.US10Y.iloc[j-1]-.10
        shock.loc[idx[j],'IWM']=shock.IWM.iloc[j-1]
        shock.loc[idx[j],'SPY']=shock.SPY.iloc[j-1]*.99
    st,ev,sel=tech_shock_study(shock)
    recent=ev[ev.signal_date==idx[-2]]
    assert len(recent)==9 and recent.IWM_forward.isna().all()
    for _,row in ev.dropna(subset=['IWM_forward']).iterrows():
        i=shock.index.get_loc(row.signal_date)
        assert row.entry_date==idx[i+1] and row.exit_date==idx[i+1+row.horizon]
        np.testing.assert_allclose(row.IWM_forward,shock.IWM.iloc[i+1+row.horizon]/shock.IWM.iloc[i+1]-1)
    for _,group in sel.groupby(['rule','horizon']):
        q=group.sort_values('entry_date')
        assert (q.entry_date.iloc[1:].to_numpy()>q.exit_date.iloc[:-1].to_numpy()).all()
    setup=t['Setup_Comparison'].query('case == "Current" and lookback_sessions == 5').iloc[0]
    np.testing.assert_allclose(setup.IWM_return,p.IWM.iloc[-1]/p.IWM.iloc[-6]-1)
    print('PASS: correlation, SMA, yield units, next-close execution, forward windows, non-overlap, no-lookahead, missing sessions, invalid data, setup lookbacks and unmatured shock outcomes.')


def main():
    parser=argparse.ArgumentParser(description=__doc__,formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--backend',choices=['auto','xbbg','pdblp','blpapi'],default=BACKEND)
    parser.add_argument('--end',default=END_DATE)
    parser.add_argument('--out',default=OUTPUT_DIR)
    parser.add_argument('--self-test',action='store_true')
    args=parser.parse_args()
    if args.self_test:self_test();return
    # Preflight dependencies before making Bloomberg calls.
    import openpyxl
    import matplotlib
    try:
        from zoneinfo import ZoneInfo
        today=datetime.now(ZoneInfo('America/New_York')).date()
    except Exception:
        # Fixed UTC-5 fallback is conservative around the DST midnight boundary.
        today=datetime.now(timezone(timedelta(hours=-5))).date()
    asof=pd.Timestamp(today);latest=asof-pd.Timedelta(days=1)
    end=min(pd.Timestamp(args.end).normalize(),latest) if args.end else latest
    if args.end and pd.Timestamp(args.end)>latest:LOG.warning('Requested end clipped to yesterday: %s',end.date())
    if end<=pd.Timestamp(SIGNAL_START):raise ValueError('End date must be after SIGNAL_START')
    cfg=dict(start=pd.Timestamp(START_DATE),end=end,asof=asof,tickers=dict(TICKERS),source='',warnings=[])
    raw=pull_bloomberg(cfg,args.backend)
    t=analyze(raw,cfg)
    out=save_report(t,cfg,args.out or Path(__file__).resolve().parent.parent/'iwm_results')
    print('Excel: '+str(out/'rotation_research.xlsx'))
    print('HTML:  '+str(out/'report.html'))


if __name__=='__main__':
    logging.basicConfig(level=logging.INFO,format='%(levelname)s: %(message)s')
    try:main()
    except Exception as exc:
        LOG.error('%s: %s',type(exc).__name__,exc)
        raise SystemExit(1)
