"""Additional stock-level conviction, disagreement and positioning-change factors."""

import numpy as np
import pandas as pd
import pyarrow.parquet as pq


def add_manager_signals(
    portfolio_path,
    stock_summary,
    neutral_band_pct=0.1,
    conviction_cutoff=0.90,
    stock_id_column="security_id",
    aum_column="fund_aum",
):
    """Rank absolute active weights within each fund/month, then summarize by stock."""
    keys = ["analysis_date", stock_id_column]
    fund_keys = ["analysis_date", "factset_fund_id"]
    columns = keys + [
        "factset_fund_id", "portfolio_weight_pct", "benchmark_weight_pct",
        "active_value", aum_column,
    ]
    new_columns = [
        "high_conviction_overweight_pct", "mean_conviction_rank_pct",
        "active_weight_std_pct", "active_weight_iqr_pct",
        "median_conviction_pool_count",
    ]

    available = set(pq.ParquetFile(portfolio_path).schema_arrow.names)
    if missing := set(columns) - available:
        raise ValueError(f"Missing fund_stock columns: {sorted(missing)}")
    if not 0 < conviction_cutoff < 1:
        raise ValueError("conviction_cutoff must be between 0 and 1")

    summary = stock_summary.copy()
    summary["analysis_date"] = pd.to_datetime(summary["analysis_date"])
    if summary.duplicated(keys).any():
        raise ValueError("Duplicate stock/date rows in stock_summary")
    if set(new_columns) & set(summary.columns):
        raise ValueError("Manager signal columns already exist in stock_summary")

    dates = pq.read_table(portfolio_path, columns=["analysis_date"]).to_pandas()
    dates = pd.to_datetime(dates["analysis_date"]).drop_duplicates().sort_values()
    monthly = []

    for date in dates:
        frame = pq.read_table(
            portfolio_path, columns=columns, filters=[("analysis_date", "=", date)]
        ).to_pandas()
        if frame.empty:
            continue
        frame["analysis_date"] = pd.to_datetime(frame["analysis_date"])
        if frame.duplicated(fund_keys + [stock_id_column]).any():
            raise ValueError(f"Duplicate fund/stock records for {date.date()}")

        for column in ["portfolio_weight_pct", "benchmark_weight_pct", "active_value", aum_column]:
            frame[column] = pd.to_numeric(frame[column], errors="raise")
        valid = (
            frame["portfolio_weight_pct"].notna()
            & frame["benchmark_weight_pct"].notna()
            & frame["active_value"].notna()
            & frame[aum_column].gt(0)
        )
        active_pct = 100 * frame["active_value"] / frame[aum_column].where(frame[aum_column].gt(0))
        valid &= np.isfinite(active_pct)
        frame["active_pct"] = active_pct.where(valid)

        # Pool: all valid stocks in the same fund/month, including benchmark-only rows.
        by_fund = frame.groupby(fund_keys, sort=False)
        frame["conviction_rank"] = by_fund["active_pct"].transform(
            lambda x: x.abs().rank(pct=True)
        )
        frame["conviction_pool_count"] = by_fund["active_pct"].transform("count").where(valid)
        frame["high_conviction_overweight"] = (
            valid & frame["conviction_rank"].gt(conviction_cutoff)
            & frame["active_pct"].gt(neutral_band_pct)
        )

        grouped = frame.groupby(keys, dropna=False, sort=False)
        extra = grouped.agg(
            valid_fund_count=("active_pct", "count"),
            high_conviction_count=("high_conviction_overweight", "sum"),
            mean_conviction_rank=("conviction_rank", "mean"),
            active_weight_std_pct=("active_pct", "std"),
            median_conviction_pool_count=("conviction_pool_count", "median"),
        )
        extra["active_weight_iqr_pct"] = (
            grouped["active_pct"].quantile(0.75)
            - grouped["active_pct"].quantile(0.25)
        ).where(extra["valid_fund_count"].ge(2))
        extra["high_conviction_overweight_pct"] = (
            100 * extra["high_conviction_count"]
            / extra["valid_fund_count"].where(extra["valid_fund_count"].gt(0))
        )
        extra["mean_conviction_rank_pct"] = 100 * extra["mean_conviction_rank"]
        monthly.append(extra.reset_index()[keys + new_columns])

    if not monthly:
        return summary.assign(**{column: np.nan for column in new_columns})
    extra = pd.concat(monthly, ignore_index=True)
    if extra.duplicated(keys).any():
        raise ValueError("Duplicate stock/date rows in manager signal output")
    return summary.merge(extra, on=keys, how="left", validate="one_to_one")


def add_positioning_changes(stock_summary, stock_id_column="security_id"):
    """Calculate adjacent-month stock positioning changes in percentage points."""
    keys = ["analysis_date", stock_id_column]
    metrics = ["aum_weighted_active_weight_pct", "overweight_pct"]
    if missing := set(keys + metrics) - set(stock_summary.columns):
        raise ValueError(f"Missing stock sentiment columns: {sorted(missing)}")

    result = stock_summary.copy()
    result["analysis_date"] = pd.to_datetime(result["analysis_date"])
    if result.duplicated(keys).any():
        raise ValueError("Duplicate stock/date rows before positioning changes")
    result = result.sort_values([stock_id_column, "analysis_date"])
    by_stock = result.groupby(stock_id_column, sort=False)
    prior_date = by_stock["analysis_date"].shift(1)
    consecutive = result["analysis_date"].eq(prior_date + pd.offsets.MonthEnd(1))

    # A missing month must not be treated as a one-month change.
    result["active_weight_change_1m_pp"] = (
        result[metrics[0]] - by_stock[metrics[0]].shift(1)
    ).where(consecutive)
    result["overweight_pct_change_1m_pp"] = (
        result[metrics[1]] - by_stock[metrics[1]].shift(1)
    ).where(consecutive)
    return result
