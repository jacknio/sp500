请基于现有 Mutual Fund Positioning Project，增加三个 Stock-level Sentiment Signals。

**要求：**
- 先检查现有 `main.py`、`sentiment.py` 的实际代码和数据结构。
- 尽量不要修改现有代码，尤其不要改 `pull_data.py`、`cleandata.py` 和原有计算逻辑。
- 新建一个简单的 `sentiment_signals.py`，把新增计算写成独立 Functions。
- 在 `main.py` 现有 Sentiment Aggregation 完成后调用，输出新的 `aggregate_stock_extended.parquet`，保留原始 Output。
- 代码简单易读，少量必要的英文 Comments，不要引入复杂 Class 或额外模块。

### 1. Manager Conviction

对于每个 Fund × Month：

- 使用全部有效的 Fund × Stock 记录作为 Ranking Pool，包括 Benchmark-only Stocks。
- 按 `abs(active_weight_pct)` 排名，识别 Top 10% 的主动偏离仓位。
- 正向 Overweight 才计入 High-conviction Overweight。
- 对每只 Stock 汇总 `high_conviction_overweight_pct`。
- 同时保留每只基金的 Ranking Pool Size，方便检查样本覆盖。

### 2. Manager Disagreement

对于每个 Stock × Month：

- 计算所有有效基金的 Active Weight Standard Deviation。
- 计算 Active Weight IQR（P75 − P25）。
- 输出 `active_weight_std_pct`、`active_weight_iqr_pct`。
- 基金等权，缺失值不能当成零。

### 3. Positioning Change

比较同一只 Stock 连续两个月的 Aggregate Sentiment：

- `active_weight_change_1m_pp`：AUM-weighted Active Weight 的月度变化。
- `overweight_pct_change_1m_pp`：Overweight Fund Percentage 的月度变化。
- 第一月或缺失相邻月份时保留 NaN。
- 不要把两期样本变化直接解释为基金实际买卖行为。

### 4. 验证与交付

请先确认现有字段名称、计算单位、有效基金定义，以及每月基金股票记录的唯一性，**不要直接假设我之前提供的代码正确**。

新增 Signals 必须与原 `sentiment.py` 使用相同的 Valid Contributor 标准和 Stock Identifier。

完成后：

1. 告诉我新文件放在哪里。
2. 告诉我 `main.py` 具体在哪一行插入 Import 和 Function Calls。
3. 给出两个关键 Debug Breakpoints，以及应检查的 DataFrame 和 Columns。
4. 用少量真实数据验证 Conviction 排名、Std、IQR、Monthly Changes。
5. 如果发现不确定的数据口径或字段问题，先说明，不要自行假设或大幅修改项目。

**核心原则：只增加这三个研究方向，最小化代码改动，确保代码容易人工 Review 和 Debug。**
