# BLS AI and labor demand: reproduce the historical vs projected employment chart

Updated October 2026. Source for both the 2025 base and the 2035 forecast: BLS Employment Projections, 2025–2035.

## 1. Download data

1. **BLS Table 2.1 — Employment by major industry sector:** https://www.bls.gov/emp/tables/employment-by-major-industry-sector.htm . Click **Other available formats (XLSX)**. Save as `BLS_Table_2_1.xlsx`. This file has industry, 2015 actual, 2025 base, 2035 projection, and 2015–25 and 2025–35 CAGR.
2. **BLS historical annual industry employment:** https://www.bls.gov/emp/data/industry-out-and-emp.htm . Under **Data for Researchers** click **Industry output and employment data for the U.S. economy for historical years, and for the projected year (ZIP)**. Download link: https://www.bls.gov/emp/industry-employment/industry.zip . Extract and identify the *historical industry employment* series using the included industry description/sectoring plan. Use annual 2015–2025 values, same industry identifier as Table 2.1. Do not confuse industry codes in the researcher data with 2-digit NAICS; check the sectoring plan.
3. **Documentation on AI in BLS projections:** https://www.bls.gov/emp/factsheets/ai-impacts.htm . AI is one of several modeled labor-demand factors, but BLS forecasts are not AI-only counterfactuals.
4. **Direct occupational example:** https://www.bls.gov/opub/ted/2026/total-employment-projected-to-grow-3-5-percent-between-2025-and-2035.htm . The BLS projects office/administrative support employment down 4.0% (-752,100 jobs) over 2025–35 and explicitly mentions AI-enabled automation among contributing factors. Note: SOC 43 occupational category is NOT NAICS 56 administrative-services industry.
5. **Optional production/output link:** https://www.bls.gov/emp/tables/industry-employment-and-output.htm . Table 2.11 has 2025/2035 output and employment for each industry, useful for illustrating output growth vs headcount growth.

## 2. Pick your example industry

Finance and insurance (NAICS 52), values in **thousands of jobs** from BLS Table 2.1:

| Historical 2015 | Base 2025 | Forecast 2035 | Historical CAGR | Forward CAGR |
|---:|---:|---:|---:|---:|
| 6,034.7 | 6,739.3 | 6,893.6 | `(6739.3/6034.7)^(1/10)-1` | `(6893.6/6739.3)^(1/10)-1` |

Other examples: Professional/scientific/technical (NAICS 54): 8,658.4 → 10,779.2 → 11,705.9; Manufacturing (31–33): 12,309.8 → 12,633.7 → 12,705.8. All from the BLS Table 2.1.

## 3. Fastest chart (uses only BLS Table 2.1)

Create a new Excel sheet `Employment Trend`. Type:

| A | B | C |
|---|---|---|
| Year | Historical (actual anchor) | BLS projection (anchor) |
| 2015 | 6034.7 | *(blank)* |
| 2025 | 6739.3 | 6739.3 |
| 2035 | *(blank)* | 6893.6 |

Select `A1:C4` -> **Insert -> Line with Markers**. Format **Historical** as solid and **BLS projection** as dashed (right-click the projected line -> Format Data Series -> Line -> Dash type). IMPORTANT: each line connects two anchor data points. The 2015–2025 portion is NOT an observed annual sequence in this quick view; it is a connector between observed endpoints.

## 4. Full historical annual chart (BLS research ZIP)

Set up the sheet:

- H2 = 2015 employment from Table 2.1 (`6034.7`)
- H3 = 2025 employment from Table 2.1 (`6739.3`)
- H4 = 2035 projected employment from Table 2.1 (`6893.6`)
- H5 = `=(H4/H3)^(1/10)-1` (BLS ten-year CAGR, derived)
- A2:A22 = years 2015 through 2035, one year per row
- B2:B12 = annual 2015–2025 historical employment from the official BLS research ZIP (same industry definition). **Do not invent missing annual historical values or interpolate them.**
- B13:B22 = blank
- C2:C11 = blank
- C12 = `=$H$3` (duplicate 2025 anchor, to join the line)
- C13 = `=C12*(1+$H$5)`; drag to C22 (2035)
- C22 should be approximately H4, allowing for rounding.

Select A1:C22 -> Insert -> Line chart. Make **B Historical** solid and **C Forecast** dashed. Title: `Finance & Insurance Employment: Historical 2015–25 vs BLS Projected 2025–35`. Y-axis: employment, thousands. Footnote: `Sources: U.S. BLS Employment Projections (2025–2035), Table 2.1 and Industry Researcher Data. 2026–2034 forecast annual values interpolated from BLS 2025/2035 endpoints at constant CAGR.`

## 5. Optional output/job comparison

For Finance and insurance, Table 2.11 gives 2025 output = 3,285.1 billion chained-2017 dollars and 2035 output = 3,915.2, and employment = 6,739.3 thousand vs 6,893.6 thousand.

- **Jobs growth:** `=6893.6/6739.3-1`
- **Real output growth:** `=3915.2/3285.1-1`
- **Change in labor intensity:** `=(6893.6/3915.2)/(6739.3/3285.1)-1`

This suggests fewer jobs per unit of real output, a labor-productivity narrative, but it does **not** identify AI's causal contribution.

## 6. Interpretation for your report

Safe: `BLS expects a materially slower rate of employment growth in selected industries, while explicitly incorporating AI-related productivity and displacement in qualitative occupational assessments.`

Avoid: `AI alone will eliminate the gap between historical trend and the BLS forecast.` BLS provides no AI-free employment counterfactual. Many factors, including demographics and demand, affect employment. Across the whole economy BLS projects **job growth**, not an absolute AI-driven reduction in total jobs.
