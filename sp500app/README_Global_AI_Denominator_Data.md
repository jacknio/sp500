# Global AI Denominator — Data Download & Replication README

**Project:** AI Monetization / AI-Exposed Wage Pool / CapEx Revenue Hurdle  
**Reference date:** 2026-10-22  
**Target year:** 2030  
**Currency:** nominal USD, unless otherwise stated  
**Purpose:** Build auditable U.S., Global, and Global Ex-China denominators. Keep all downloaded source files unchanged and put calculations in separate worksheets.

> **中文使用说明**：按本文的 A → B → C 顺序做。第一轮只需下载 IMF GDP、ILOSTAT Labor Share 和 ILO GitHub Exposure 三组数据，就可以做出 Global / Ex-China 的初版表。World Bank 行业权重和 ILO 行业×职业就业数据是第二轮验证，不必第一天就全部做完。
>
> **重要**：本 README 的 IMF Excel **三个 sheet 名称**已经与 IMF 官方下载页核对，但由于无法在当前环境中读取原始 xlsx 的内部单元格，本文件不假定固定列号或行号。下载后以表格中的实际字段名为准，不要凭空使用 A、B、C 或第 2030 列的位置。

---

## 0. What you are trying to calculate / 最终要得到什么

| Output | Definition | Model role |
|---|---|---|
| `US_GDP_2030` | U.S. nominal GDP in 2030 | GDP denominator / baseline |
| `World_GDP_2030` | World nominal GDP in 2030 | Global GDP denominator |
| `China_GDP_2030` | China's nominal GDP in 2030 | Used for exclusion |
| `World_ExChina_GDP_2030` | World GDP − China GDP | Global Ex-China denominator |
| `US_AI_Exposed_Wages_2030` | Existing U.S. exposed-wage model, projected consistently to 2030 | Wage denominator baseline |
| `World_AI_Exposed_Wages_2030` | U.S. wage pool × GDP ratio × labor-share ratio × relative wage-weighted exposure | Scenario, **not** direct official forecast |
| `World_ExChina_AI_Exposed_Wages_2030` | Same adjustment using Ex-China GDP and Ex-China labor share/exposure | Scenario, **not** direct official forecast |

**不要混淆：** nominal GDP、AI-exposed wages、AI vendor revenue、labor savings 分别是不同概念；不得直接相加。`AI exposure` 不是 actual productivity improvement，更不是已经实现的 cost reduction。

### Suggested local folders / 推荐保存路径

```text
Global_AI_Denominator/
├── README_Global_AI_Denominator_Data.md
├── data/
│   └── raw/
│       ├── IMF/
│       │   └── WEOOct2026all.xlsx
│       ├── ILOSTAT/
│       │   ├── labor_income_share_download.xlsx_or_csv
│       │   └── employment_by_industry_occupation_optional.csv
│       ├── ILO_AI_Exposure/
│       │   ├── Final_Scores_ISCO08_Gmyrek_et_al_2025.xlsx
│       │   ├── 4digits_with_tasks.xlsx
│       │   └── DATA_DICTIONARY.md
│       └── World_Bank/
│           └── industry_value_added_optional.csv
└── model/
    └── Global_Denominator_Model.xlsx  # 自己在公司电脑建；不要覆盖 raw
```

---

# A. IMF World Economic Outlook — 2030 Nominal GDP (PRIORITY 1)

## A1. Exact official download links / 官网从哪里进

**IMF WEO landing page:**  
https://data.imf.org/Datasets/WEO

**Direct Excel download (October 2026 WEO Entire Dataset):**  
https://data.imf.org/-/media/iData/External-Storage/Documents/0FA3DC49A04D446E9D69DBEB29979E1C/en/WEOOct2026all.xlsx

**WEO database FAQ:**  
https://data.imf.org/Datasets/WEO/Frequently-Asked-Questions

**IMF file-navigation / transition guide:**  
https://data.imf.org/-/media/iData/External-Storage/Documents/2C2B2F3E671A4B11ABE756E2441FD6F0/en/WEO-Database-Transition-Guide.pdf

### Download clicks / 按什么按钮

1. 打开 `https://data.imf.org/Datasets/WEO`。
2. 先确认页面标题写 **October 2026 World Economic Outlook**（避免选到 April 2026 / October 2025 旧版）。
3. 页面下方找 **Explore the latest WEO**。
4. 点击 **October 2026 WEO Entire Dataset in Excel**。这是官方提供的全量 Excel，不是只下载一个国家的 dashboard。
5. 下载后文件预计叫 `WEOOct2026all.xlsx`；原样放入 `data/raw/IMF/`。
6. 如果公司网络挡住直接下载，使用上方 Landing Page 链接点击；不要改用另一个 vintage 的数据并与 October 2026 混合。

### Workbook navigation / 进入 Excel 后看哪里

IMF 官方页面明确写明 workbook 包含 **3 个 datasheets**：

| Excel sheet | Contains | 本项目怎样用 |
|---|---|---|
| **`Countries`** | Individual-country annual WEO indicator series | 找 `United States`、`China` 的 2030 nominal GDP；必要时汇总全部国家 |
| **`Country Groups`** | IMF group aggregates such as World / advanced economies / regions | 优先寻找 **World** 的 nominal GDP（如果这一指标有提供） |
| **`Commodity Prices`** | Commodity-price series | **本项目不需要** |

**解释 `Country Groups`：** 它不是让你自己逐国相加的“country list”。它放的是 IMF 已经编好的经济体组别/区域汇总数据（例如 `World`）。但**不是每个 indicator 都一定有 World 汇总**，所以必须先核对实际是否有 `NGDPD`，没有就回 `Countries` 求和。

## A2. Exact GDP indicator / 必须找哪个指标

Use:

- **Indicator / WEO subject code:** `NGDPD`
- **Description:** `Gross domestic product, current prices` expressed in **U.S. dollars**
- **Year:** `2030`
- **Units:** normally **billions of U.S. dollars**; **verify the workbook's Units/Scale column** before calculating

**Do NOT use:**

- `PPPGDP`: PPP-based GDP in international dollars — **不是名义美元 GDP**；
- `NGDP`: GDP in current **local currency**;
- `NGDP_RPCH`: real GDP **growth rate** (percentage), not GDP level;
- `NGDPDPC`: GDP **per capita**, not aggregate GDP.

**Excel 找法（不依赖固定列号）：**

1. 在目标 sheet 打开 `Ctrl + F`，先搜 `NGDPD`。
2. 确认同一行 / record 的 `Subject Descriptor`、`WEO Subject Code`、`Indicator` 或同义字段，明确是 *GDP current prices, U.S. dollars*。
3. 再筛选 `Country` / `Country Group` 或同义字段，并定位 `2030` 这一年。若数据是宽表：2030 通常是列；若是长表：2030 是 `Time Period` / `Year` 的取值。根据实际文件来。
4. 检查 `Unit`, `Scale`, `Units` 字段；**如果是 billions USD，用该值除以 1,000 才得到 USD trillion**。如果已经是 USD trillion，就不要再除。
5. 如果同一个名称出现多个指标，务必优先按 `NGDPD` + `2030` + `Unit` 三个条件识别。

## A3. Pull the three GDP numbers / 具体筛选顺序

### (1) U.S. 2030 GDP

1. 点击 Excel 底部 `Countries`。
2. 筛选 `United States`（国家名称可能是 `United States`，不要选 `United States Virgin Islands`）。
3. 筛选 `NGDPD`，找到 `2030`。
4. 记录原始值、原始单位和转换后 USD trillion。
5. 在自己 Excel 的 `01_IMF_GDP` tab 记录 `US_GDP_2030`。

### (2) China 2030 GDP

1. 仍在 `Countries`。
2. 筛选 `China`（**只选国家 China**，不额外加上 `Hong Kong SAR` / `Macao SAR` / `Taiwan Province of China`）。
3. `NGDPD` + `2030`；记录 `China_GDP_2030`。
4. 若你老板定义的 *Ex-China* 范围不同，必须另写 Ex-China coverage 注释；不要擅自增加扣减的经济体。

### (3) World 2030 GDP

**First choice:**

1. 点击 `Country Groups` sheet。
2. 筛选 `World` / total global aggregate（以实际 group 名称为准）。
3. 看有没有 `NGDPD` 的 `2030`：若有且单位是名义 USD，使用它。
4. 注意：世界 `real GDP growth (%)` **不是**世界 `nominal GDP (USD)`，绝不能替换。

**Fallback if `NGDPD` is missing for World:**

1. 回 `Countries` sheet。
2. 筛选出 **`NGDPD` + `2030`** 的所有 individual-country rows。
3. **仅汇总国家，不加入任何 Country Groups rows**。Excel 可以用 `SUM` 对筛选后的数值列求和，或把相关字段复制到自己 workbook 的 `Raw_IMF_Countries` 表后用 `SUMIFS`。
4. 检查缺失值、重复国家和单位。对 `2030` 缺失的国家要单独列出，不能自动当作零且不披露。
5. 结果命名为 `World_GDP_2030_CountrySum`，并注明 `WEO-covered economies, available data only`。它不一定与官方 World composite 完全一致。
6. 最好再与 IMF WEO DataMapper / 可靠全球统计口径交叉检查。**不要在 World NGDPD 缺失时，把 PPP GDP 强行换成名义美元 GDP。**

### (4) Global Ex-China 2030 GDP

在你自己的模型里计算（IMF 不一定直接给 *World excluding China*）：

```excel
=World_GDP_2030-China_GDP_2030
```

如果所有值存成 USD trillion，那么结果也就是 USD trillion。若存成 USD billion，则先相减、最后整体除 1,000。

### (5) Optional: Global Ex-U.S. and Global Ex-U.S.-Ex-China

```excel
=World_GDP_2030-US_GDP_2030
=World_GDP_2030-US_GDP_2030-China_GDP_2030
```

提醒：`Global Ex-China` **仍包含美国**；不要把美国 $4.8T 再与 Ex-China 全球工资池相加。

## A4. How to copy to your model / 建议做一个最简单的审计表

建立 `01_IMF_GDP` sheet：

| Input name | Original source sheet | Geography | Indicator | Year | Raw value | Unit / scale | USD trillion | Vintage |
|---|---|---|---|---:|---:|---|---:|---|
| US GDP | Countries | United States | NGDPD | 2030 | [paste] | [check] | [formula] | Oct-2026 |
| China GDP | Countries | China | NGDPD | 2030 | [paste] | [check] | [formula] | Oct-2026 |
| World GDP | Country Groups **or documented country sum** | World | NGDPD | 2030 | [paste] | [check] | [formula] | Oct-2026 |
| World Ex-China | Calculated | World ex China | Derived | 2030 | n/a | USD trillion | =World-China | Oct-2026 |

**QA checks**

- World GDP > U.S. GDP and > China GDP.
- World Ex-China = World − China, using the **same IMF vintage, year and units**.
- All inputs are nominal USD GDP at market exchange rates (not PPP and not real GDP).
- The IMF workbook is a projection vintage, **not an observed 2030 outcome**.

---

# B. ILOSTAT — Labor Income Share of GDP (PRIORITY 2)

## B1. Official page and which dataset to download

**Official ILOSTAT Data catalogue:**  
https://ilostat.ilo.org/data/

Go to **Indicator catalogue** (scroll down; do **not** use the “Quick downloads” labour force participation data).

Find one of these relevant series:

1. **`SDG indicator 10.4.1: Labour income share as a percent of GDP (%)`**, annual; OR
2. **`Labour income share as a percent of GDP — ILO modelled estimates`**, annual, if a consistent modeled worldwide series is needed.

Click the indicator row and use **Downloads** or open the **Data Explorer**, filtering the relevant countries and years. Save the spreadsheet/CSV in `data/raw/ILOSTAT/`.

**Important:** These are published estimates for historical/current years, not an official 2030 Labor Share forecast. Applying a recent labor share to 2030 GDP is **your model assumption**, not an ILO 2030 projection.

## B2. What to collect

| Item | Geography | Data field | Year | Use |
|---|---|---|---|---|
| `US_Labor_Share` | U.S. | Labor income as % of GDP | Latest comparable published year | U.S. denominator |
| `World_Labor_Share` | World (if available) | Same definition | Same reference year | Global adjustment |
| `China_Labor_Share` | China | Same definition | Same reference year | Calculate Ex-China adjustment |
| `Country_Labor_Share` | Optional country list | Same definition | Latest comparable year | Country-level validation |

**Formatting:** if downloaded value is `52.6`, use Excel `52.6%` or `=52.6/100`, **not** the number 52.6 in multiplicative formulas.

**Definition check:** ILO labor income share includes employee compensation **plus imputed labor income of the self-employed**; employee compensation includes employer social contributions. This may not match a U.S. OEWS wage-only dataset exactly. Document any adjustment before comparing pooled dollar values.

## B3. Ex-China labor share — not identical to World labor share

If you have `World_Labor_Share` and `China_Labor_Share` on consistent definitions, compute an **illustrative 2030 ex-China share** by GDP weighting:

```excel
=(World_GDP_2030*World_Labor_Share-China_GDP_2030*China_Labor_Share)/(World_GDP_2030-China_GDP_2030)
```

**Caveat:** this is a **scenario that applies historical labor shares to the 2030 GDP structure**; it is not a directly observed or ILO-forecast 2030 share. Prefer labor-share observations from the **same historical year**; record which year and whether country definitions match. If data are too patchy, leave this as a sensitivity assumption instead of filling in an invented value.

---

# C. ILO / GitHub — Occupation-level GenAI Exposure (PRIORITY 3)

## C1. GitHub entry point and exact filenames

**Repository:**  
https://github.com/pgmyrek/2025_GenAI_scores_ISCO08

**Data dictionary (please read first):**  
https://github.com/pgmyrek/2025_GenAI_scores_ISCO08/blob/main/DATA_DICTIONARY.md

**Main score workbook — download this first:**  
https://github.com/pgmyrek/2025_GenAI_scores_ISCO08/blob/main/Final_Scores_ISCO08_Gmyrek_et_al_2025.xlsx

**Simplified occupation/task workbook — optional:**  
https://github.com/pgmyrek/2025_GenAI_scores_ISCO08/blob/main/4digits_with_tasks.xlsx

**Original ILO research paper:**  
https://www.ilo.org/publications/generative-ai-and-jobs-refined-global-index-occupational-exposure

### GitHub clicks / 如何下载

**Option A (recommended):**

1. 打开 Repository 链接。
2. 在文件列表里点击 `Final_Scores_ISCO08_Gmyrek_et_al_2025.xlsx`。
3. 在文件页面按 **Download raw file**（或下载图标）。有时 GitHub 直接预览 Excel 不方便，下载后本地打开更好。
4. 同样下载 `DATA_DICTIONARY.md`（变量解释）和可选的 `4digits_with_tasks.xlsx`。
5. 另一种方式：repo 页面绿色 **Code** 按钮 → **Download ZIP** → 解压，再拿上述文件。

保存原文件到 `data/raw/ILO_AI_Exposure/`；不要覆盖原表或修改官方列名。

## C2. What does the Excel contain?

按照该 GitHub 项目的 `DATA_DICTIONARY.md`，主文件有 **3,265 task-level records**, covering **427 ISCO-08 four-digit occupations**。每行是一个任务，**不是一个员工，不是一个行业**。

Useful columns (exact spelling from published dictionary):

| Column | Interpretation | How to use |
|---|---|---|
| `ISCO_08` | Four-digit occupation code | Join key / link to ILOSTAT occupational employment |
| `Title` | Occupation title | Readable label |
| `Task_ISCO` | Task description | Detail; optional |
| `score_2025` | Final 2025 task-level exposure score (0–1) | Task analysis |
| `mean_score_2025` | Mean task exposure score for that occupation | Better candidate for occupation-level mapping |
| `potential25` | Occupational exposure category / gradient | Categorical cross-check |
| `SD_2025` | Standard deviation of task-level scores within occupation | Sensitivity / heterogeneity |

**Critical**: `mean_score_2025` 和 `potential25` 可能在同一职业的每一个 task row 上重复。绝对不能把 3,265 行当成 3,265 个 occupation 后直接相加！

### Make an occupation-level lookup table

1. 在自己的 workbook 新建 `03_ILO_Exposure`。
2. 从原表复制 `ISCO_08`, `Title`, `mean_score_2025`, `potential25`。
3. **Data → Remove Duplicates → 仅按 `ISCO_08` 去重**，保留每种职业 1 行。
4. 核对职业代码 unique 数量，应与源资料所列约 427 个职业一致。
5. 用 `ISCO_08` 与国家/行业职业人数匹配；不同分类版本（ISCO-08 vs 其他）必须有正式 crosswalk。
6. `mean_score_2025` 是**任务分数的平均值**，不是精确的“可节省总工时比例”。若你的美国模型使用的是 `exposure-weighted time`，需要注意**概念不完全相同**，不能默默直接替换。

---

# D. ILOSTAT — Employment by Industry × Occupation (OPTIONAL: BOTTOM-UP)

Official entry page: https://ilostat.ilo.org/data/

Indicator catalogue → search:

**`Employment by economic activity and occupation (thousands)`** (Annual)

This is the bridge you need when you want to go beyond a simple GDP ratio and **really estimate country × industry × knowledge worker × AI exposure**.

### Click / export procedure

1. 在 ILOSTAT 的 **Indicator catalogue** 搜 `Employment by economic activity and occupation`。
2. 选 **Annual**（避免月度/季度数据与你的工资年度数据错位）。
3. 点击 **Downloads**，或进入 Data Explorer，筛选国家、年份、industry、occupation，并导出 CSV/Excel。
4. 保留字段：`country`, `year`, `industry code`, `occupation code`, `employment`, `unit`, `sex`（如果有）。
5. 如果出现 `Total` 与细分行，**只选择一个一致的分组级别**；不要把 total + subgroups 重复加总。
6. 检查单位：`thousands` = 千人。例如 Excel 显示 `125.0` 意味着 125,000 workers。要变成人数乘以 1,000。
7. 如果某国只有 ISCO 1-digit / 2-digit 的职业数据，就不能直接与 ILO `ISCO_08` 4-digit scores 一一匹配；必须先聚合/映射并披露信息损失。

### Economic logic / 你真正想算的 bottom-up 公式

```text
Country × Industry × Occupation employment
× Annual labor compensation per employee
× Comparable occupation AI-exposure weight
= Country × Industry × Occupation AI-exposed compensation
```

Then sum all combinations, excluding China where needed.

**Data caveat:** ILOSTAT employment is available for many countries, but not every country has a complete fine-grained industry × occupation matrix. Wage-by-occupation-and-industry coverage is often weaker still. That is exactly why the GDP-based top-down model is the first pass.

---

# E. World Bank — Industry Value Added Weights (OPTIONAL: FASTER INDUSTRY VALIDATION)

If you want **industry economic weights** but do not want to rebuild all global wage data, use World Bank World Development Indicators (WDI), with the caveat that the following is only a broad **3-sector** split.

| Industry | WDI indicator code | Direct source page |
|---|---|---|
| Agriculture, forestry, fishing | `NV.AGR.TOTL.ZS` | https://data.worldbank.org/indicator/NV.AGR.TOTL.ZS |
| Industry, including construction | `NV.IND.TOTL.ZS` | https://data.worldbank.org/indicator/NV.IND.TOTL.ZS |
| Services | `NV.SRV.TOTL.ZS` | https://data.worldbank.org/indicator/NV.SRV.TOTL.ZS |

### How to download and use

1. 点上方一个指标网页，例如 Services。
2. 选择 **Download → Excel** 或 **CSV**。
3. 取与你 ILO exposure / employment 数据最接近的历史年份；不要假装这三个指标本身给出 2030 行业预测。
4. 把该百分比转换为小数：`64.0% → 0.64`。
5. 注意 value added sector weights **不一定精确加总为 100% of GDP**（可能存在 product taxes less subsidies 等 national accounts 项目），不能强制补到 100%。
6. 如果要精细到 IT、Financials、Health Care 等行业，再转向 OECD STAN / TiM、UN National Accounts、ILOSTAT 等更细 ISIC 数据。

**OECD TiM 2025 edition (optional detailed industry employment/compensation):**  
https://data-explorer.oecd.org/vis?df%5Bag%5D=OECD.STI.PIE&df%5Bds%5D=dsDisseminateFinalDMZ&df%5Bid%5D=DSD_TIM_2025%40DF_TIM_2025&df%5Bvs%5D=1.0

OECD TiM 2025 edition covers 50 industries and historical series through **2022** (not a 2030 forecast). Use it as a *structure and calibration reference*, not as ready-made 2030 projections.

---

# F. Your first-pass model — simple, auditable Excel formulas

**Recommended workbook tabs**

| Sheet | Purpose |
|---|---|
| `00_Readme` | Model definitions, source links, vintage, assumptions |
| `01_IMF_GDP` | U.S. / China / World nominal GDP 2030 |
| `02_ILO_Labor_Share` | U.S. / China / World labor income share, with actual observation years |
| `03_ILO_Exposure` | Unique ISCO-08 occupational exposure lookup |
| `04_TopDown_Calc` | U.S. wage pool → Global and Global Ex-China scenarios |
| `05_Sensitivity` | Relative Exposure Factor × wage growth × labor share |
| `06_Country_Industry` | Later bottom-up validation; optional |

**Simple `04_TopDown_Calc` input layout** (enter true Excel percentage values, e.g. 3.5%, not 3.5):

| Cell | Label | Value / source |
|---|---|---|
| `B2` | U.S. AI-Exposed Wage Pool — base year (USD tn) | Existing U.S. model |
| `B3` | Nominal Wage Growth | Scenario, e.g. 3.5% |
| `B4` | U.S. Wage Pool Base Year | Your current model vintage, e.g. 2025 |
| `B5` | Target Year | 2030 |
| `B6` | U.S. GDP (2030, USD tn) | IMF |
| `B7` | World GDP (2030, USD tn) | IMF |
| `B8` | China GDP (2030, USD tn) | IMF |
| `B9` | U.S. Labor Share | ILOSTAT / reconciled comparable definition |
| `B10` | World Labor Share assumption for 2030 | Historical ILOSTAT benchmark, held constant or forecast |
| `B11` | China Labor Share assumption for 2030 | Historical ILOSTAT benchmark, held constant or forecast |
| `B12` | Relative Wage-Weighted Exposure Factor — Global Ex-China vs U.S. | **Scenario, not an official published rate** |
| `B13` | Relative Wage-Weighted Exposure Factor — World vs U.S. | **Scenario, not an official published rate** |

Put these outputs below (example locations; formulas can be copied directly after setting inputs):

| Cell | Label | Formula | Meaning |
|---|---|---|---|
| `B16` | U.S. Exposed Wages 2030 | `=B2*(1+B3)^(B5-B4)` | If using wage growth only; keep employment and exposure assumptions separate |
| `B17` | World GDP Ex-China 2030 | `=B7-B8` | Still includes United States |
| `B18` | Ex-China Labor Share scenario | `=(B7*B10-B8*B11)/B17` | 2030 GDP-weighted share under chosen labor-share assumptions |
| `B19` | Ex-China GDP / U.S. GDP | `=B17/B6` | Geographic scale factor |
| `B20` | Ex-China AI-Exposed Wage Pool 2030 | `=B16*B19*(B18/B9)*B12` | **Scenario** |
| `B21` | World AI-Exposed Wage Pool 2030 | `=B16*(B7/B6)*(B10/B9)*B13` | **Scenario** |

### Very important interpretation of the exposure factor

`Relative Wage-Weighted Exposure Factor` means **relative exposure-weighted labor income per dollar of labor income** in the target geography versus the U.S., using **comparable definitions**.

It is **NOT** equivalent to any one of the following:

- share of workers who have tried ChatGPT;
- share of firms reporting AI adoption;
- share of occupations with *some* GenAI exposure;
- measured hours actually automated;
- captured vendor revenue; or
- actual annual labor savings.

For a first-pass sensitivity, you may enter a range as **explicit assumptions** (e.g., 60%, 80%, 100%), but do not label them as ILO estimates. Use ILO occupational / country composition research to inform the factor later.

**Special note on wage growth:** If your original U.S. $4.8T model already represents **2030** wages, do **not** grow it again from 2025 to 2030. If it is a 2025 wage pool, explicitly document whether you assume 2030 employment and exposed-time shares remain unchanged. The wage CAGR alone is not the growth rate of the total national wage bill unless employment/hours are also held fixed.

---

# G. QA / sanity checks before sending to Keith

- [ ] IMF vintage is **October 2026**, target year **2030**, and GDP is `NGDPD` nominal USD (not PPP).
- [ ] World, China and U.S. GDP use the **same vintage and unit**.
- [ ] World `NGDPD` is truly reported; if not, `World_CountrySum` is clearly labeled as an approximate covered-country sum.
- [ ] **Ex-China definition** is explicit; the U.S. is still included in Ex-China.
- [ ] The source labor-share observation year is recorded for every country.
- [ ] Percent inputs are decimals in Excel (`52.6%`, not `52.6`).
- [ ] No conversion between local-currency wages and USD without a documented FX rate / basis.
- [ ] No double counting `World_ExChina` plus `U.S.` or unique ISCO-08 occupations vs repeated task rows.
- [ ] AI-exposed wages are clearly labeled **potential wage exposure**, not actual AI revenue/savings.
- [ ] Global and Ex-China outputs are **model scenarios**, not official IMF/ILO forecasts.
- [ ] Every model input has a source URL, retrieval date, unit, vintage, and any mapping/assumption described.

## Suggested boss-facing methodology descriptions

**Global nominal GDP**  
“2030 nominal GDP in U.S. dollars from the IMF October 2026 WEO; Global Ex-China is calculated as World GDP less China GDP.”

**Global AI-exposed wages — top-down**  
“U.S. AI-exposed compensation scaled by the 2030 GDP ratio, adjusted for labor income share and relative wage-weighted AI exposure.”

**Global AI-exposed wages — bottom-up validation**  
“Country and industry labor compensation weighted by occupation-level GenAI exposure, aggregated across countries; excludes China in the ex-China scenario.”

---

# H. Official links / source registry

| ID | Publisher / dataset | Link | Download/use |
|---|---|---|---|
| `IMF_WEO` | IMF World Economic Outlook | https://data.imf.org/Datasets/WEO | **Required:** Oct 2026 entire workbook |
| `IMF_WEO_XLSX` | IMF WEO Oct 2026 Entire Dataset | https://data.imf.org/-/media/iData/External-Storage/Documents/0FA3DC49A04D446E9D69DBEB29979E1C/en/WEOOct2026all.xlsx | Direct xlsx link |
| `IMF_FAQ` | WEO Database FAQ | https://data.imf.org/Datasets/WEO/Frequently-Asked-Questions | Indicator / forecast / aggregate definitions |
| `ILOSTAT` | ILOSTAT Indicators | https://ilostat.ilo.org/data/ | Labor share; optional industry×occupation employment |
| `ILO_SDG_DEF` | UN/ILO Labor Income Share Definition | https://unstats.un.org/sdgs/metadata/files/Metadata-10-04-01.pdf | Check compensation scope |
| `ILO_AI_REPO` | ILO-NASK GitHub 2025 scores | https://github.com/pgmyrek/2025_GenAI_scores_ISCO08 | Download AI exposure scores |
| `ILO_AI_MAIN` | ILO AI Task Scores xlsx | https://github.com/pgmyrek/2025_GenAI_scores_ISCO08/blob/main/Final_Scores_ISCO08_Gmyrek_et_al_2025.xlsx | Main exposure data |
| `ILO_AI_DICTIONARY` | ILO AI Data Dictionary | https://github.com/pgmyrek/2025_GenAI_scores_ISCO08/blob/main/DATA_DICTIONARY.md | Exact variable definitions |
| `WDI_AGRI` | World Bank Agriculture Value Added % GDP | https://data.worldbank.org/indicator/NV.AGR.TOTL.ZS | Optional industry breakdown |
| `WDI_INDUSTRY` | World Bank Industry Value Added % GDP | https://data.worldbank.org/indicator/NV.IND.TOTL.ZS | Optional industry breakdown |
| `WDI_SERVICES` | World Bank Services Value Added % GDP | https://data.worldbank.org/indicator/NV.SRV.TOTL.ZS | Optional industry breakdown |
| `OECD_TIM` | OECD TiM 2025 edition | https://data-explorer.oecd.org/vis?df%5Bag%5D=OECD.STI.PIE&df%5Bds%5D=dsDisseminateFinalDMZ&df%5Bid%5D=DSD_TIM_2025%40DF_TIM_2025&df%5Bvs%5D=1.0 | Optional detailed industry validation |

**Data provenance:** Source locations and public file/indicator names checked against official IMF, ILO/ILOSTAT, World Bank, OECD and the linked GitHub repository as of 2026-10-22. Workbook-specific column coordinates are intentionally not asserted without direct inspection of downloaded bytes. Preserve the source vintage when reproducing the analysis.
