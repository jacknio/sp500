# Step by Step — Oil Shock, Inflation & Demand Replication

**Research snapshot:** October 18, 2026  
**Workbook:** `Oil_Inflation_Replication_v5.xlsx`  
**Goal:** Rebuild a short, source-backed validation of the oil-price / inflation-transmission thesis using EIA, BEA, BLS and Census data.

> **Workbook safety rule:** ALL URLs and downloading instructions are in this README, not the `.xlsx`. The workbook has no hyperlinks, URL text, `WEBSERVICE`, Power Query connections, or external-workbook references. Only paste *values* downloaded from official sources into the blue input cells. Do not embed the source CSVs or source URLs inside the workbook.

## 1. Folder setup and download procedure

Recommended local folder on your company computer:

```text
Oil_Inflation_Research/
├── README.md
├── Oil_Inflation_Replication_v5.xlsx
└── raw_data/
    ├── eia_table_01_07.xlsx
    ├── eia_table_03_01.xlsx
    ├── GDPCA.csv
    ├── DGOERC1A027NBEA.csv
    ├── PCECA.csv
    ├── CPIAUCSL.csv
    ├── CPILFESL.csv
    ├── CPIENGSL.csv
    ├── CUSR0000SAS4.csv
    ├── RRSFS.csv
    ├── PCESC96.csv
    ├── RSAFS.csv
    ├── RSGASS.csv
    ├── RSMVPD.csv
    ├── RSFSDP.csv
    └── RSBMGESD.csv
```

**FRED / BEA / BLS / Census series:** Open a link from the dataset list below and download its CSV. The direct `fredgraph.csv` links below work without an API key. In the browser, you can alternatively open the `series/` page and select **Download → CSV**. If Excel automatically misinterprets the date or numeric columns, use **Data → From Text/CSV**, choose **Comma** delimiter, and import as a new scratch sheet/file; then copy only values into the provided research workbook.

**EIA tables:** Open the EIA browser page, choose **Annual** (not monthly), click **Download → Excel** or **CSV**, and save the downloaded table separately in `raw_data/`. EIA downloads can include series metadata and different columns; use the row labels and units specified below rather than relying on fixed column positions.

**Note on revisions:** Sources update retrospectively. This workbook is a snapshot as of October 18, 2026; downloaded values on a later date may differ slightly. Record your pull date and do not silently mix vintages.

## 2. Source-to-cell map (every input series; 16 entries)

| Workbook | Institution | Original dataset / series code | Original page | Direct download / action | Needed observations |
|:---|:---|:---|:---|:---|:---|
| 01 | **EIA** | Monthly Energy Review, **Table 1.7** (`T01.07`) | [EIA Table 1.7](https://www.eia.gov/totalenergy/data/browser/index.php?tbl=T01.07) | Select **Annual → Download → Excel** (or [EIA CSV](https://www.eia.gov/totalenergy/data/browser/csv.php?tbl=T01.07)) | 1970, 1975, 1980–2025; primary energy and published energy intensity |
| 01, 02 | **BEA via FRED** | **GDPCA** — Real Gross Domestic Product | [GDPCA](https://fred.stlouisfed.org/series/GDPCA) | [GDPCA.csv](https://fred.stlouisfed.org/graph/fredgraph.csv?id=GDPCA) | Selected years 1970–2025 |
| 02 | **EIA** | Monthly Energy Review, **Table 3.1** (`T03.01`) | [EIA Table 3.1](https://www.eia.gov/totalenergy/data/browser/index.php?tbl=T03.01) | Select **Annual → Download → Excel** (or [EIA CSV](https://www.eia.gov/totalenergy/data/browser/csv.php?tbl=T03.01)) | Selected years 1970–2025; total **petroleum products supplied**, not crude production |
| 03 | **BEA via FRED** | **DGOERC1A027NBEA** — PCE: Gasoline and Other Energy Goods | [Series page](https://fred.stlouisfed.org/series/DGOERC1A027NBEA) | [DGOERC1A027NBEA.csv](https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGOERC1A027NBEA) | Annual 1970–2025, current-dollar $bn |
| 03 | **BEA via FRED** | **PCECA** — Total Personal Consumption Expenditures | [Series page](https://fred.stlouisfed.org/series/PCECA) | [PCECA.csv](https://fred.stlouisfed.org/graph/fredgraph.csv?id=PCECA) | Annual 1970–2025, current-dollar $bn |
| 04 | **BLS via FRED** | **CPIAUCSL** — All-Items CPI | [Series page](https://fred.stlouisfed.org/series/CPIAUCSL) | [CPIAUCSL.csv](https://fred.stlouisfed.org/graph/fredgraph.csv?id=CPIAUCSL) | Monthly 2018-01 to 2026-09; SA index |
| 04 | **BLS via FRED** | **CPILFESL** — Core CPI Excluding Food and Energy | [Series page](https://fred.stlouisfed.org/series/CPILFESL) | [CPILFESL.csv](https://fred.stlouisfed.org/graph/fredgraph.csv?id=CPILFESL) | Monthly 2018-01 to 2026-09; SA index |
| 04 | **BLS via FRED** | **CPIENGSL** — Energy CPI | [Series page](https://fred.stlouisfed.org/series/CPIENGSL) | [CPIENGSL.csv](https://fred.stlouisfed.org/graph/fredgraph.csv?id=CPIENGSL) | Monthly 2018-01 to 2026-09; SA index |
| 04 | **BLS via FRED** | **CUSR0000SAS4** — Transportation Services CPI | [Series page](https://fred.stlouisfed.org/series/CUSR0000SAS4) | [CUSR0000SAS4.csv](https://fred.stlouisfed.org/graph/fredgraph.csv?id=CUSR0000SAS4) | Monthly 2018-01 to 2026-09; SA index |
| 05 | **BEA via FRED** | **PCEC96** — Real Personal Consumption Expenditures | [Series page](https://fred.stlouisfed.org/series/PCEC96) | [PCEC96.csv](https://fred.stlouisfed.org/graph/fredgraph.csv?id=PCEC96) | Monthly 2020-01 to 2026-08; chained 2017 $bn, SAAR |
| 05 | **St. Louis Fed via FRED** | **RRSFS** — Advance Real Retail and Food Services Sales | [Series page](https://fred.stlouisfed.org/series/RRSFS) | [RRSFS.csv](https://fred.stlouisfed.org/graph/fredgraph.csv?id=RRSFS) | Monthly 2020-01 to 2026-09; millions of 1982–84 CPI-adjusted dollars, SA |
| 05 | **BEA via FRED** | **PCESC96** — Real Personal Consumption Expenditures: Services | [Series page](https://fred.stlouisfed.org/series/PCESC96) | [PCESC96.csv](https://fred.stlouisfed.org/graph/fredgraph.csv?id=PCESC96) | Monthly 2020-01 to 2026-08; chained 2017 $bn, SAAR |

**Important:** the new 05 sheet uses three **real** demand series only. It no longer relies on a nominal retail control-group construction, because nominal retail dollars can confuse price and quantity. The purpose is to judge whether the *level and path of actual consumer demand* looks hot enough to amplify oil pass-through.

## 3. Tab-by-tab reproduction in Excel

### `01_Energy_GDP` — total energy intensity

1. From EIA **Table 1.7**, select **Annual** and export. Locate **Total Primary Energy Consumption (quadrillion Btu)** and the EIA-published **Consumption per Dollar of Real GDP (thousand Btu per dollar)**. Copy the **Year** into `A8:A55`, primary energy into `B8:B55`, published intensity into `D8:D55`. The workbook intentionally uses a selected historical sample, not every year from 1970 to 2025.
2. From **GDPCA.csv**, match each year and paste BEA real GDP ($bn, 2017 chain dollars) into `C8:C55`.
3. `E8 = B8*1000/C8` (quad Btu / billion dollars, converted to thousand Btu per dollar).  
   `F8 = D8/$D$8*100`; `G8 = 1-D8/$D$8`. Drag down.
4. Check **1970 EIA published 12.42**, **2025 4.04** (000 Btu/$), approximately a **67.5% decline**. The EIA-published D column is the benchmark; the recomputation in E can differ slightly due to GDP-vintage differences.
5. This demonstrates less **total energy per real GDP**, not less crude oil specifically.

### `02_Petroleum_GDP` — petroleum-specific volume intensity

1. From EIA **Table 3.1**, switch to **Annual** and locate **Petroleum Products Supplied** (**thousands of barrels per day**), using the **total petroleum** row, not individual fuels.
2. Match the 24 selected years listed in `A8:A31`; enter petroleum supplied into `B8:B31`.
3. For those same years copy **GDPCA** into `C8:C31`.
4. `D8 = B8/C8`; `E8 = D8/$D$8*100`; `F8 = 1-D8/$D$8`; `G8 = $D$8/D8`. Drag down.
5. This is **petroleum product demand proxied by products supplied divided by real GDP**, not crude oil input coefficients or a formal oil-import dependence metric.

### `03_Gas_PCE` — consumer gasoline / energy-goods budget share

1. Download **DGOERC1A027NBEA.csv**. Paste annual 1970–2025 gasoline **and other energy goods** spending ($bn current dollars) into `B8:B63`.
2. Download **PCECA.csv**. Paste annual 1970–2025 total **nominal** PCE ($bn) into `C8:C63`.
3. `D8 = B8/C8`; `E8 = (D8-$D$18)*100` (percentage point difference vs **1980**); `F8 = D8/$D$18*100`. Drag down. **1980 = row 18**.
4. Anchors: **1980 = 101.898 / 1750.667 ≈ 5.82%**; **2025 = 418.814 / 20870.150 ≈ 2.01%**. Both numerator and denominator are nominal, so the ratio is internally consistent.
5. This is a **spending share**, affected by quantities **and** prices, not gallons consumed or total household indirect energy exposure.

### `04_Inflation` — 2018–2026 longer CPI history

1. Download the **four BLS/FRED CPI series** from the table above. Use **original index levels**, not FRED's transformed percent-change view; all series are **seasonally adjusted**.
2. Match by `YYYY-MM` and paste **headline / core / energy / transportation services** index levels into `B8:E112`, **2018-01 through 2026-09**.
3. Beginning in row 20 (**2019-01**), calculate **YoY**: `F20 = IF(OR(B20="",B8=""),"",B20/B8-1)`; similarly use `C`→`G`, `D`→`H`, `E`→`I`. Copy down. The charts compare headline with core, and energy with transport separately to avoid scale distortion.
4. **Keep 2025-10 missing CPI observations blank** if still missing from the official series. Do not interpolate or use zero to fill missing historical CPI readings.
5. Reading the chart: price changes in energy categories can be much larger than in core CPI. That co-movement **does not by itself quantify causal pass-through**.

### `05_Demand_Retail` — consumer demand backdrop, 2020–2026 real time series (v6)

**Why change?** The earlier 05 sheet compared 2022 with 2026 directly, but that makes it harder to tell the broader demand story. The new version uses a clean **2020–2026 monthly time series** of three real demand indicators: total real PCE, real retail sales, and real services PCE. This is better aligned with the thesis that *consumer demand is not as overheated as in the 2021–2022 phase*, which may help limit pass-through from oil into broader inflation. **This remains descriptive evidence, not a causal estimate.**

#### How to rebuild the new `05_Demand_Retail`

1. Download three CSV files: **PCEC96**, **RRSFS**, and **PCESC96** from FRED.
2. Keep observations from **2020-01** through the latest available month. For the current workbook snapshot, `PCEC96` and `PCESC96` run through **2026-08**, while `RRSFS` runs through **2026-09**.
3. In Excel, place the monthly dates in column **A** and the three raw series in columns **B**, **D**, and **F**.
4. Create simple index columns so every series starts at **100 in 2020-01**:
   - `C13 = IF(B13="","",B13/$B$13*100)`
   - `E13 = IF(D13="","",D13/$D$13*100)`
   - `G13 = IF(F13="","",F13/$F$13*100)`
   Drag each formula down.
5. Make one **line chart** using the three index columns (**C, E, G**) against the date column (**A**).
6. Optional quick read boxes used in the workbook:
   - retail peak = `MAX(D13:D93)`
   - Sep 2026 vs retail peak = `D93/J8-1`
   - Aug 2026 real PCE vs Jan 2020 = `B92/B13-1`
   - Aug 2026 services vs Jan 2020 = `F92/F13-1`

#### Interpretation

- **Real retail sales** are the cleanest proxy here for goods demand. They surged in 2021, peaked in 2022, and then moved sideways; by **Sep 2026** they were still below the **Apr 2022** peak.
- **Real PCE** and **real services PCE** continued to rise, but on a steadier slope than the post-reopening burst.
- That combination is more consistent with **firm but not overheating consumer demand**, which supports the argument that oil-price shocks may face a more limited pass-through into broader inflation than they did in 2021–2022.


