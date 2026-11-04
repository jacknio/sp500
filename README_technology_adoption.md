# Technology adoption curves: Generative AI vs. earlier digital technologies

**Data snapshot:** November 4, 2026  
**Geography:** United States  
**Coverage:** Personal computers, Internet, home broadband, social networking, smartphones, and generative AI.

This project compares the diffusion of major digital technologies using **years since a chosen mass-market introduction date** rather than calendar years. It uses historical observations from the U.S. Census Bureau, World Bank / ITU, the Real-Time Population Survey (hosted on FRED), and Pew Research Center. **No future AI adoption points are forecast.**

> **Research takeaway:** Generative AI has a short *observed* adoption history relative to earlier digital technologies, even though initial diffusion has been very fast. This is evidence that its long-run economic adoption path is not yet well established; **it is not, by itself, evidence against an AI valuation bubble or infrastructure oversupply**.

## Charts

### 1. Full adoption history (up to 25 years after the selected launch year)

![Technology adoption curves](charts/technology_adoption_full.png)

### 2. Early adoption (first 10 years)

![Technology adoption, first 10 years](charts/technology_adoption_first_ten_years.png)

### 3. Years to first *observed* 50% adoption

![Years to first observed 50%](charts/technology_years_to_50pct.png)

> **Important:** This last figure uses the **first available measurement at or above 50%**, not a model-estimated crossing date. The starting year and denominator vary by technology, so the ranking is illustrative, not a controlled comparison of diffusion speed.

## Download the original data

| Series | Data publisher / type | Population and measurement | Original webpage | How to download |
|---|---|---|---|---|
| **Generative AI** | Bick, Blandin & Deming's Real-Time Population Survey, distributed via Federal Reserve Bank of St. Louis FRED (academic survey, **not an official federal government survey**) | U.S. adults ages **18–64** reporting any GenAI use, at work or outside work; quarterly | [FRED series RPSGENAIUSAGESHAREALL](https://fred.stlouisfed.org/series/RPSGENAIUSAGESHAREALL) · [Research project](https://www.genaiadoptiontracker.com/) | [Direct CSV](https://fred.stlouisfed.org/graph/fredgraph.csv?id=RPSGENAIUSAGESHAREALL) or click **Download** on FRED |
| **Internet** | International Telecommunication Union (ITU) / World Bank, distributed by FRED | Individuals using the Internet per 100 people in the U.S.; annual, **all ages** | [FRED series ITNETUSERP2USA](https://fred.stlouisfed.org/series/ITNETUSERP2USA) · [World Bank indicator](https://data.worldbank.org/indicator/IT.NET.USER.ZS?locations=US) | [Direct CSV](https://fred.stlouisfed.org/graph/fredgraph.csv?id=ITNETUSERP2USA) |
| **Personal computers** | **U.S. Census Bureau**, Current Population Survey | Percentage of U.S. **households** with a computer | [Census historical tables, 1984–2009](https://www.census.gov/data/tables/time-series/demo/computer-internet/computer-use-1984-2009.html) · [Census report (PDF)](https://www.census.gov/content/dam/Census/library/publications/2013/demo/p20-569.pdf) | Open the historical tables page and download the relevant table / report. The values transcribed into this project are in the Excel workbook. |
| **Home broadband** | **Pew Research Center** (independent nonprofit; not a government agency) | U.S. adults with home broadband | [Internet / broadband fact sheet](https://www.pewresearch.org/internet/fact-sheet/internet-broadband/) | Use the fact sheet's chart and any available **Download data** control; otherwise use the published dated observations / workbook. |
| **Social networking** | **Pew Research Center** | U.S. adults using a social networking site | [Social Media Usage: 2005–2015](https://www.pewresearch.org/internet/2015/10/08/social-networking-usage-2005-2015/) | Use report data tables or downloadable chart data where available; otherwise use the transcribed observations in the workbook. |
| **Smartphones** | **Pew Research Center** | U.S. adults reporting smartphone ownership | [Mobile phone ownership chart](https://www.pewresearch.org/chart/mobile-phone-ownership-2/) | Use chart's data/download option if available; otherwise use the transcribed observations in the workbook. |

The **FRED link is an official Federal Reserve data distribution site**, but the **GenAI survey itself is conducted by academic researchers**. Distinguish the data host from the original survey producer in reports.

## Files in this repository

```text
.
├── README.md
├── charts/
│   ├── technology_adoption_full.png
│   ├── technology_adoption_first_ten_years.png
│   └── technology_years_to_50pct.png
└── data/
    └── technology_adoption_official_sources_2026.xlsx
```

The workbook contains:

| Worksheet | What it includes |
|---|---|
| `Readme` | Scope, key definitions, and limitations |
| `Raw observations` | Every plotted data point, publication, denominator, metric, and original source URL |
| `Excel chart matrix` | Chart-ready table with `Years since launch` on the X-axis and one column per technology |
| `First above 50pct` | Launch anchors, first observed year at/above 50%, elapsed years |
| `FRED GenAI quarterly` | Quarterly observations from July 2024 through July 2026 |

## How the charts were built

1. **Select a mass-market start year** for each technology; these anchors are analytical conventions, not the technology's date of invention:

   | Technology | Chosen year zero | Rationale / limitation |
   |---|---:|---|
   | Personal computer | **1981** | IBM PC launch; Census household observations begin **1984**, not year zero |
   | Internet | **1995** | Approximate U.S. mass-commercialization milestone; Internet existed earlier |
   | Home broadband | **2000** | Approximate beginning of broad consumer rollout |
   | Social networking | **2004** | Common early social-networking era anchor; sites existed earlier and first selected observation is **2005** |
   | Smartphone | **2007** | iPhone launch; selected Pew ownership series begins **2011** |
   | Generative AI | **2022** | ChatGPT public release on **Nov. 30, 2022**; first selected RPS observation is **July 2024** |

2. Compute `Years since launch = Observation calendar year − Chosen year zero`. Thus **2026 corresponds to year 4 for GenAI** on this calendar-year axis, even though as of November 4, 2026 it has been slightly less than four full years since ChatGPT's release.
3. Plot each published adoption or ownership observation against the shifted axis, leaving years **without observations blank**. Lines connecting observations are visual guides; they are not interpolated annual estimates.
4. For the overview, GenAI uses the **July (Q3)** reading for 2024, 2025, and 2026: **44.6%, 54.5%, 60.8%**. Full quarterly history is in the workbook. The survey began tracking in **July 2024**, so only about two years of regularly published observations are available.
5. Do **not** extend the GenAI curve beyond 2026, fit a predictive S-curve without disclosing assumptions, or treat the present value as an eventual saturation estimate.
6. In the `First above 50pct` worksheet, use the first available observed percentage **≥50%**. Do not call it the exact date of reaching 50%.

### Quick Excel reproduction

1. Open [`data/technology_adoption_official_sources_2026.xlsx`](data/technology_adoption_official_sources_2026.xlsx).
2. Select the `Excel chart matrix` worksheet.
3. Insert a **Line with markers** chart. Use **Years since launch** for X-axis categories and the six technology columns as separate series. Configure empty cells as **gaps**, not zero.
4. Set the Y-axis to 0–100% and the X-axis to 0–25 years (or 0–10 years for the zoom).
5. Highlight Generative AI and label the last real observation, `2026: 60.8%`; the GenAI curve should **stop**, not continue as a forecast.

## Comparability and interpretation caveats

- **This is not a strict apples-to-apples chart**. PC adoption is **household ownership**; Internet adoption measures **individual users of all ages**; GenAI measures **working-age adults (18–64)** using it for any purpose; broadband, social networking, and smartphone series generally measure **U.S. adults**. Methods and question wording also differ.
- **Launch anchors are not standardized.** Some technologies launched well before widespread commercial use. Moving a start year can change the visual or the `years to 50%` ranking materially.
- **AI is not “only two years old.”** Mass-market ChatGPT dates to **2022**. **Two years** refers only to the length of the **July 2024–July 2026 RPS survey record** in this project. More recent AI capabilities, such as reasoning and agents, have their own later development dates but are not measured separately here.
- **Use ≠ monetization or productivity.** Reporting AI usage does not establish paid spending, agent deployment, business process automation, economic value realized, GPU demand, or infrastructure utilization.
- **Short time series ≠ early S-curve location.** Fast initial adoption may bring a technology to the steep part of its adoption curve quickly. The eventual ceiling and inflection point for GenAI cannot be inferred confidently from the available points.
- **Adoption and financial bubbles can coexist.** A technology can continue diffusing during and after an investment bubble; the chart alone cannot establish whether AI equities are overvalued.
- **No invented data.** Missing years have not been backfilled. The figures are descriptive visualizations, not fitted logistic forecasts.

## Suggested caption for an investment presentation

**Chart title:** *Generative AI has the shortest observed history of major digital technologies*  
**Subtitle:** *U.S. technology adoption / ownership aligned by selected mass-market launch year; observed survey data only.*  
**Footnote:** *Sources: U.S. Census Bureau, ITU / World Bank via FRED, Bick et al. Real-Time Population Survey via FRED, and Pew Research Center. Definitions and population bases vary across series. GenAI observations begin July 2024; no forward estimates are shown.*

## Reproducibility / publication note

- The spreadsheet is a **curated research transcription**, not a verbatim government download. For updates, refresh the original publisher's figures and preserve the question wording, respondent age range, date, and data vintage.
- For third-party figures, datasets, and citations, check each publisher's copyright and re-use terms before external redistribution.
- Reference date: **November 4, 2026**.
