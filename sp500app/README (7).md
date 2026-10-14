# AI-Exposed Wage Pool | Rebuild Guide (2025 baseline → 2030)

> **用途**：在公司电脑上仅使用公开来源，从零复刻 `AI_Exposed_Wage_Pool_with_2030_Projection.xlsx` 的数据层与计算层。  
> **发布日期/数据快照**：本指南按 **2026-10-14** 可获得的版本固定来源：QCEW 2025 Annual、OEWS May 2025、Penn Wharton 2025、Census BTOS `202620`、BLS Employment Projections 2025–2035。  
> **模型衡量**：潜在受 AI 影响的**任务对应工资池**，**不是**失业人数、工资节约金额、已实现效率提升或 AI 公司收入。

---

## 0. 开始之前：准备 Excel 与文件夹（约 2 分钟）

先在公司电脑上建立下列目录，按文件名保存原始下载文件；**不需要把所有原始数据传到 GitHub**。

```text
ai-exposed-wage-pool/
├── README.md                    # 把本文件命名为 README.md 放在仓库根目录
├── model/
│   └── AI_Exposed_Wage_Pool.xlsx  # 公司电脑上自己复刻的模型
└── data/                        # 本地留存原始文件；大型文件不要提交到 GitHub
    ├── qcew_2025/
    ├── oews_may_2025/
    ├── wharton_2025/
    ├── census_btos_202620/
    └── bls_projection_2025_2035/
```

**复刻顺序**：`DATA Wharton` → `DATA QCEW` → `DATA OEWS` → `DATA Census` → `DATA BLS Projections` → `01 Industry Model` → `Projection Assumptions` → `Five-Year Projection`。

**重要质量提醒（请先看）**：先前提供的 2030 Excel 的 `DATA OEWS` 出现了 **13,457 条记录，但只有 7,805 个唯一的「行业 × 详细职业 SOC」组合**，其中 **5,652 个组合重复**。这属于工作簿构建过程的重复导入，不能当作真实的重复就业。**不要直接整表复制这个 Tab 作为公司版原始数据**。请按本指南从官方 OEWS 重新筛选，每个 NAICS × SOC 保留唯一记录，并用 Excel 的 `COUNTIFS` 检查。先前计算结果也应在去重后重新核对。

---

## 1. 下载 Penn Wharton：AI Exposure by Occupation（先做这个）

**官方网址**：

- 文章：[The Projected Impact of Generative AI on Future Productivity Growth](https://budgetmodel.wharton.upenn.edu/p/2025-09-08-the-projected-impact-of-generative-ai-on-future-productivity-growth/)
- [直接下载补充 CSV：supp_exposure_by_occupation.csv](https://budgetmodel.wharton.upenn.edu/legacy/csv/supp_exposure_by_occupation.csv)

**在官网怎么点**

1. 打开文章，滚动到 **Economic Activity Exposed to AI**。
2. 在 *Table 1* 前，点击 **“supplementary data file available for download”**。
3. 保存 `supp_exposure_by_occupation.csv` 到 `data/wharton_2025/`。
4. Excel → **Data → From Text/CSV** 打开文件（或直接打开 CSV）。

**应当看到的原始字段**

| 原始列 | 作用 |
|---|---|
| `Occupation Code` | 6 位 SOC，例如 `11-1011`；**按文本导入** |
| `Occupation Title` | 职业名称 |
| `Exposure to AI Automation (% of tasks)` | 例如 `47.75`，代表 47.75% |

**放入 Excel：`DATA Wharton`**

| 列 | 内容 | 从原始字段 |
|---|---|---|
| A | SOC code | `Occupation Code` |
| B | Occupation | `Occupation Title` |
| C | Wharton AI task exposure | 原始百分数 **÷ 100** |
| D | Wharton source | 文章网址 |

建议 `C2`：`=原始百分比单元格/100`，向下拖到全部职业（若原始单元格已经是 Excel 百分比，则不要再除以 100）。设置 C 列为百分比格式。

**检查**：784 个职业；A2 应为 `11-1011`，C2 为 **47.75%**；SOC 不要丢掉中间的 `-`。Excel 内 `DATA Wharton` 第 1 行是标题，数据第 2–785 行。

---

## 2. 下载 BLS QCEW：2025 年各 Industry 的真实 Annual Wages

**官方网址**：[QCEW Downloadable Data Files](https://www.bls.gov/cew/downloadable-data-files.htm)  
**直达 ZIP**：[2025 Annual — CSVs By Industry](https://data.bls.gov/cew/data/files/2025/csv/2025_annual_by_industry.zip)

**在官网怎么点**

1. 打开官方页面，找到 **NAICS-Based** 数据表。
2. 找到 **2025** 年那一行。
3. 选择 **CSVs → By Industry → Annual Averages**，注意**不是** By Area / Quarterly / All industries 单独 CSV。
4. 保存 `2025_annual_by_industry.zip` 到 `data/qcew_2025/` 并解压。
5. 解压后会有很多按行业拆分的 CSV（整个 ZIP 很大，无需全部加载进 Excel）。

**筛选方法（每个行业只取美国全国、私人部门的一行）**

例如，先打开 `2025.annual 11 NAICS 11 Agriculture, forestry, fishing and hunting.csv`：

| QCEW 字段 | 必须选择 |
|---|---|
| `area_fips` | `US000`（全美国；不是纽约州或某县） |
| `own_code` | `5`（私人部门） |
| `industry_code` | `11`（当前正在处理的行业） |
| `year` | `2025` |
| `qtr` | `A`（Annual） |
| `size_code` | `0`（All establishment sizes） |
| `total_annual_wages` | **直接复制这个字段**，不要乘 52 |

将每个目标 NAICS 行业重复执行，复制到 `DATA QCEW`：

| DATA QCEW 列 | 内容 |
|---|---|
| A | NAICS（文本，例如 `31-33`） |
| B | Industry |
| C | QCEW 2025 private annual wages ($) |
| D | 官网链接 |

**本模型需要的 19 个 NAICS（顺序很重要）**：

```text
11, 21, 22, 23, 31-33, 42, 44-45, 48-49, 51, 52,
53, 54, 55, 56, 61, 62, 71, 72, 81
```

**检查点**：

- NAICS 11: `60,088,224,865` 美元。
- NAICS 31-33: `1,125,542,676,332` 美元。
- 19 个行业工资合计：`10,515,645,524,505` 美元（约 **$10.516T**）。
- 美国私人部门“所有行业”控制总额：NAICS `10` / `US000` / `own_code=5` = `10,532,191,402,092` 美元（约 **$10.532T**）。
- 两者相差约 **$16.546B**，来自未分类 NAICS `99`，**不要把差额强行加进某个行业**。

> 注意：`31-33`、`44-45`、`48-49` 是官方合并的两位数 NAICS 行业代码，不要将其中各行业和合并行业同时相加，否则 double count。

---

## 3. 下载 BLS OEWS：Industry × Occupation 就业人数与工资（最容易出错的一步）

**官方网址**：[BLS Occupational Employment and Wage Statistics Tables](https://www.bls.gov/oes/tables.htm)  
**May 2025 下载入口**：**National industry-specific and by ownership → XLSX**  
**直接文件**：[May 2025 National Industry / Ownership ZIP](https://www.bls.gov/oes/special-requests/oesm25in4.zip)

**在官网怎么点**

1. 进入 OEWS Tables，找到最上方 **May 2025**。
2. 点击 **National industry-specific and by ownership → XLSX**。
3. 下载的是 ZIP，解压后找：
   - `natsector_M2025_dl.xlsx` — 19 个 NAICS 主要行业的职业结构。
   - `nat3d_owner_M2025_dl.xlsx` — 拆分教育、医院等所有制时使用。
4. 两个文件保存到 `data/oews_may_2025/`。

**原始表中的关键字段**：

| 原始列 | 内容 | 注意 |
|---|---|---|
| `NAICS` | 行业编码 | 与模型 19 个行业一致 |
| `NAICS_TITLE` | 行业名字 | 辅助核查 |
| `OWN_CODE` | 所有制 | 有些 aggregate sector 用特殊合并码，不能统一强行筛成 5 |
| `OCC_CODE` | SOC code | 必须是详细职业，例如 `11-1011` |
| `OCC_TITLE` | 职业名字 | — |
| `O_GROUP` | 职业层级 | **只保留 `detailed`**，排除 `total` / `major` / `minor` / `broad` |
| `TOT_EMP` | Employment | 遇 `**`/空值不要当成 0 |
| `A_MEAN` | Annual mean wage | 遇 `*`/`**`/空值不要当成 0 |

**将数据映射至 `DATA OEWS`**：

| 列 | 内容 | 方法 |
|---|---|---|
| A | NAICS | `NAICS`（需要映射至 19 行业） |
| B | Industry | `NAICS_TITLE` |
| C | SOC code | `OCC_CODE`，文本 |
| D | Occupation | `OCC_TITLE` |
| E | Employees | `TOT_EMP` |
| F | Annual mean wage ($) | `A_MEAN` |
| G | Known occupation wages ($) | Excel 公式 `E × F` |
| H | Wharton task exposure | 按 SOC `XLOOKUP` 到 `DATA Wharton` |
| I | Matched occupation wages ($) | 仅 H 有值时保留 G |
| J | Exposure-weighted wages ($) | `I × H` |
| K | Mapping status | `Matched` / `Missing OEWS wage` / `Missing Wharton SOC` |

**核心公式（从第 2 行开始，全部可拖拽）**：

`DATA OEWS!G2`:

```excel
=IF(OR(E2="",F2=""),"",E2*F2)
```

`DATA OEWS!H2`:

```excel
=IFERROR(XLOOKUP(C2,'DATA Wharton'!$A$2:$A$785,'DATA Wharton'!$C$2:$C$785),"")
```

`DATA OEWS!I2`:

```excel
=IF(OR(G2="",H2=""),"",G2)
```

`DATA OEWS!J2`:

```excel
=IF(OR(I2="",H2=""),"",I2*H2)
```

`DATA OEWS!K2`:

```excel
=IF(G2="","Missing OEWS wage",IF(H2="","Missing Wharton SOC","Matched"))
```

**所有制修正（必须单独处理）**：

1. **NAICS 61（教育）**：用 `nat3d_owner_M2025_dl.xlsx` 中 `NAICS=611000`、`OWN_CODE=5`、`O_GROUP=detailed` 的私人部门教育记录，**替换**原本可能混有公立教育的 sector 61 职业记录。
2. **NAICS 62（医疗）**：原 sector 62 混合了某些政府雇员。用 ownership 文件中 `NAICS=622000`、`OWN_CODE=2 或 3`、`O_GROUP=detailed` 的 **government hospitals**，按 SOC 扣掉对应的 government-hospital occupation wage bill（政府医院职业就业人数 × 其年平均工资）。严格复刻时，这个调整应先在职业层形成 adjusted wage bill，再应用 Wharton exposure；**不要只对 62 的总额扣一次**。
3. **48-49、71、72**：现有公开 sector 职业结构可能混有部分政府雇员；私人部门 QCEW 的年度工资数据仍然正确，但 OEWS 的 Exposure 权重是近似估计。记录 caveat，不要将其写成完美的 private-only match。

> 如果你只是先复刻一个可审计的 *first pass*，先完成除 61、62 外的 17 个行业；随后专门进行上述所有制修正。不要把不一致的行业假装已经对齐。

### 3A. 防重复 QA：务必执行

- 每个 **NAICS × 6-digit SOC** 应只有一条记录，不要同时保留 sector 与 3-digit 的重复数据，也不要重复粘贴相同明细。
- 在 `DATA OEWS` 的一个空白列（例如 L2）输入并向下拖：

```excel
=COUNTIFS($A$2:$A$8000,A2,$C$2:$C$8000,C2)
```

- 任何结果 **>1** 都需要调查、去重，**不能直接继续建模**。如果行数超过 8,000，扩展公式中的末行。
- 之前清洗后的逐职业审计文件具有 **7,805 条唯一 NAICS × SOC 记录**；这个数量用于核对本项目版本，未来 BLS 发布新版本后可能变化。
- 检查 NAICS 11 / SOC 11-1011：`Employees=170`，`Annual Mean Wage=$277,000`，应只出现 **一次**。

> 不建议用 Excel **Remove Duplicates** 不加判断地删所有数据：如果你合并了不同所有制的原始明细，先区分所有制和行业层级，再决定是否替换、扣减或去重。

---

## 4. 下载 Census BTOS：Industry AI Use（固定 202620 版本）

**官方网址**：[U.S. Census — Business Trends and Outlook Survey — Data Downloads](https://www.census.gov/hfp/btos/data_downloads)

1. 打开页面，找到 **BTOS Data Downloads**，进入 **Sector** 数据下载。
2. 下载 `Sector.xlsx`（在本项目原始上传中叫 `Sector (3).xlsx`），保存到 `data/census_btos_202620/`。
3. 打开工作表 **Response Estimates**（不要误用 `Response Standard Errors`）。
4. 找到最左边：`Sector`、`Question ID`、`Answer ID`、`Answer`。
5. 筛选 **Question ID = 7**、**Answer ID = 1**、**Answer = Yes**。
6. 找到表头 **202620** 这一列（当前项目选定的 survey wave）。**不要随手选更新的 202621 / 202622，否则无法复现这版 Excel**。
7. 按 19 个行业收集 AI Use 百分比，填入 `DATA Census` 的 C 列，确保是 Excel 真正的百分比数值，而非文本。

| DATA Census 列 | 内容 |
|---|---|
| A | 模型 NAICS（11、21、…、31-33、…） |
| B | Industry |
| C | 202620 企业使用 AI 的比例 |
| D | 下载链接 |

**行业代码 mapping**：Census 使用 `31` 对应模型 `31-33`；`44` 对应 `44-45`；`48` 对应 `48-49`。其余按照同一两位 NAICS 匹配。

**校验**：`Information (51) = 50.2%`，`Finance & Insurance (52) = 40.1%`，`Professional Services (54) = 42.7%`。

> **非常重要**：这个 Census 比例测量的是**使用 AI 的企业占比**，不是某个行业使用 AI 的员工比例，更不是工资可被节省的比例。主模型 Gross Exposure **不乘** Census AI Use；只在情景列演示乘法。
>
> 你另外下载的 `AI_Supplement_Table_2026.xlsx` 是另一套一次性补充调查，可用作 cross-check，但**不是**这版 `DATA Census` C 列的来源，不能混用不同调查时点。

---

## 5. 下载 BLS Employment Projections：2025–2035 → 2030

**官方网址**：[BLS Table 2.1 — Employment by Major Industry Sector](https://www.bls.gov/emp/tables/employment-by-major-industry-sector.htm)  
**直接下载**：[BLS 2025–2035 Industry Employment XLSX](https://www.bls.gov/emp/industry-employment/industry.xlsx)

1. 打开官网，看到 **Table 2.1 Employment by major industry sector, 2025 and projected 2035**。
2. 点击上方 **Other available formats: (XLSX)**。
3. 下载 Excel，保存到 `data/bls_projection_2025_2035/`。
4. 在官方数据中，只选与 19 个行业对应的行，避免把 `Total`、`Goods-producing` 等汇总行与其子行业同时相加。
5. 复制 **Employment 2025** 与 **Employment 2035**（单位：**thousands / 千人**）。

填入 `DATA BLS Projections`：

| 列 | 内容 |
|---|---|
| A | NAICS model code |
| B | BLS Industry Sector |
| C | 2025 Employment (000s) |
| D | 2035 Employment (000s) |
| E | Employment CAGR |
| F | Scope / mapping note |
| G | 官网链接 |

`E5` 公式（向下拖到 E23）：

```excel
=(D5/C5)^(1/(2035-2025))-1
```

**校验**：Construction (NAICS 23) 的就业人数约 `8,266.7`（千人，2025）→ `8,677.4`（千人，2035）。

> BLS 提供的是**就业人数预测**，不是工资增长预测，也不是 AI adoption 预测。我们将官方 2025–2035 CAGR 作为 2026–2030 的机械插值。教育/医疗等行业一定检查与私人部门工资口径是否匹配。

---

## 6. 先在 `01 Industry Model` 完成 2025 Baseline

在主表中，A5:A23 = 19 个 NAICS；B5:B23 = 对应的行业名称；C–K 使用 `XLOOKUP`、`SUMIF`、简单乘除法。**不使用 `LET`、`INDEX/MOD` 或复杂数组**。

以下是第 5 行的示例公式，锁定数据范围、保持当前行相对引用；向下复制到第 23 行。**将 `8000` 替换为你清洗后 `DATA OEWS` 的真实末行**（如果是 7,805 条，从标题下第二行开始，最后一行为 7,806）。

`C5` — QCEW industry wage bill（$bn）：

```excel
=IFERROR(XLOOKUP(A5,'DATA QCEW'!$A$2:$A$20,'DATA QCEW'!$C$2:$C$20)/1000000000,"")
```

`D5` — OEWS known wage bill（$bn）：

```excel
=SUMIF('DATA OEWS'!$A$2:$A$8000,A5,'DATA OEWS'!$G$2:$G$8000)/1000000000
```

`E5` — SOC / Wharton matched wage bill（$bn）：

```excel
=SUMIF('DATA OEWS'!$A$2:$A$8000,A5,'DATA OEWS'!$I$2:$I$8000)/1000000000
```

`F5` — Exposure-weighted wage bill（$bn）：

```excel
=SUMIF('DATA OEWS'!$A$2:$A$8000,A5,'DATA OEWS'!$J$2:$J$8000)/1000000000
```

`G5` — Wharton match coverage：

```excel
=IF(D5=0,"",E5/D5)
```

`H5` — AI Task Exposure（工资加权）：

```excel
=IF(E5=0,"",F5/E5)
```

`I5` — **Gross AI-Exposed Wage Pool**（$bn）：

```excel
=C5*H5
```

`J5` — Census BTOS 企业采用率（仅供参考）：

```excel
=IFERROR(XLOOKUP(A5,'DATA Census'!$A$2:$A$20,'DATA Census'!$C$2:$C$20),"")
```

`K5` — **Illustrative** firm-use-adjusted proxy（不是 savings）：

```excel
=I5*J5
```

模型的经济逻辑：

```text
职业工资 = Employment × Average Annual Wage
行业 Exposure Share = Σ(职业工资 × 职业任务暴露率) / Σ(有 Wharton 匹配的职业工资)
Gross Exposed Wage Pool = QCEW 真实行业工资 × 行业 Exposure Share
```

**不应**把 OEWS occupation payroll 再加到 QCEW industry payroll 上；OEWS 只负责计算各行业的加权暴露比例。

---

## 7. 建 2026–2030 Five-Year Projection

### 7A. `Projection Assumptions`：先输入并标注为假设

| Cell | Input | 默认值 | 来源 |
|---|---|---:|---|
| C5 | Nominal wage growth per employee per year | **3.0%** | **Assumption**，不是官方预测 |
| C6 | AI-using firms increase per year | **+5 percentage points** | **Assumption**，不是官方预测 |
| C7 | Task exposure increase per year | **0 percentage points** | **Assumption**，保持 Wharton 2025 暴露率 |

`Projection Assumptions` 第 13–31 行可分别设置 19 个行业覆盖全局参数，例如 C13=`=$C$5`、D13=`=$C$6`、E13=`=$C$7`，拖拽向下，之后再对个别行业覆盖输入。

### 7B. `Five-Year Projection`：每个行业 2025–2030 六行

表头在第 16 行，A–L 列分别为：NAICS、Industry、Year、BLS Jobs CAGR、Wage/Worker CAGR、Wage Bill、Baseline Exposure、Exposure change、Projected Exposure、Gross Exposure、AI-Using Firms、Use-Adjusted Proxy。

以 NAICS 11 为例：

- Row 17 = 2025 **base year**，F17 从 `01 Industry Model!C5` lookup，G17 从 `01 Industry Model!H5` lookup。
- Row 18 = 2026：

```excel
=F17*(1+D18)*(1+E18)
```

  F18 表示 2026 预测 Wage Bill；D18 是官方就业 CAGR，E18 是名义人均工资增长**假设**。

- 2026 的任务暴露率（I18）：

```excel
=MIN(1,MAX(0,G18+(C18-2025)*H18))
```

- 2026 Gross Exposure（J18）：

```excel
=F18*I18
```

- 2026 的 BTOS Use（K18）从 `DATA Census` 的 202620 横截面读取，**为一个时点值，并非 2026 全年平均采用率**。
- 2027 的企业 AI Use（K19）基于前一年线性增加（C6，+5pp），最多 100%：

```excel
=MIN(1,K18+XLOOKUP(A19,'Projection Assumptions'!$A$13:$A$31,'Projection Assumptions'!$D$13:$D$31))
```

- 2027 Gross Exposure（J19）：`=F19*I19`；Use-Adjusted Proxy（L19）：`=J19*K19`。

向下复制时记得**每个行业第一行（2025 基期）使用 XLOOKUP，而每个行业第二行（2026）开始才使用上一年工资滚动公式**，不要让新行业的 2025 工资错误地接上上一个行业的 2030 年。

---

## 8. 最后五个检查：通过后再给老板

- [ ] **QCEW Scope**：`US000`、`own_code=5`、`qtr=A`、2025；19 行业合计约 **$10.516T**；全国私人部门总额约 **$10.532T**，差额属于 NAICS 99。
- [ ] **OEWS No Duplicates**：只有 detailed SOC；NAICS × SOC 重复检查应全部为 1；教育、医院私有/公有口径已单独处理。
- [ ] **Wharton Match**：784 个 SOC；Exposure 格式为 0–100%；缺失职业**不当成 0% exposure**，而是从 matched denominator 排除。
- [ ] **Census Fixed Vintage**：使用 `202620`、Q7 / Yes，行业 51=50.2%、52=40.1%、54=42.7%；清楚标注**企业采用率，不是员工采用率**。
- [ ] **Projection Assumptions**：BLS 只提供就业 CAGR；3% wage growth、5pp annual adoption、0pp exposure change 都明示为 assumptions；2025 Base / 2026 Reference / 2030 Forecast 不混用。

### 口径局限（研究报告里应写明）

1. Wharton exposure 是任务可以受 AI 自动化影响的**潜力**，不是现实可节省的工资比例。
2. QCEW wages 与 OEWS mean annual wages 来自不同统计项目，**不能直接相加**；使用 OEWS 只是为了推导职业工资权重。
3. Census BTOS 测量的是企业使用率；直接乘工资池只是一种 illustrative proxy，没有因果识别。
4. BLS 就业预测不是由 AI adoption 驱动的预测；如果要归因“AI 创造了多少生产率”，后续还需要单独的 real productivity / time saved / realization assumptions。

---

## 9. 把 README 放上 GitHub，方便公司电脑看

**最简单（GitHub 网页操作）**：

1. 在自己的 GitHub 建一个 repository，例如 `ai-wage-exposure-rebuild`。
2. 下载本文件，**重命名为 `README.md`**。
3. 在 GitHub repository 页面点击 **Add file → Upload files**。
4. 拖进去 `README.md`，点击 **Commit changes**。
5. GitHub 仓库首页就会自动渲染 README。打开公司电脑浏览器进入这个 repository，逐步骤照着做即可。

**建议只提交** `README.md`、不含公司机密的模板与必要说明。不要将 UBS 内部研究、内部数据库、客户资料、公司电脑信息提交到个人或公开仓库；如果公司不允许访问个人 GitHub，就使用公司批准的 GitLab/知识库。**QCEW 原始 ZIP 超过 100 MB，GitHub 普通上传也不合适**，官方源链接已经在本文件中。

---

## 10. 完整官方数据源索引（便于重新下载）

| Source | 数据文件/版本 | 下载链接 |
|---|---|---|
| Penn Wharton | `supp_exposure_by_occupation.csv`, 2025 | https://budgetmodel.wharton.upenn.edu/legacy/csv/supp_exposure_by_occupation.csv |
| BLS QCEW | `2025_annual_by_industry.zip`, Annual | https://data.bls.gov/cew/data/files/2025/csv/2025_annual_by_industry.zip |
| BLS OEWS | `oesm25in4.zip`, May 2025 national industry & ownership | https://www.bls.gov/oes/special-requests/oesm25in4.zip |
| Census BTOS | `Sector.xlsx`, **column 202620** | https://www.census.gov/hfp/btos/data_downloads |
| BLS Employment Projections | `industry.xlsx`, 2025–2035 | https://www.bls.gov/emp/industry-employment/industry.xlsx |

> 如果官网未来更新，优先锁定本模型对应的旧版本（**2025 Annual / May 2025 / 202620 / 2025–2035**），只有在主动决定重新基准化时才一起更新年份与调查期。
