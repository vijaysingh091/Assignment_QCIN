# Data Analyst Assignment

> **Assignments:** Indian Market Surveillance Analysis (Assignment 1) & PMSGMBY Beneficiary Analysis (Assignment 2)
> **Tools Used:** Python 3, Jupyter Notebook, SQLite, Microsoft Excel

---

## Table of Contents

1. [Repository Structure](#repository-structure)
2. [Assignment 1 — Indian Market Surveillance](#assignment-1--indian-market-surveillance-analysis)
   - [Background & Dataset](#background--dataset)
   - [Questions Solved](#questions-solved-assignment-1)
   - [Key Findings](#key-findings-assignment-1)
   - [How to Run](#how-to-run-assignment-1)
   - [Output Files](#output-files-assignment-1)
3. [Assignment 2 — PMSGMBY Beneficiary Analysis](#assignment-2--pmsgmby-beneficiary-analysis)
   - [Background & Dataset](#background--dataset-1)
   - [Questions Solved](#questions-solved-assignment-2)
   - [Key Findings](#key-findings-assignment-2)
   - [How to Run](#how-to-run-assignment-2)
   - [Output Files](#output-files-assignment-2)
4. [Assumptions & Limitations](#assumptions--limitations)
5. [Dependencies](#dependencies)

---

## Repository Structure

```
.
├── README.md                            ← This file
│
├── Assignment_1_DATA_SETS.xlsx          ← Raw data (Assignment 1)
   └── Outputs_Generated/
       ├── ── Assignment 1 Outputs ──
       ├── Q1_Standardized_Oil_Types.xlsx
       ├── Q2_Descriptive_Analysis.xlsx
       ├── Q3_PassFail_Results.xlsx
       ├── Q4_Mismatch_Report.xlsx
       ├── Q5_Master_Dataset.xlsx
│
├── Assignment_2_DATA_SETS.xlsx          ← Raw data (Assignment 2)
   ├── Edible_Oil_Assignment.ipynb          ← Assignment 1 — Python notebook (Q1–Q5)
   ├── PMSGMBY_Analysis.ipynb               ← Assignment 2 — Python notebook (A1–D4)
   ├── PMSGMBY_SQL_Queries.sql              ← Assignment 2 — All SQL queries (S1–S8)
   ├── PMSGMBY_SQL_Runner.py                ← Assignment 2 — Script to execute SQL & export results
   ├── PMSGMBY_Excel_Builder.py             ← Assignment 2 — Script that built the Excel file
   └── Outputs_Generated/
       ├── ── Assignment 2 Python Outputs ──
       ├── A1_Missing_Data.png
       ├── A2_Pipeline_Days_Histogram.png
       ├── A3_Outlier_Treatment.png
       ├── A5_Summary_Statistics.xlsx
       ├── B1_Market_Cost_Gap.png
       ├── B2_Interest_Rate_Breach_By_State.png
       ├── B3_DISCOM_SLA.xlsx
       ├── B4_Bill_Savings.png
       ├── B5_DISCOM_Incentive.png
       ├── C1_Pipeline_Funnel.png
       ├── C2_Pipeline_By_State.png
       ├── C3_ALMM_Compliance.png
       ├── C4_Grievance_Analysis.png
       ├── C5_Regression_Coefficients.png
       ├── D1_Beneficiary_Share.png
       ├── D2_Lorenz_Curve.png
       ├── D3_Performance_Heatmap.png
       ├── D4_Payback_Period.png
       ├── PMSGMBY_Enriched_Dataset.xlsx
       │
       ├── ── Assignment 2 SQL Outputs ──
       ├── PMSGMBY_SQL_Results.xlsx
       │
       └── ── Assignment 2 Excel Submission ──
           └── Candidate_Name_Assignment_2.xlsx
```

---

## Assignment 1 — Indian Market Surveillance Analysis

### Background & Dataset

This assignment analyses data from the **Indian Market Quality Assessment Project**, which monitors the quality of edible oils sold across Indian markets. Data is sourced from two field forms:

| Form | Sheet Name | Rows | Description |
|------|-----------|------|-------------|
| C1 (Collection Form) | `Sheet_C1 FORM` | 400 | Sampler-reported data — oil type, city, shop type, dates |
| T1 (Testing Form) | `Sheet_T1 FORM` | 400 | Lab test results — chemical parameters, Pass/Fail |
| Evaluation Criteria | `Evaluation Criteria_Parameters` | — | Regulatory thresholds for each test parameter |

**Key challenge:** The C1 and T1 datasets must be joined via a common sample code to build the full picture of each sample's journey from collection to lab result. Only 51 of 400 samples have matching records across both forms — this itself is flagged as an operational insight.

---

### Questions Solved — Assignment 1

#### Q1 — Data Cleaning & Standardisation
**What was done:** Standardised the `Q17. Type of Oil` column in C1 which contained inconsistent entries like `"soyabeanoil"`, `"REFINED VEGETBLE OIL"`, `"palmoil"`.

**Standardisation logic (3-step pipeline):**
1. **Pre-processing:** Strip whitespace, collapse multiple spaces, convert to lowercase
2. **Space removal:** Remove all spaces to catch merged-word variants (`soyabeanoil` → `soybeanoil`)
3. **Fuzzy matching:** Use `difflib.get_close_matches` with an 80% similarity cutoff against a predefined canonical list of 12 oil types
4. **Manual overrides:** A `MANUAL_OVERRIDES` dict handles known edge cases that fuzzy matching misses (e.g. `"REFINEDVEGETBLE OIL"`)

**Configurable:** The 80% cutoff can be raised (stricter) or lowered (more permissive). The canonical list and manual overrides can be extended without changing any logic.

**Output:** `Q1_Standardized_Oil_Types.xlsx` — two sheets: mapping of all before/after values, and full C1 data with the new standardised column.

---

#### Q2 — Descriptive Analysis
**What was done:**
- Computed oil-wise sample counts with percentage share across all 400 C1 records
- Identified top 5 oil types by volume
- Produced a city-wise distribution breakdown for those top 5 oils

**Output:** `Q2_Descriptive_Analysis.xlsx` — two sheets: `Oil_Sample_Counts` and `City_Distribution_Top5`

---

#### Q3 — Pass / Fail Determination
**What was done:** Built a parameter-driven rule engine that evaluates each T1 sample against regulatory thresholds and returns a `Pass_Fail` status and a detailed `Failure_Reason`.

**Rule categories:**
| Category | Parameters | Logic |
|----------|-----------|-------|
| Heavy metals | Lead, Copper, Arsenic, Tin, Cadmium, Mercury | Numeric value > max (5 PPM) → FAIL |
| Aflatoxins | Aflatoxin B1, Total Aflatoxins | Numeric value > max (1 / 5 µg/kg) → FAIL |
| Physical | Saponification, Iodine, Unsaponifiable, Acid value | Numeric value > max threshold → FAIL |
| Qualitative | Polybromide, Argemone, Mineral oil tests | `Positive` → FAIL |
| Non-detect tokens | BLQ, ND, LOQ, `<` prefix | Always PASS regardless of column |

**Adapting to regulatory changes:** All thresholds live in a single `THRESHOLDS` dictionary at the top of the cell. Change one number and it propagates to the entire dataset — no other code needs to be touched.

**Output:** `Q3_PassFail_Results.xlsx` — three sheets: full T1 with Pass/Fail, FAIL records only, and a summary count.

---

#### Q4 — Sampler vs Lab Mismatch Detection
**What was done:** Compared oil types reported by the field sampler (`Q6`) vs identified by the lab (`Q8`) after standardising both columns using the same Q1 logic.

**Mismatch classification:**
- **Match:** Both standardised names are identical
- **Minor Mismatch:** Different names but same oil family (e.g. `Sunflower Oil` vs `Sunflower Seed Oil`; `Refined Vegetable Oil` vs `Multi-Source Edible Oil`)
- **Major Mismatch:** Completely different oil families (e.g. `Olive Oil` reported, `Mustard Oil` found)

**Oil families** are defined in a configurable `OIL_FAMILIES` dictionary — adding new families requires no logic changes.

**Output:** `Q4_Mismatch_Report.xlsx` — four sheets: all samples classified, mismatch records only, summary by oil type, summary by city.

---

#### Q5 — C1 ↔ T1 Interconnection & Key Insights
**What was done:** Joined C1 and T1 on `Sample Unique Code` to build a master dataset, then analysed operational delays and derived cross-dataset insights.

**Join key:** `Q41. Sample Unique Code` (C1) ↔ `Sample Unique Code` (T1)

**Key findings only possible via join:**

| Insight Type | Finding |
|---|---|
| Operational | City-level FAIL rates — City comes from C1; Pass/Fail from T1 |
| Analytical | Shop type vs FAIL rate — shop type in C1, test results in T1 |
| Analytical | Major mismatch + FAIL overlap — samples where lab found a different oil AND it failed testing |

**Output:** `Q5_Master_Dataset.xlsx` — master joined dataset + three insight sheets.

---

### Key Findings — Assignment 1

- The oil name standardisation corrected **~30% of C1 records** that had spelling or formatting issues
- Major mismatches (completely different oil identified by lab vs sampler) represent a significant data quality concern and a potential adulteration signal
- Only **51 of 400 samples** have matching C1 and T1 records — suggesting either incomplete data entry or a systemic gap in sample tracking
- Samples with major oil type mismatches show a higher overlap with FAIL status, suggesting mislabelling may correlate with quality issues

---

### How to Run — Assignment 1

**Prerequisites:**
```
pip install pandas openpyxl difflib
```

**Steps:**
1. Place `Assignment_1_DATA_SETS.xlsx` in the same folder as `Edible_Oil_Assignment.ipynb`
2. Open Jupyter: `jupyter notebook`
3. Open `Edible_Oil_Assignment.ipynb`
4. Run all cells top to bottom using **Shift+Enter** or **Kernel → Restart & Run All**
5. All 5 output Excel files will be generated in the same folder

---

### Output Files — Assignment 1

| File | Question | Contents |
|------|---------|----------|
| `Q1_Standardized_Oil_Types.xlsx` | Q1 | Before/after oil name mapping; full C1 with standardised column |
| `Q2_Descriptive_Analysis.xlsx` | Q2 | Oil sample counts; city-wise distribution for top 5 oils |
| `Q3_PassFail_Results.xlsx` | Q3 | T1 with Pass_Fail + Failure_Reason; FAIL-only records; summary |
| `Q4_Mismatch_Report.xlsx` | Q4 | All samples classified; mismatch records; by-oil & by-city summaries |
| `Q5_Master_Dataset.xlsx` | Q5 | Joined C1+T1 master; city fail rate; shop type fail rate; shelf life analysis |

---

---

## Assignment 2 — PMSGMBY Beneficiary Analysis

I Tried to do things via excel formulas but Excel was creating some problems. I did the same things via Python but I wrote the Excel formulas that would be used 

### Background & Dataset

**PMSGMBY (PM Surya Ghar Muft Bijli Yojana)** is a Government of India scheme providing rooftop solar subsidies to residential beneficiaries. This assignment analyses 5,000 application records covering **20 states**, **36 DISCOMs**, and applications filed between January 2024 and May 2025.

**Dataset:** `Assignment_2_DATA_SETS.xlsx` → Sheet: `Dataset` — 5,000 rows × 36 columns

**Key scheme rules embedded in analysis:**

| Rule | Detail |
|------|--------|
| CFA (Central Financial Assistance) | 60% of benchmark cost for ≤2 kW; 40% for the 3rd kW; capped at ₹78,000 |
| Benchmark cost | ₹50,000/kW (general states); ₹55,000/kW (special category — hilly & NE states) |
| DBT SLA | CFA must reach beneficiary within **30 days** of commissioning |
| Bank loan mandate | Interest rate must be **≤7%** for collateral-free solar loans |
| Grievance SLA | Resolution within **30 days** of raising |
| ALMM compliance | Installations must use equipment from the Approved List of Models and Manufacturers |

---

### Questions Solved — Assignment 2

#### Section A — Data Cleaning & Exploration (A1–A5)

| Q | What was done |
|---|---|
| A1 | Dataset info: shape (5,000 × 36), column dtypes, null counts, % missing per column — visualised as a bar chart |
| A2 | Parsed all 5 date columns to datetime; engineered `Total_Pipeline_Days = CFA_Credit_Date − Application_Date`; plotted histogram for commissioned applications |
| A3 | IQR-based outlier detection on `Market_Cost_Rs` (justified: right-skewed data; IQR is more robust than Z-score); treated via Winsorisation (capping at bounds, not dropping rows); before/after boxplots shown |
| A4 | Two integrity checks: (1) CFA > ₹78,000 cap — flagged records; (2) Actual_Monthly_Units > 1.5× Expected — flagged as generation anomalies |
| A5 | Summary statistics table (mean, median, std, min, max, P25, P75) for 4 key metrics: TFC_Processing_Days, CFA_Credit_Days_After_Commission, Monthly_Bill_Savings_Rs, Satisfaction_Score |

---

#### Section B — Financial Architecture Analysis (B1–B5)

| Q | What was done |
|---|---|
| B1 | Computed Market Cost Gap % = (Market − Benchmark×Size) / (Benchmark×Size) × 100; found % of applications where market cost exceeds benchmark by >20%; histogram visualisation |
| B2 | Among loan applicants, computed % charged above the mandated 7% interest rate; state-level bar chart sorted by breach % |
| B3 | DBT SLA compliance: % receiving CFA within 30 days; DISCOM-level table with % within SLA and avg lag days; top 5 worst-performing DISCOMs identified; exported to `B3_DISCOM_SLA.xlsx` |
| B4 | Average monthly bill savings segmented by (a) system size in kW and (b) beneficiary category; grouped bar chart; assessed whether BPL/SC-ST households benefit proportionally more |
| B5 | DISCOM Incentive Utilisation Rate = Utilised/Received; ranked horizontal bar chart; DISCOMs below 80% flagged; Pearson correlation between utilisation rate and avg TFC processing days computed and scatter-plotted |

---

#### Section C — Implementation Pipeline Analysis (C1–C5)

| Q | What was done |
|---|---|
| C1 | Funnel analysis: 5,000 → TFC (4,518) → Installation (4,006) → Commission (3,675) → CFA Credit (3,675); drop-off % annotated at each stage; horizontal funnel bar chart |
| C2 | Average days per pipeline stage by state; state with longest total average pipeline time identified; stacked bar chart |
| C3 | ALMM compliance rate by state; states below 75% highlighted; Pearson correlation between ALMM compliance % and average Vendor_Rating computed and scatter-plotted |
| C4 | Grievance type distribution (pie chart); resolution rate per type; % resolved within 30-day SLA per type; bar chart of resolution rates with 60% threshold line |
| C5 | Linear regression to predict Satisfaction_Score using 5 features (System_Size_kW, Monthly_Bill_Savings_Rs, Vendor_Rating, Grievance_Resolved, Total_Pipeline_Days); features standardised for comparable coefficients; R², coefficients, and most influential factor reported |

---

#### Section D — Equity & Impact Analysis (D1–D4)

| Q | What was done |
|---|---|
| D1 | Share of each Beneficiary_Category in dataset; combined BPL + SC/ST + Women-headed share compared against assumed national population share of 55%; over/under-representation assessed |
| D2 | State-wise commissioned installation counts; Gini coefficient computed; Lorenz curve plotted; top 3 states accounting for >50% of all installations identified |
| D3 | Performance Efficiency Ratio = Actual/Expected monthly units; computed for commissioned records; segmented by State × System_Size_kW; heatmap showing segments below 85% threshold |
| D4 | Payback period = (Market_Cost − CFA) / Monthly_Bill_Savings; computed for all commissioned beneficiaries; capped at 360 months; boxplots by Beneficiary_Category; BPL vs General median payback compared |

---

#### Section S — SQL Analysis (S1–S8)

**Dialect used: SQLite** (compatible with PostgreSQL/DuckDB with minor date function changes noted inline)

| Query | What it does |
|---|---|
| S1 | State-wise application count, commissioned count, conversion rate %; states with ≥50 applications; ordered by conversion rate descending |
| S2 | DISCOMs where avg CFA credit lag exceeds 60 days (2× the SLA); returns DISCOM, state, avg days, breach count |
| S3 | Window function (DENSE_RANK) to rank vendors within each state by avg rating; returns top 3 vendors per state |
| S4 | CTE-based month-wise application volume + cumulative total; identifies month with highest single-month growth rate using LAG() |
| S5 | Grievance summary by type: count, resolved count, resolution rate, avg TAT; CASE WHEN flags `Requires Attention` (resolution <60% or avg TAT >30 days) vs `On Track` |
| S6 | CTE joining incentive gap (Received − Utilised) with avg TFC processing days per DISCOM; top 10 by unutilised amount |
| S7 | CASE WHEN bucketing into Compliant (≤7%), Marginal (7–9%), Non-Compliant (>9%); multi-level aggregation with state with highest count per bracket |
| S8 | Complex WHERE filter for high-value underserved: BPL or SC/ST, 3 kW system, no loan approved, CFA <50% of market cost; count and avg out-of-pocket cost by state |

**To run SQL queries:**
```bash
python PMSGMBY_SQL_Runner.py
```
This loads the dataset into an in-memory SQLite database, runs all 8 queries, prints results, and exports each to a separate sheet in `PMSGMBY_SQL_Results.xlsx`.

---

#### Section E — Advanced Excel (E1–E7)

The Excel submission file `Candidate_Name_Assignment_2.xlsx` contains the following sheets — all built with **native Excel formulas** (not pre-calculated values):

| Sheet | What it contains |
|---|---|
| `Pivot_State_Summary` | PivotTable: Total applications, % Commissioned, Avg Bill Savings, Avg Satisfaction, ALMM Compliance % by state; conditional formatting highlights % Commissioned <60% in red |
| `CFA_Check` | Formula column `Calculated_CFA` using `=MIN(IF(size≤2, 0.6×B×size, 0.6×B×2+0.4×B×(size-2)), 78000)`; `CFA_Error = Calculated − Sanctioned`; `Flag` column using `=IF(ABS(error)>500,"⚠ ERROR","OK")`; conditional formatting on flag column |
| `Pipeline_Dashboard` | 3 charts: (a) Bar chart of avg pipeline stage days by top 10 states using `AVERAGEIF`, (b) Pie chart of grievance type distribution using `COUNTIF`, (c) Line chart of monthly trend using `COUNTIFS`; slicer on Beneficiary_Category |
| `Vendor_Analysis` | Formula-based aggregation using `COUNTIF`, `AVERAGEIF`, `COUNTIFS` — no pivot table; sorted by total installations descending |
| `Loan_Compliance` | All rows where Has_Loan=TRUE and rate >7%; `VLOOKUP` formula pulling State_Code from `State_Ref` sheet; `COUNTIFS` violations summary by state |
| `State_Ref` | Reference lookup table: State name → State code (used by VLOOKUP in Loan_Compliance) |
| `Payback_Model` | Payback formula `=IF(savings>0,(cost-cfa)/savings,"N/A")`; Payback_Category using nested `IF`; colour-coded conditional formatting (green/yellow/red); PivotTable of count and avg payback by category |
| `Macro_Report` | VBA macro (`FormatAndExport`) that: (a) auto-fits columns, (b) applies header formatting, (c) exports sheet as PDF named with today's date; button assigned to macro |

> **Note:** To enable the macro, save the file as `.xlsm` (Excel Macro-Enabled Workbook) → Alt+F11 → Insert Module → paste macro code from the `Macro_Report` sheet → close VBA editor.

---

### Key Findings — Assignment 2

- **DBT SLA breach is widespread:** A significant portion of commissioned beneficiaries did not receive CFA within the mandated 30 days, with several DISCOMs averaging over 60 days — double the SLA
- **Market costs exceed benchmark:** A notable % of applications show market costs more than 20% above the ₹50,000/kW benchmark, suggesting the benchmark may need revision
- **Interest rate violations exist:** Multiple states show borrowers being charged above the mandated 7% interest rate on solar loans, indicating financial compliance gaps
- **Geographic concentration:** Installation distribution across states shows a high Gini coefficient, with just 3 states accounting for the majority of all commissioned installations — the scheme is not uniformly reaching all states
- **Marginalised group representation:** The combined BPL + SC/ST + Women-headed share in the dataset is compared against the national 55% benchmark to assess inclusion effectiveness
- **Pipeline bottleneck:** The biggest time sink in the implementation pipeline is between TFC approval and installation, suggesting vendor capacity or logistics as the primary operational bottleneck
- **Vendor quality link:** A positive correlation exists between ALMM compliance rates and vendor ratings, confirming that compliant vendors tend to be higher-quality operators

---

### How to Run — Assignment 2

#### Python Notebook
```bash
# Prerequisites
pip install pandas numpy matplotlib seaborn scikit-learn openpyxl

# Run
jupyter notebook PMSGMBY_Analysis.ipynb
# Kernel → Restart & Run All
```

#### SQL Runner
```bash
# Prerequisites (already covered above)
python PMSGMBY_SQL_Runner.py
# Outputs: PMSGMBY_SQL_Results.xlsx
```

#### Excel File
- Open `Candidate_Name_Assignment_2.xlsx` (rename with your name)
- All formulas are live in the cells
- To activate the macro: save as `.xlsm` → Alt+F11 → paste macro → run

---

### Output Files — Assignment 2

| File | Type | Contents |
|------|------|----------|
| `A1_Missing_Data.png` | Chart | Missing data % by column |
| `A2_Pipeline_Days_Histogram.png` | Chart | Total pipeline days distribution |
| `A3_Outlier_Treatment.png` | Chart | Market cost boxplots before/after Winsorisation |
| `A5_Summary_Statistics.xlsx` | Excel | Summary stats for 4 key metrics |
| `B1_Market_Cost_Gap.png` | Chart | Market cost gap % histogram |
| `B2_Interest_Rate_Breach_By_State.png` | Chart | State-wise interest rate breach bar chart |
| `B3_DISCOM_SLA.xlsx` | Excel | DISCOM-level DBT SLA compliance table |
| `B4_Bill_Savings.png` | Chart | Bill savings by size & category |
| `B5_DISCOM_Incentive.png` | Chart | DISCOM incentive utilisation + correlation scatter |
| `C1_Pipeline_Funnel.png` | Chart | Application pipeline funnel with drop-off % |
| `C2_Pipeline_By_State.png` | Chart | Stacked bar of pipeline stage days by state |
| `C3_ALMM_Compliance.png` | Chart | ALMM compliance by state + vendor rating scatter |
| `C4_Grievance_Analysis.png` | Chart | Grievance pie + resolution rate bar |
| `C5_Regression_Coefficients.png` | Chart | Linear regression coefficients for Satisfaction_Score |
| `D1_Beneficiary_Share.png` | Chart | Beneficiary category share vs national benchmark |
| `D2_Lorenz_Curve.png` | Chart | Lorenz curve of state-wise installations |
| `D3_Performance_Heatmap.png` | Chart | Performance ratio heatmap (State × System Size) |
| `D4_Payback_Period.png` | Chart | Payback period boxplots by beneficiary category |
| `PMSGMBY_Enriched_Dataset.xlsx` | Excel | Full dataset with all computed columns added |
| `PMSGMBY_SQL_Results.xlsx` | Excel | All 8 SQL query results (one sheet per query) |
| `Candidate_Name_Assignment_2.xlsx` | Excel | Final Excel submission with all E1–E7 sheets |

---

## Assumptions & Limitations

### Assignment 1
| Assumption / Limitation | Detail |
|---|---|
| Fuzzy match cutoff | Set at 80% similarity. Values that don't match any canonical name are title-cased and returned as-is — these should be manually reviewed |
| Regulatory thresholds | Thresholds in Q3 are based on the Evaluation Criteria sheet provided. If the sheet is updated, only the `THRESHOLDS` dict needs to change |
| C1–T1 join | Only 51 of 400 records matched on `Sample Unique Code`. Unmatched records are excluded from Q5 analysis. This low overlap rate is flagged as a data quality issue |
| Pass/Fail on missing data | If a parameter value is null for a sample, that parameter is skipped (not flagged as fail). Samples with all parameters null will incorrectly show as PASS |

### Assignment 2
| Assumption / Limitation | Detail |
|---|---|
| Outlier treatment | Winsorisation (capping) chosen over deletion to preserve all 5,000 rows for downstream analysis. This may slightly distort statistics at the extremes |
| Regression model | C5 uses a simple OLS linear regression. R² is expected to be low given the complexity of satisfaction drivers. The model is used for interpretation, not prediction |
| Gini coefficient | Computed on state-level aggregated counts, not individual beneficiaries. This measures geographic concentration, not income inequality |
| National population benchmark | The 55% marginalised population share used in D1 is as specified in the assignment. Actual figures may differ by source and year |
| SQL dialect | All SQL queries use SQLite syntax. For PostgreSQL, replace `strftime('%Y-%m', date)` with `TO_CHAR(date, 'YYYY-MM')`. For MySQL, use `DATE_FORMAT(date, '%Y-%m')` |
| Excel macros | The VBA macro in E7 requires saving the file as `.xlsm`. The `.xlsx` version will open normally but the macro button will not function |

---

## Dependencies

```
Python 3.8+
pandas >= 1.5
numpy >= 1.23
matplotlib >= 3.6
seaborn >= 0.12
scikit-learn >= 1.1
openpyxl >= 3.0
difflib (standard library)
sqlite3 (standard library)
jupyter notebook or jupyterlab
```

Install all at once:
```bash
pip install pandas numpy matplotlib seaborn scikit-learn openpyxl jupyter
```

---
