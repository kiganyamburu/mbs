# HW Set #6 – MBS Prepayment Speed & Yield Analysis

A Python-based analysis toolkit for **Mortgage-Backed Securities (MBS)** prepayment and yield study. This project generates charts, tables, and a formatted Word document summarizing findings across four sections of the homework assignment.

---

## Project Structure

```
mbs/
├── hw6_solution.py          # Main solution script (charts + Word doc)
├── populate_excel.py        # Populates Excel workbooks with analysis data
├── add_charts_to_q3.py      # Adds charts into the Q3 Excel file (openpyxl)
├── add_charts_win32.py      # Adds charts into the Q3 Excel file (win32com)
├── sec2_chart.py            # Section 2: CPR vs PMMS chart (CUSIP 31418C5Z3)
├── update_sec2.py           # Updates Section 2 data in Excel
├── verify.py                # Verification / sanity checks
├── charts/                  # Output directory for all generated chart PNGs
├── HW Q3 Problem_Agency MBS Research 2.xlsx   # Source data (Pool Data + Pool Hist)
├── HW6_Section3_Analysis.xlsx                 # Section 3 analysis workbook
├── HW6_MBS_Solution.docx                      # Generated Word document output
└── HW  MBS Prepayment Speed, MBS Yield.docx  # Original assignment
```

---

## Sections Covered

| #   | Section                     | Points | Description                                                |
| --- | --------------------------- | ------ | ---------------------------------------------------------- |
| 1   | Prepayment Measurement      | 0.5    | SMM ↔ CPR ↔ PSA conversions, WAL calculation               |
| 2   | Pool Prepayment History     | 0.5    | CUSIP 31418C5Z3 CPR vs. 30-yr PMMS chart                   |
| 3   | Agency MBS Pool Analysis    | 1.5    | Issuance trends, coupon stats, CPR time series, WAC impact |
| 4   | Bloomberg Yield Table Study | 0.5    | Yield/Duration/AL analysis across rate scenarios           |

---

## Key Formulas

```
SMM → CPR:   CPR  = 1 − (1 − SMM)^12
CPR → SMM:   SMM  = 1 − (1 − CPR)^(1/12)
CPR → PSA:   PSA% = CPR / (min(age, 30) × 0.2%) × 100
WAL:         WAL  = Σ(month × principal) / Σ(principal)
```

---

## Requirements

```bash
pip install pandas numpy matplotlib openpyxl python-docx
```

> **Windows only:** `add_charts_win32.py` additionally requires Microsoft Excel installed and `pywin32`:
>
> ```bash
> pip install pywin32
> ```

---

## Usage

### 1. Generate Full Solution (Charts + Word Doc)

```bash
python hw6_solution.py
```

Outputs:

- `charts/` — all PNG charts
- `HW6_MBS_Solution.docx` — formatted Word document

### 2. Populate Excel Analysis File

```bash
python populate_excel.py
```

### 3. Add Charts to Q3 Excel Workbook

```bash
# Using openpyxl (cross-platform)
python add_charts_to_q3.py

# Using win32com (Windows + Excel required)
python add_charts_win32.py
```

### 4. Section 2 – CPR vs PMMS Chart

```bash
python sec2_chart.py
```

### 5. Verify Outputs

```bash
python verify.py
```

---

## Data Sources

- **`HW Q3 Problem_Agency MBS Research 2.xlsx`** — 2,234 Agency MBS pools with issuance and monthly historical data (Jan 2023 – Apr 2025)
- **Bloomberg Terminal screenshots** — Yield table data for Section 4 (CUSIP 01F040610, 4.0% Fannie Mae 30-yr)
- **FRED / Recursion Analyzer** — 30-yr PMMS rates for Section 2 (external access required)

---

## Generated Charts

| File                       | Description                                            |
| -------------------------- | ------------------------------------------------------ |
| `q1_smm_vs_cpr.png`        | SMM vs CPR nonlinear relationship                      |
| `q3_1_issuance.png`        | Monthly Agency MBS issuance by coupon (stacked bar)    |
| `q3_3a_balance.png`        | Monthly total outstanding balance time series          |
| `q3_3b_cpr.png`            | Balance-weighted aggregate CPR time series             |
| `q3_4_wac_cpr.png`         | WAC vs CPR average/median bar chart                    |
| `q3_4_wac_cpr_scatter.png` | WAC vs CPR scatter plot                                |
| `q4_psa_avg_life.png`      | PSA vs Average Life                                    |
| `q4_psa_mod_duration.png`  | PSA vs Modified Duration (3 price scenarios)           |
| `q4_yield_vs_scenario.png` | Yield vs Interest Rate Scenario (par/premium/discount) |
