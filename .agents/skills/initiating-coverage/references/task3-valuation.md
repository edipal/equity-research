<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/initiating-coverage/references/task3-valuation.md), Apache-2.0; modified for Google Antigravity. -->

# Task 3: Valuation Analysis - Detailed Workflow

Step-by-step instructions for executing Task 3 (Valuation Analysis) of the `initiating-coverage` skill. All outputs are staged as `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`.

## Task Overview

**Purpose:** Perform multi-methodology valuation (DCF + Trading Comparables, plus Precedent Transactions as cross-check when M&A is relevant), add canonical valuation tabs to Task 2's financial model, and derive a 12-month Price Target and standardized rating (`Buy / Hold / Sell`).

**Prerequisites (MANDATORY — stop if Task 2 is incomplete):**
- Verified Task 2 model at `research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx` containing `Revenue Model`, `Income Statement`, `Cash Flow Statement`, `Balance Sheet`, `Scenarios` and passing `scripts/recalc_xlsx.py`.
- If Task 2 does not exist, STOP and inform the user to complete Task 2 first.

**Core Methodology Delegation:**
- **DCF Mechanics:** Delegate ALL core DCF methodology, formulas, and layout conventions to the **`dcf-model`** skill (`.agents/skills/dcf-model/SKILL.md`, `references/wacc-and-valuation-bridge.md`, `references/dcf-layout-and-formulas.md`, and `scripts/build_dcf_skeleton.py`). Do NOT deviate from:
  - **Gross-debt WACC weights:** `Total Capital = Market Cap (E) + Total Gross Debt (D)`; `Equity Weight = E/(E+D)`, `Debt Weight = D/(E+D)` (bounded `0%–100%`, never negative).
  - **Mid-year explicit UFCF discounting** (`t - 0.5`: `0.5, 1.5, ..., N - 0.5`); **end-of-year `(1 + WACC)^N` discounting for Terminal Value** (both perpetuity growth AND exit multiple — NEVER `N - 0.5` for exit multiple).
  - **Terminal growth `g`:** Strictly `< WACC` and `<= long-term nominal GDP / Rf` (typically `2.0%–3.5%`).
  - **TV `%` of EV:** Target `50%–70%`; flag and investigate if `< 40%` or `> 80%`.
  - **Complete Equity Bridge & Diluted Shares:** `EV - Total Gross Debt - Preferred Stock - NCI + Cash & Marketable Securities + Equity Investments ÷ Fully Diluted Shares (TSM + if-converted)`. Never deduct ASC 842 operating lease liabilities when UFCF already expenses lease rent. Keep SBC as a real economic expense in UFCF.
  - **Centered Odd Sensitivity Grid (5×5):** Formula-driven axes centered on live `WACC` and `Terminal Growth`, full DCF recalculation in all 25 cells, highlighted center cell matching Base DCF price.
- **Comps Mechanics:** Follow **`comps-analysis`** for EV construction, NTM calendarization, `"NM"` error handling (`IFERROR(..., "NM")`), and mandatory Max / 75th / Median / 25th / Min statistics.
- **Valuation Theory Reference:** See [valuation-methodologies.md](valuation-methodologies.md).

---

## Standardized Rating Scale (Use Consistently Across Tasks 3 & 5)

| Rating | 12-Month Expected Total Return Threshold | Interpretation |
|---|---|---|
| **Buy** | `>= +15.0%` | Risk-adjusted upside exceeds institutional hurdle rate |
| **Hold** | `-10.0%` to `+15.0%` | Valuation broadly reflects fundamentals; balanced risk/reward |
| **Sell** | `<= -10.0%` | Material downside risk to intrinsic / peer-implied value |

*(Use `Buy / Hold / Sell` consistently across Task 3, Task 5, and Page 1 — do not mix with `Outperform / Neutral / Underweight` in the same report.)*

---

## Step-by-Step Valuation Workflow

### Step 1: Link Model Projections from Task 2

Pull live cross-sheet links from Task 2's workbook (`Income Statement`, `Cash Flow Statement`, `Balance Sheet`, `Scenarios`):
- 5-year projected Revenue, EBIT, Tax Rate, D&A, EBITDA, CapEx, and ΔNWC (`[FY+1]E–[FY+5]E`)
- Latest balance sheet Gross Debt, Preferred Stock, Non-Controlling Interest, Cash & Marketable Securities, Equity Investments
- Diluted share count inputs (basic shares + ITM options/RSUs for TSM + convertibles if-converted test)
- Bear / Base / Bull scenario trajectories

### Step 2: Build `WACC` and `DCF` Tabs (Per `dcf-model` Spec)

**Worked Illustrative Example (Every figure mathematically reconciled):**
Assume illustrative current share price `$42.00`, fully diluted shares `100.0M`, Total Gross Debt `$500.0M`, Cash & Marketable Securities `$50.0M` (`Net Debt = $450.0M`), Preferred `$0`, NCI `$0`, Equity Investments `$0`:

1. **WACC (Gross-Debt Basis on `WACC` tab):**
   - `Market Cap (E) = $42.00 × 100.0M = $4,200.0M ($4.2B)`
   - `Total Gross Debt (D) = $500.0M` → `Total Capitalization (E + D) = $4,700.0M`
   - `Equity Weight = $4,200.0 / $4,700.0 = 89.36%` | `Debt Weight = $500.0 / $4,700.0 = 10.64%` (sum `= 100.00%`)
   - `Cost of Equity (Ke) = 4.20% Rf + 1.30 Beta × 5.50% ERP = 11.35%`
   - `After-Tax Cost of Debt = 6.50% Pre-Tax Kd × (1 - 25.0% Tax Rate) = 4.875%`
   - `WACC = (89.36% × 11.35%) + (10.64% × 4.875%) = 10.14% + 0.52% = 10.66%`

2. **5-Year Mid-Year UFCF & End-of-Year Terminal Value (`DCF` tab):**
   - Illustrative 5-year UFCFs (`[FY+1]E–[FY+5]E`): `$250M`, `$320M`, `$390M`, `$450M`, `$500M`
   - Mid-year discount periods (`0.5, 1.5, 2.5, 3.5, 4.5`) at `WACC = 10.66%`:
     - PV of Explicit UFCFs: `$237.7M + $274.9M + $302.8M + $315.7M + $317.0M = $1,447.9M`
   - Perpetuity Growth Terminal Value (`g = 2.50%`):
     - `Undiscounted TV = $500.0M × (1.025) / (0.1066 - 0.025) = $6,280.6M`
     - `PV of TV (end of Year 5) = $6,280.6M / (1.1066)^5 = $3,784.9M`
   - `Enterprise Value = $1,447.9M + $3,784.9M = $5,232.8M` (`TV % of EV = 72.3%` — within `40%–80%` guardrail; close to `50%–70%` target, note in narrative)
   - Complete Equity Bridge:
     - `EV $5,232.8M - Gross Debt $500.0M + Cash $50.0M = Implied Equity Value $4,782.8M`
     - `Implied Share Price (Base DCF) = $4,782.8M / 100.0M = $47.83` (`+13.9%` vs. `$42.00`)
   - Cross-check Exit Multiple TV (`EBITDA_5 × Exit Multiple`, strictly discounted at `(1 + WACC)^5`, NEVER `4.5`).

3. **Centered 5×5 Live-Formula Sensitivity Table (`DCF` inline or optional `Sensitivity` tab):**
   Axes centered on live `WACC = 10.66%` (`±50bps / ±100bps`) and live `g = 2.50%` (`±50bps / ±100bps`):

```text
Implied Share Price ($)          Terminal Growth Rate (g)
WACC        1.50%     2.00%     2.50%     3.00%     3.50%
 9.66%     $49.55    $52.32    $55.47    $59.09    $63.31
10.16%     $46.28    $48.68    $51.40    $54.49    $58.05
10.66%     $43.37    $45.47   [$47.83]   $50.50    $53.54   <- Base WACC row
11.16%     $40.76    $42.61    $44.68    $46.99    $49.61
11.66%     $38.42    $40.06    $41.87    $43.90    $46.18
                                  ^
                             Base g col (center cell = $47.83)
```

   DCF Low / Base / High range (using `±50bps` WACC & `±50bps` g around base): **`$42.61 – $47.83 – $54.49`**.

### Step 3: Build `Comparable Companies` Tab (Per `comps-analysis` Spec)

Select **5–10 public peers** (where a valid peer group exists; document selection rationale and any peer exclusions). Compute full EV (`Market Cap + Gross Debt + Preferred + NCI - Cash & Short-Term Investments`) and calendarized NTM/LTM multiples (`EV/Revenue`, `EV/EBITDA`, `P/E`). Include the **mandatory statistical summary** (`Max`, `75th Percentile`, `Median`, `25th Percentile`, `Min`) using `IFERROR(..., "NM")`.

**Worked Illustrative Example (`Target NTM EBITDA = $550.0M`, `Net Debt = $450.0M`, `100.0M` diluted shares):**
- **25th Percentile (`12.5x` NTM EV/EBITDA):** `EV = $6,875.0M` → `Equity = $6,425.0M` → **`$64.25/sh`**
- **Median (`13.8x` NTM EV/EBITDA):** `EV = $7,590.0M` → `Equity = $7,140.0M` → **`$71.40/sh`**
- **75th Percentile (`14.1x` NTM EV/EBITDA):** `EV = $7,755.0M` → `Equity = $7,305.0M` → **`$73.05/sh`**

Justify any applied premium/discount vs. peer median based on relative revenue growth, EBITDA margin, ROIC, and moat durability.

### Step 4: Precedent Transactions (Optional Cross-Check Only)

Include precedent M&A deals (target ~5–10 deals **where relevant recent transactions exist** — never invent deals) ONLY when sector M&A provides meaningful valuation context.

**Worked Illustrative Example (`Target LTM EBITDA = $500.0M`, `Median Precedent = 15.8x LTM EV/EBITDA`):**
- `Implied EV = $500.0M × 15.8x = $7,900.0M`
- `Less: Net Debt ($450.0M)` → `Implied Equity Value = $7,450.0M`
- `Implied Share Price = $7,450.0M / 100.0M = $74.50/sh`

**Control-Premium Weighting Guardrail:** Precedent transaction multiples embed a `20%–40%` control premium and buyer synergies. When near-term M&A is **unlikely**, assign **`0%` weight** to Precedent Transactions in the standalone 12-month Price Target and display it strictly as a **cross-check / takeover floor** on the football field. Only assign positive weight (`10%–25%`) when an active strategic review or high-probability M&A catalyst is explicitly part of the investment thesis.

### Step 5: Build `Valuation Summary` Tab & Reconcile Price Target

Combine standalone methodologies cleanly and cross-check against Bear/Base/Bull scenarios:

```text
VALUATION SUMMARY (ILLUSTRATIVE RECONCILIATION)

Method                       Low      Base     High     Weight   Weighted Base
DCF Analysis (Perpetuity)    $42.61   $47.83   $54.49    50%     $23.92
Trading Comps (NTM EBITDA)   $64.25   $71.40   $73.05    50%     $35.70
Precedent Trans. (LTM)*      $67.50   $74.50   $81.50     0%      $0.00  (cross-check only; M&A unlikely)
------------------------------------------------------------------------------
Blended 12M Price Target     $53.43   $59.62   $63.77   100%     $59.62 -> Rounded: $60.00

Current Share Price ([Date]): $42.00
Implied 12M Upside:           +42.9% ($60.00 / $42.00 - 1) -> Rating: BUY (>= +15% threshold)
```

**Scenario Probability-Weighted Cross-Check (verify arithmetic exactly):**
```text
Scenario    Probability   Rev CAGR   Term. EBITDA Margin   Implied Value   Weighted Contrib.
Bear Case       20%         18%             28%               $42.00            $8.40
Base Case       60%         25%             32%               $59.00           $35.40
Bull Case       20%         32%             36%               $82.00           $16.40
--------------------------------------------------------------------------------------------
Probability-Weighted Expected Value (100%):                                    $60.20
```
*(Always verify `SUMPRODUCT(Probability, Implied Value)` exactly — e.g., `0.20 × $42.00 + 0.60 × $59.00 + 0.20 × $82.00 = $60.20`, aligning closely with the `$60.00` blended target!)*

### Step 6: Document Catalysts, Risks & Draft Notice

Include 3–5 dated catalysts and 3–5 quantified upside/downside risks in `<TICKER>_Valuation_Analysis_<Date>.md`, labeled **`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`**, with `[ASSUMPTION]` tags on all analyst projections.

### Step 7: Mandatory Automated QA Gates

After adding the 4 canonical valuation tabs (`WACC`, `DCF`, `Comparable Companies`, `Valuation Summary` — plus inline 5×5 sensitivity on `DCF` or optional `Sensitivity` tab) to `<TICKER>_Financial_Model_<Date>.xlsx`, run BOTH QA scripts:

```bash
# 1. Recalculate via LibreOffice and verify zero formula errors / zero uncached formulas
python3 scripts/recalc_xlsx.py research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx 30

# 2. Validate DCF sheet structure, WACC range, g < WACC, TV % of EV, and cached values
python3 .agents/skills/dcf-model/scripts/validate_dcf.py research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx
```

Both commands MUST exit `0` (`"status": "success"` and `"status": "PASS"`) before handing off Task 3.

---

## Output Files & Success Criteria

1. `research/companies/<TICKER>/reports/<TICKER>_Valuation_Analysis_YYYY-MM-DD.md` (target ~4–6 pages, draft notice included)
2. Canonical tabs added to Task 2's existing workbook (`research/companies/<TICKER>/models/<TICKER>_Financial_Model_YYYY-MM-DD.xlsx`):
   - `WACC`
   - `DCF` *(with inline 5×5 live sensitivity table per `dcf-model`, or optional dedicated `Sensitivity` tab)*
   - `Comparable Companies`
   - `Valuation Summary` *(+ optional `Precedent Transactions` cross-check tab)*

**Checklist before proceeding to Task 4:**
- [ ] `scripts/recalc_xlsx.py` exits `0` (`"status": "success"`) and `validate_dcf.py` exits `0` (`"status": "PASS"`)
- [ ] Gross-debt WACC, mid-year explicit FCF, end-of-year `(1+WACC)^N` TV discounting, `g <= 3.5%` and `< WACC`, centered 5×5 live sensitivity grid
- [ ] Comps table includes Max / 75th / Median / 25th / Min statistical summary
- [ ] Precedent M&A assigned `0%` weight (cross-check only) when M&A is unlikely
- [ ] Standardized `Buy / Hold / Sell` rating consistent with upside/downside threshold
- [ ] All arithmetic cross-checked and all forward drivers tagged `[ASSUMPTION]`
