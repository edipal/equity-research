<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/initiating-coverage/references/task3-valuation.md), Apache-2.0; modified for Google Antigravity. -->

# Task 3: Valuation Analysis - Detailed Workflow

Step-by-step instructions for executing Task 3 (Valuation Analysis) of the `initiating-coverage` skill. All outputs are staged as `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`.

## Task Overview

**Purpose:** Perform multi-methodology valuation (DCF + Trading Comparables, plus Precedent Transactions as cross-check when M&A is relevant), add canonical valuation tabs to Task 2's financial model, and derive a 12-month Price Target and standardized rating (`Buy / Hold / Sell`).

**Prerequisites (MANDATORY — stop if Task 2 is incomplete):**
- Verified Task 2 model at `research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx` containing `Revenue Model`, `Income Statement`, `Cash Flow Statement`, `Balance Sheet`, `Scenarios` and passing `scripts/recalc_xlsx.py`.
- If Task 2 does not exist, STOP and inform the user to complete Task 2 first.

**Core Methodology Delegation (Integrated 9-Tab Initiation Workbook):**
- **DCF Mechanics:** Follow the **`dcf-model`** skill (`.agents/skills/dcf-model/SKILL.md`, `references/wacc-and-valuation-bridge.md`, `references/dcf-layout-and-formulas.md`) for **all valuation methodology, discounting, equity bridge, sensitivity, and QA conventions**, while adapting the `DCF` tab's input block to link live to Task 2's 3-statement tabs rather than duplicating `build_dcf_skeleton.py`'s standalone hardcoded driver blocks (which are designed only for standalone 2-tab DCF workbooks built without a 3-statement model):
  - **Single Scenario Selector (NEVER duplicate):** Maintain **ONE** workbook-wide scenario selector — named range `Case` (`1=Bear, 2=Base, 3=Bull`, guarded by `1..3` Excel `DataValidation`) on Task 2's `Scenarios` tab, referenced by `DCF` via live green cross-sheet link (`=Scenarios!<case_cell>`). Never create a second independent scenario switch on `DCF`.
  - **Live Cross-Sheet Links to Task 2 (NEVER re-key Task 2 drivers):** Pull 5-year explicit UFCF lines (Revenue, EBIT, Tax Rate / Cash Taxes / NOPAT, D&A, EBITDA, CapEx, ΔNWC) and balance-sheet bridge items via live green cross-sheet formulas directly from `Income Statement`, `Cash Flow Statement`, and `Balance Sheet`. Drive terminal parameters (Bear/Base/Bull terminal growth `g` and exit multiple) from the single `Case` selector and bind the canonical workbook named ranges (`Case` on `Scenarios`, `DCF_WACC` on `WACC`, `TerminalGrowth`, `TerminalGrowth_Bear/Base/Bull` where used, `PV_TerminalValue`, `EnterpriseValue`) so `validate_dcf.py` validates the integrated workbook cleanly.
  - **Gross-capitalization WACC weights:** `Total Capital = Market Cap (E) + Total Gross Debt (D) + Preferred Stock (P)`; `Equity Weight = E/(E+D+P)`, `Debt Weight = D/(E+D+P)`, `Preferred Weight = P/(E+D+P)` (all bounded `0%–100%`, sum `= 100%`; NCI excluded from parent WACC and deducted in the equity bridge).
  - **Mid-year explicit UFCF & Perpetuity TV discounting:** Discount explicit annual UFCFs at `t - 0.5` (`0.5, 1.5, ..., N - 0.5`) and **Perpetuity Growth TV** consistently at **`(1 + WACC)^(N - 0.5)`** (`4.5` for a 5-year model). Discount **Exit Multiple TV** strictly at **end-of-year `(1 + WACC)^N`** (`5.0` — NEVER `N - 0.5` for exit multiple).
  - **Terminal growth `g`:** Strictly `< WACC` and `<= long-term nominal GDP / Rf` (typically `2.0%–3.5%`).
  - **TV `%` of EV (Mature-Company Diagnostic / Disclosure Check — NOT a Valuation Target):** `50%–70%` is a mature-company diagnostic benchmark (`validate_dcf.py` warns if `< 50%` or `> 70%`). **NEVER adjust WACC, terminal growth `g`, or operating assumptions just to force TV/EV inside `50%–70%`.** When TV/EV falls outside `50%–70%` (e.g., `70%–85%+` for high-growth/reinvestment-heavy companies, or `< 50%` for mature/declining cash cows), retain fundamental assumptions, disclose the TV concentration and economic driver explicitly, and cross-check via Exit Multiple TV and trading comps. *(Selecting a 7–10 year explicit forecast horizon for an early-stage/high-growth company is an upfront modeling choice based on time-to-steady-state, NOT a post-hoc fix to silence a warning.)*
  - **Complete Equity Bridge & Diluted Shares:** `EV - Total Gross Debt - Preferred Stock - NCI + Cash & Marketable Securities + Equity Investments ÷ Fully Diluted Shares (Basic + Unvested RSUs + ITM Option TSM + if-converted)`. Never deduct ASC 842 operating lease liabilities when UFCF already expenses lease rent. Keep SBC as a real economic expense in UFCF.
  - **Centered Odd Sensitivity Grid (5×5):** Formula-driven axes centered on live `WACC` and `Terminal Growth`, horizon-independent `=IF(WACC<=g,"NM", ... SUMPRODUCT(...) ...)` full DCF recalculation in all 25 cells, highlighted center cell matching Base DCF price.
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

### Step 1: Link Model Projections from Task 2 (No Duplicated Driver Blocks)

Pull live green cross-sheet links directly from Task 2's workbook (`Income Statement`, `Cash Flow Statement`, `Balance Sheet`, `Scenarios`) into `DCF` and `WACC` — **never re-key Task 2's Bear/Base/Bull operating driver blocks** from `build_dcf_skeleton.py`:
- **Single scenario selector:** Reference `Scenarios`'s named `Case` cell (`1=Bear, 2=Base, 3=Bull`) from `DCF` so toggling `Case` once updates Task 2 statements and Task 3 valuation simultaneously.
- **Explicit UFCF inputs (`[FY+1]E–[FY+5]E`):** Link Revenue, EBIT, Effective/Marginal Tax Rate (for Cash Taxes / NOPAT), D&A, EBITDA, CapEx, and ΔNWC directly to `Income Statement`, `Cash Flow Statement`, and `Balance Sheet`.
- **Balance-sheet bridge items:** Link latest Gross Debt, Preferred Stock, Non-Controlling Interest, Cash & Marketable Securities, and Equity Investments directly to `Balance Sheet`.
- **Diluted share count inputs:** Basic shares + unvested RSUs (`$0` strike) + ITM options (TSM) + convertibles if-converted test.
- **Named ranges for `validate_dcf.py`:** Define `Case` (on `Scenarios`), `DCF_WACC` (on `WACC`), `TerminalGrowth` (and `TerminalGrowth_Bear`/`Base`/`Bull` where used), `PV_TerminalValue`, and `EnterpriseValue`.

### Step 2: Build `WACC` and `DCF` Tabs (Per `dcf-model` Methodology)

**Worked Illustrative Example (Every figure mathematically reconciled):**
Assume illustrative current share price `$42.00`, fully diluted shares `100.0M`, Total Gross Debt `$500.0M`, Cash & Marketable Securities `$50.0M` (`Net Debt = $450.0M`), Preferred `$0`, NCI `$0`, Equity Investments `$0`:

1. **WACC (3-Tranche Gross Capitalization `E + D + P` on `WACC` tab):**
   - `Market Cap (E) = $42.00 × 100.0M = $4,200.0M ($4.2B)`
   - `Total Gross Debt (D) = $500.0M` | `Preferred Stock (P) = $0.0M` → `Total Capitalization (E + D + P) = $4,700.0M`
   - `Equity Weight = $4,200.0 / $4,700.0 = 89.36%` | `Debt Weight = $500.0 / $4,700.0 = 10.64%` | `Preferred Weight = 0.00%` (sum `= 100.00%`)
   - `Cost of Equity (Ke) = 4.20% Rf + 1.30 Beta × 5.50% ERP = 11.35%`
   - `After-Tax Cost of Debt = 6.50% Pre-Tax Kd × (1 - 25.0% Tax Rate) = 4.875%`
   - `WACC = (89.36% × 11.35%) + (10.64% × 4.875%) + (0.00% × Kp) = 10.14% + 0.52% + 0.00% = 10.66%`

2. **5-Year Mid-Year UFCF & Mid-Year Perpetuity Terminal Value (`DCF` tab):**
   - Illustrative 5-year UFCFs (`[FY+1]E–[FY+5]E`): `$250M`, `$320M`, `$390M`, `$450M`, `$500M`
   - Mid-year discount periods (`0.5, 1.5, 2.5, 3.5, 4.5`) at `WACC = 10.66%`:
     - PV of Explicit UFCFs: `$237.7M + $274.9M + $302.8M + $315.7M + $317.0M = $1,447.9M` *(unrounded sum `$1,447.94M`)*
   - Perpetuity Growth Terminal Value (`g = 2.50%`, mid-year `N - 0.5 = 4.5` discounting):
     - `Undiscounted TV = $500.0M × (1.025) / (0.1066 - 0.025) = $6,280.6M`
     - `PV of TV (mid-year perpetuity, t = 4.5) = $6,280.6M / (1.1066)^4.5 = $3,981.5M`
   - `Enterprise Value = $1,447.9M + $3,981.5M = $5,429.4M` (`TV % of EV = 73.3%` — above the `50%–70%` mature-company diagnostic benchmark; do **not** tune WACC or `g` to clear the warning — disclose the growth-weighted TV concentration in narrative and cross-check via Exit Multiple TV and trading comps)
   - Complete Equity Bridge:
     - `EV $5,429.4M - Gross Debt $500.0M + Cash $50.0M = Implied Equity Value $4,979.4M`
     - `Implied Share Price (Base DCF) = $4,979.4M / 100.0M = $49.79` (`+18.6%` vs. `$42.00`)
   - Cross-check Exit Multiple TV (`EBITDA_5 × Exit Multiple`, strictly discounted at end-of-year `(1 + WACC)^5`, NEVER `4.5`).

3. **Centered 5×5 Live-Formula Sensitivity Table (`DCF` inline or optional `Sensitivity` tab):**
   Axes centered on live `WACC = 10.66%` (`±50bps / ±100bps`) and live `g = 2.50%` (`±50bps / ±100bps`):

```text
Implied Share Price ($)          Terminal Growth Rate (g)
WACC        1.50%     2.00%     2.50%     3.00%     3.50%
 9.66%     $51.40    $54.30    $57.60    $61.40    $65.81
10.16%     $48.07    $50.59    $53.44    $56.69    $60.43
10.66%     $45.10    $47.31   [$49.79]   $52.60    $55.80   <- Base WACC row
11.16%     $42.44    $44.40    $46.57    $49.01    $51.78
11.66%     $40.05    $41.78    $43.70    $45.84    $48.25
                                  ^
                             Base g col (center cell = $49.79)
```

   DCF Low / Base / High range (using `±50bps` WACC & `±50bps` g around base): **`$44.40 – $49.79 – $56.69`**.

### Step 3: Build `Comparable Companies` Tab (Per `comps-analysis` Spec)

Select **5–10 public peers** (where a valid peer group exists; document selection rationale and any peer exclusions). Compute full EV (`Market Cap + Gross Debt + Preferred + NCI - Cash & Short-Term Investments - Equity Investments`) and calendarized NTM/LTM multiples (`EV/Revenue`, `EV/EBITDA`, `P/E`). Include the **mandatory statistical summary** (`Max`, `75th Percentile`, `Median`, `25th Percentile`, `Min`) using `IFERROR(..., "NM")`.

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
DCF Analysis (Perpetuity)    $44.40   $49.79   $56.69    50%     $24.90
Trading Comps (NTM EBITDA)   $64.25   $71.40   $73.05    50%     $35.70
Precedent Trans. (LTM)*      $67.50   $74.50   $81.50     0%      $0.00  (cross-check only; M&A unlikely)
------------------------------------------------------------------------------
Blended 12M Price Target     $54.33   $60.60   $64.87   100%     $60.60 -> Rounded: $61.00

Current Share Price ([Date]): $42.00
Implied 12M Upside:           +45.2% ($61.00 / $42.00 - 1) -> Rating: BUY (>= +15% threshold)
```

**Scenario Probability-Weighted Cross-Check (verify arithmetic exactly):**
```text
Scenario    Probability   Rev CAGR   Term. EBITDA Margin   Implied Value   Weighted Contrib.
Bear Case       20%         18%             28%               $44.00            $8.80
Base Case       60%         25%             32%               $60.50           $36.30
Bull Case       20%         32%             36%               $80.00           $16.00
--------------------------------------------------------------------------------------------
Probability-Weighted Expected Value (100%):                                    $61.10
```
*(Always verify `SUMPRODUCT(Probability, Implied Value)` exactly — e.g., `0.20 × $44.00 + 0.60 × $60.50 + 0.20 × $80.00 = $61.10`, aligning closely with the `$61.00` blended target!)*

### Step 6: Document Catalysts, Risks & Draft Notice

Include 3–5 dated catalysts and 3–5 quantified upside/downside risks in `<TICKER>_Valuation_Analysis_<Date>.md`, labeled **`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`**, with `[ASSUMPTION]` tags on all analyst projections.

### Step 7: Mandatory Automated QA Gates

After adding the 4 canonical valuation tabs (`WACC`, `DCF`, `Comparable Companies`, `Valuation Summary` — plus inline 5×5 sensitivity on `DCF` or optional `Sensitivity` tab) to `<TICKER>_Financial_Model_<Date>.xlsx`, run BOTH QA scripts:

```bash
# 1. Recalculate via LibreOffice so cached values are populated directly in the canonical file
#    (use --in-place on clean openpyxl output, or stage to /tmp/<name>_staged.xlsx, write --output directly to the canonical path, and delete the temp file)
python3 scripts/recalc_xlsx.py research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx 30 --in-place

# 2. Validate DCF sheet structure, WACC range, g < WACC, TV % of EV diagnostic, and cached values
python3 .agents/skills/dcf-model/scripts/validate_dcf.py research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx
```

Both commands MUST exit `0` (`"status": "success"` and `"status": "PASS"`) before handing off Task 3.

---

## Output Files & Success Criteria

1. `research/companies/<TICKER>/reports/<TICKER>_Valuation_Analysis_YYYY-MM-DD.md` (target ~4–6 pages, draft notice included)
2. Canonical tabs added to Task 2's existing workbook (`research/companies/<TICKER>/models/<TICKER>_Financial_Model_YYYY-MM-DD.xlsx`):
   - `WACC`
   - `DCF` *(linked live to Task 2 statements and `Scenarios`'s single `Case` selector, with inline 5×5 live sensitivity table per `dcf-model` or optional dedicated `Sensitivity` tab)*
   - `Comparable Companies`
   - `Valuation Summary` *(+ optional `Precedent Transactions` cross-check tab)*

**Checklist before proceeding to Task 4:**
- [ ] `scripts/recalc_xlsx.py` populates cached formula values directly into the canonical `<TICKER>_Financial_Model_<Date>.xlsx` (`"status": "success"`, `uncached_formulas: 0`) and `validate_dcf.py` exits `0` (`"status": "PASS"`)
- [ ] Single `Case` selector on `Scenarios` linked by `DCF`; explicit UFCF and bridge items linked live to `Income Statement`, `Cash Flow Statement`, and `Balance Sheet` (zero duplicated hardcoded operating driver blocks)
- [ ] Gross-capitalization (`E + D + P`) WACC, mid-year explicit FCF (`t - 0.5`) and mid-year perpetuity TV (`N - 0.5`) discounting, strict end-of-year `(1+WACC)^N` exit-multiple TV discounting, `g <= 3.5%` and `< WACC`, TV `%` of EV disclosed (never assumption-tuned), centered `IF(WACC<=g,"NM", SUMPRODUCT(...))` 5×5 live sensitivity grid
- [ ] Comps table includes Max / 75th / Median / 25th / Min statistical summary
- [ ] Precedent M&A assigned `0%` weight (cross-check only) when M&A is unlikely
- [ ] Standardized `Buy / Hold / Sell` rating consistent with upside/downside threshold
- [ ] All arithmetic cross-checked and all forward drivers tagged `[ASSUMPTION]`
