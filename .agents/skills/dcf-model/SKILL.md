---
name: dcf-model
description: Build institutional Discounted Cash Flow (DCF) valuation models in Excel (.xlsx) with gross-debt WACC, Bear/Base/Bull scenarios, full enterprise-to-equity bridge, and live-formula 5x5 sensitivity analysis. Use for intrinsic equity valuation, WACC builds, and FCF projection models. Not for peer multiples (use comps-analysis) or full 3-statement balance-sheet models (use three-statement-model).
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/dcf-model/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# DCF Model Builder

Build institutional-grade, two-sheet (`DCF` + `WACC`) Excel valuation models programmatically with Python (`openpyxl`), live formulas, full source provenance, and automated LibreOffice + logic verification.

## Quick Reference & Supporting Docs

- **Starter skeleton generator:** `python3 .agents/skills/dcf-model/scripts/build_dcf_skeleton.py --output <path.xlsx> [--case 2]`
- **WACC, FCF, TV & Equity Bridge methodology:** [references/wacc-and-valuation-bridge.md](references/wacc-and-valuation-bridge.md)
- **Exact row map, formulas, sensitivity & styling:** [references/dcf-layout-and-formulas.md](references/dcf-layout-and-formulas.md)
- **Error & valuation debugging:** [TROUBLESHOOTING.md](TROUBLESHOOTING.md)

---

## Non-Negotiable Rules

1. **Live Excel formulas only:** Every projection, subtotal, NOPAT, UFCF, discount factor, PV, terminal value, valuation bridge line, WACC calculation, and sensitivity cell MUST be a live Excel formula (`ws["C74"] = "=B74*(1+C62)"`). Never pre-compute results in Python and write numbers.
2. **Gross-capitalization WACC weights (never negative):** Compute WACC weights on **Market Cap (`E`) + Total Gross Debt (`D`) + Preferred Stock (`P`)**: `Equity Weight = E / (E + D + P)`, `Debt Weight = D / (E + D + P)`, `Preferred Weight = P / (E + D + P)` (all bounded `0%–100%`, sum `= 100%`). Exclude NCI from parent WACC (consolidated UFCFs discounted at consolidated WACC; NCI deducted in the equity bridge). Never use `Net Debt / EV`. Link `DCF!C21 = WACC!C30` as the single source of WACC.
3. **Discounting & terminal growth discipline:**
   - Discount explicit annual UFCFs using the **mid-year convention** (`t - 0.5`: `0.5, 1.5, ..., N - 0.5`).
   - Discount **Perpetuity Growth Terminal Value** consistently at **mid-year `N - 0.5`** (`1 / (1 + WACC)^(N - 0.5)`, dynamically referencing the final cell of the `Period` row `$G$87`), because `UFCF_N * (1 + g) / (WACC - g)` values a mid-year terminal cash-flow stream one period prior to Year `N + 0.5`.
   - Discount **Exit Multiple Terminal Value** strictly at **end of Year `N`**: `1 / (1 + WACC)^N` (`N = 5.0` — **never** `N - 0.5` for an exit-multiple TV).
   - Keep perpetuity terminal growth `g` strictly `< WACC` and `<= long-term nominal GDP / 10Y Treasury` (typically **`2.0%–3.5%`**, e.g., Bear `2.0%`, Base `2.25%–2.5%`, Bull `2.5%–3.0%`; never `3.5%–5.0%`).
4. **Complete equity bridge, diluted shares & SBC treatment:**
   - `Equity Value = Enterprise Value - Total Gross Debt - Preferred Stock - Non-Controlling Interest + Cash & Marketable Securities + Equity Investments`.
   - **ASC 842 warning:** Do NOT deduct operating lease liabilities from EV when UFCF already expenses operating lease rent (double-counting).
   - **Diluted shares:** Separate **Unvested RSUs (`$0` strike — always 100% dilutive)** from **Options Outstanding** (`Basic Shares + Unvested RSUs + IF(Price > Strike, Options - (Options * Strike / Price), 0)`) and apply the **if-converted test** for convertibles (avoid double-counting convertible face value in debt if treated as converted).
   - **SBC:** Treat Stock-Based Compensation as a real economic expense in GAAP EBIT / UFCF (do not add back SBC; if an adjusted view is requested, reconcile explicitly alongside dilution).
5. **Single canonical layout (aligned with `build_dcf_skeleton.py`):**
   - Named `Case` cell at `DCF!B4` (`1=Bear, 2=Base, 3=Bull`) guarded by whole-number `1..3` Excel `DataValidation`, with display label `=CHOOSE(INT(Case),"Bear Case","Base Case","Bull Case")` in `DCF!C4`.
   - Vertically stacked Bear / Base / Bull assumption blocks (uniform 12-row stride, explicit projection year headers) consolidated via `=INDEX(..., (INT($B$4)-1)*12 + offset)` into an Active Case block — zero scattered nested `IF`s, zero hardcoded OpEx percentages in formulas.
   - Live horizon-independent 5×5 `WACC × Terminal Growth` sensitivity table at the bottom of `DCF` (`B124:G129`) using `=IF($B{row}<={col}${hdr},"NM", ...)` around `SUMPRODUCT` over the UFCF and Period ranges, with formula-driven axes centered on live `WACC` (`C21`) and live `Terminal Growth` (`C69`), full DCF recalculation in all 25 cells, and highlighted center cell (`E127`) matching `DCF!C108`.
6. **Cell comments on creation:** Attach `"Source: [System/Document], [Date], [Reference]"` to every blue hardcoded input as it is written; tag analyst assumptions `[ASSUMPTION]` and unverified items `[UNSOURCED]`.
7. **Lock row layout before writing formulas:** Define every section's row constants before writing any formula cell to eliminate off-by-one and `#REF!` bugs.

---

## Interactive vs. Subagent-Safe Checkpoints

- **Interactive mode (direct user session):** Confirm key milestones before building downstream sections:
  1. **Raw inputs & historicals** (revenue, margins, gross debt, cash, diluted shares, fiscal calendar)
  2. **Operating forecast** (Bear/Base/Bull revenue growth, gross margin, OpEx `%` of revenue, D&A, CapEx, NWC)
  3. **WACC build** (Rf, beta, ERP, pre-tax Kd, marginal tax rate, gross-debt weights)
  4. **Terminal value, equity bridge & comps cross-check** (TV `%` of EV, implied exit EV/EBITDA, implied P/E) prior to generating final sensitivities
- **Subagent mode (invoked via `invoke_subagent` or automated workflow):** Do not stall waiting for mid-build user input. Proceed end-to-end using primary filing data, tag every inferred or assumed driver `[ASSUMPTION]`, run both QA scripts, and present the complete assumption summary, comps cross-check, and sensitivity range in the final handoff for human analyst review.

---

## Step-by-Step DCF Workflow

### Step 1: Gather & Verify Primary Data
Follow the `AGENTS.md` source hierarchy:
1. Connected financial MCP servers (if configured in `.agents/mcp_config.json`)
2. SEC EDGAR filings (`10-K`, `10-Q`, `8-K`, `DEF 14A`) and official earnings releases
3. Workspace files under `research/companies/<TICKER>/`
4. Verified web search for live share price, 5Y monthly beta, 10Y US Treasury yield, and credit spread (cite source + date)

Confirm: most recent reported quarter, fiscal year-end month, 3–5 years of historical financials (revenue, gross margin, EBIT margin, D&A, CapEx, NWC), balance sheet cash/gross debt/preferred/NCI/equity investments, and option/RSU/convertible footnotes for diluted shares.

### Step 2: Build the `WACC` Sheet
Follow [references/wacc-and-valuation-bridge.md](references/wacc-and-valuation-bridge.md):
- CAPM Cost of Equity (`Rf + Beta * ERP`)
- After-Tax Cost of Debt (`Pre-Tax Kd * (1 - Marginal Tax Rate)`) and Cost of Preferred Stock (`Kp`, non-tax-deductible; `0.0%` if Preferred `= $0M`)
- Gross-capitalization weights (`E / (E + D + P)`, `D / (E + D + P)`, `P / (E + D + P)`, sum check `= 100.00%`), linking `E = DCF!C14`, `D = DCF!C15`, and `P = DCF!C16` (excluding NCI, which is deducted in the equity bridge)
- Output cell `WACC!C30` highlighted in `#BDD7EE`

### Step 3: Build Scenario Blocks & Active Case Consolidation
Follow [references/dcf-layout-and-formulas.md](references/dcf-layout-and-formulas.md):
- Place `Case` (`1=Bear, 2=Base, 3=Bull`) with whole-number `1..3` `DataValidation` at `DCF!B4` and `=CHOOSE(INT(Case),...)` label at `DCF!C4`.
- Populate stacked Bear / Base / Bull blocks with explicit projection year headers and realistic operating leverage (5Y standard; 7–10Y for high-growth companies).
- Wire the Active Case consolidation block with `INDEX(..., (INT($B$4)-1)*12 + offset)`.

### Step 4: Project Unlevered Free Cash Flow (UFCF)
- Model every OpEx, D&A, and CapEx line as a `% of Revenue` referencing the Active Case block (never `% of Gross Profit`, never hardcoded percentages).
- Compute `EBIT → Cash Taxes → NOPAT + D&A - CapEx - ΔNWC = UFCF` (retaining SBC as an operating expense).
- Apply mid-year discount periods (`0.5, 1.5, ..., N - 0.5`) and discount factors `=1/(1+$C$21)^Period` using `DCF!C21 = WACC!C30`.

### Step 5: Compute Terminal Value, Valuation Bridge & Comps Cross-Check
- **Primary (Perpetuity Growth):** `Undiscounted TV = UFCF_N * (1 + g) / (WACC - g)`, discounted consistently at mid-year `1 / (1 + WACC)^(N - 0.5)` (`=1/(1+$C$21)^$G$87`).
- **Cross-check (Exit Multiple):** `Undiscounted TV = EBITDA_N * Exit Multiple`, strictly discounted at end-of-year `1 / (1 + WACC)^N` (`=C116/(1+$C$21)^($G$87+0.5)`).
- **Bridge to Implied Share Price:** `EV - Gross Debt - Preferred - NCI + Cash & Marketable Securities + Equity Investments = Implied Equity Value ÷ Diluted Shares (Basic + RSUs + ITM Option TSM)`.
- **Cross-check against `comps-analysis`:** Compare implied Year 1 EV/EBITDA, implied exit EV/EBITDA, and implied P/E against peer median and 25th–75th quartiles. Verify **TV `%` of EV** sits inside the institutional **`50%–70%`** target band (`validate_dcf.py` warns if `<50%` or `>70%`).

### Step 6: Build Live-Formula 5×5 Sensitivity Table(s)
- Place at the bottom of `DCF` (starting Row `122` in the canonical layout — do NOT split to a separate sheet unless requested).
- Anchor the 5×5 `WACC × Terminal Growth` axes with live formulas around `C21` (`WACC`) and `C69` (`Terminal Growth`).
- Populate all 25 data cells via an `openpyxl` loop with horizon-independent `SUMPRODUCT` DCF formulas (no linear shortcuts, no Excel What-If Data Tables). Confirm the highlighted center cell (`E127`) equals `C108` exactly.
- Add optional secondary 5×5 tables (e.g., `WACC × Exit Multiple` or `Revenue Growth × EBIT Margin`) below Row `131` when deeper stress-testing is needed.

### Step 7: Format, Recalculate & Validate (MANDATORY QA GATE)
Apply institutional styling from `.agents/rules/financial-modeling.md` (`#1F4E79` dark navy banners, `#D9E1F2` column headers, `#F2F2F2` input fills with blue font, black formulas, green cross-sheet links, `#BDD7EE` key outputs/center cell, thin/total borders, and cell comments on all inputs).

Run BOTH automated QA checks before delivering:

```bash
# 1. Recalculate the newly openpyxl-generated workbook via LibreOffice (--in-place populates cached values on disk)
python3 scripts/recalc_xlsx.py research/companies/<TICKER>/models/<TICKER>_DCF_YYYY-MM-DD.xlsx 30 --in-place

# 2. Validate DCF structure, WACC bounds, g < WACC, TV % of EV, and cached values
python3 .agents/skills/dcf-model/scripts/validate_dcf.py research/companies/<TICKER>/models/<TICKER>_DCF_YYYY-MM-DD.xlsx
```

*(Note: `recalc_xlsx.py` is **non-destructive by default** — running it without `--in-place` or `--output` checks a temporary copy and leaves an `openpyxl`-saved file uncached on disk, which would cause `validate_dcf.py` to fail on uncached formulas. Always pass `--in-place` for newly Python-generated `.xlsx` files or `--output <recalc.xlsx>` when preserving an original workbook, then run `validate_dcf.py` on that recalculated file.)*

#### `scripts/recalc_xlsx.py` Exit Codes, Statuses & Flags

| Exit Code | JSON `status` | Meaning & Required Action |
|---|---|---|
| `0` | `"success"` | Zero formula errors; all formulas recalculated and verified via LibreOffice (or `--allow-unrecalculated` explicitly passed). Safe to proceed. |
| `1` | `"errors_found"` or `"error"` | One or more Excel errors (`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`), corrupt/missing file, or LibreOffice timeout/failure. Inspect `error_summary`, consult [TROUBLESHOOTING.md](TROUBLESHOOTING.md), fix, and re-run. |
| `2` | `"unverified"` (or CLI usage error) | LibreOffice (`soffice`) is not installed and the workbook has uncached formula values. Install LibreOffice to recalculate, or pass `--allow-unrecalculated` only if the user explicitly accepts unverified cached values. |

`recalc_xlsx.py` mode flags:
- *(default)* / `--check-only` — Non-destructive read-only check: recalculates in a temp directory and audits without mutating the file on disk (ideal for read-only audits of untouched workbooks).
- `--in-place` — Atomically replaces a newly generated `.xlsx` file with its LibreOffice-recalculated copy so cached formula values are populated on disk for `validate_dcf.py` and `openpyxl(data_only=True)` reads (refuses non-`.xlsx` files and What-If data tables).
- `--output <PATH>` — Writes the recalculated `.xlsx` workbook to a new path without modifying the original input.
- `--allow-unrecalculated` — Opt-in override returning exit `0` when LibreOffice is unavailable (always disclose `recalculated_with_libreoffice: false` to the analyst).

---

## Deliverable Path & Final Checklist

- **Default output path:** `research/companies/<TICKER>/models/<TICKER>_DCF_YYYY-MM-DD.xlsx`
- **Starter skeleton test:** You can bootstrap or compare against `python3 .agents/skills/dcf-model/scripts/build_dcf_skeleton.py --output <path.xlsx> --case 2`.

Before handing off the model, confirm:
- [ ] `scripts/recalc_xlsx.py <file.xlsx> 30 --in-place` (or `--output <recalc.xlsx>`) exits `0` (`"status": "success"`, `total_errors: 0`, `uncached_formulas: 0`)
- [ ] `validate_dcf.py` exits `0` (`"status": "PASS"`)
- [ ] Two sheets (`DCF` + `WACC`), named `Case` at `DCF!B4` with `1..3` `DataValidation`, working Bear/Base/Bull toggle
- [ ] Gross-capitalization WACC weights (`E + D + P`, all `0%–100%`, sum `= 100%`), `g < WACC`, `g <= 3.5%`
- [ ] Mid-year explicit FCF (`t - 0.5`) and mid-year perpetuity TV (`N - 0.5`) discounting; strict end-of-year `(1 + WACC)^N` exit-multiple TV discounting
- [ ] Complete equity bridge (Gross Debt, Preferred, NCI, Cash, Equity Investments), separate RSU + option TSM diluted shares, SBC kept as expense
- [ ] Live `SUMPRODUCT` 5×5 sensitivity table with center cell (`E127`) matching `DCF!C108`
- [ ] Blue inputs (all commented with sources), black formulas, green links, institutional borders and number formats
- [ ] Executive summary provided alongside the workbook (Base/Bear/Bull implied prices, key drivers, comps cross-check, and `[ASSUMPTION]` disclosures — labeled `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY`)