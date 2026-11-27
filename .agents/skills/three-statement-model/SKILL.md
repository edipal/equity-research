---
name: three-statement-model
description: Build from scratch or populate integrated 3-statement financial models (Income Statement, Balance Sheet, Cash Flow Statement, working capital, PP&E/D&A, debt/revolver, and tax/NOL schedules) in Excel (.xlsx) using Python/openpyxl. Use when constructing or completing linked historical and projected IS/BS/CF workbooks. Not for standalone intrinsic valuation (use dcf-model) or peer multiples (use comps-analysis).
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/3-statement-model/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# 3-Statement Financial Model Builder & Template Completion

Build integrated 3-statement financial models **from scratch** OR **populate existing Excel templates** with full cross-statement linkages (Income Statement, Balance Sheet, Cash Flow Statement, and supporting schedules).

## ⚠️ CRITICAL PRINCIPLES — Read Before Building or Populating

**Execution Environment (Antigravity / Python `openpyxl`):**
- Build or update `.xlsx` workbooks programmatically with Python (`openpyxl`). Write live Excel formula strings (`ws["D15"] = "=D14*(1+Assumptions!$B$5)"`), then run `python3 scripts/recalc_xlsx.py <path-to-model.xlsx> 30 --in-place` (for newly generated `.xlsx` files) or `--output <recalc.xlsx>` (when preserving an existing template/file; exit `0`, `"status": "success"`) so cached formula values are populated on disk, and run the `audit-xls` checklist before delivery.
- **Template Preservation Warning (`openpyxl`):** Loading and re-saving a complex user template via `openpyxl` can **strip existing Excel charts, embedded images, ActiveX controls, and VBA macros** (`.xlsm`). Always back up the original template first, inspect `wb.defined_names` and sheet structure programmatically, and prefer building a clean workbook from scratch when no strict template layout is required.
- **Default output path** (aligned with `AGENTS.md` / `.agents/rules/research-outputs.md`): save to `research/companies/<TICKER>/models/<TICKER>_3Statement_YYYY-MM-DD.xlsx`.

**Formulas over hardcodes (non-negotiable):**
- Every projection cell, roll-forward, linkage, subtotal, and check MUST be a live Excel formula — never a pre-computed Python value.
- The ONLY hardcoded numbers permitted are: (1) historical actuals (with filing citations in cell comments) and (2) assumption drivers (tagged `[ASSUMPTION]` in cell comments).

**Verify step-by-step (Interactive vs. Subagent Mode):**
- **Interactive mode:** Confirm at major milestones — (1) architecture/template mapping, (2) historical actuals, (3) IS projections, (4) BS balance check `= 0` every period, (5) CF cash tie-out — before finalizing.
- **Subagent mode (e.g., invoked by `model-builder`):** Do not stall waiting for mid-build confirmation. Build the complete integrated 3-statement model end-to-end, tag all assumed drivers `[ASSUMPTION]`, run `python3 scripts/recalc_xlsx.py <file.xlsx> 30 --in-place` (or `--output <recalc.xlsx>`), verify BS balance `= 0` and CF cash tie-out `= 0` across all scenarios, and present the full audit summary and assumption table in the handoff.

## Formatting — Professional Blue/Grey Palette (Default unless template/user specifies otherwise)

**Keep colors minimal.** Use only blues and greys for cell fills. Do NOT introduce greens, yellows, oranges, or multiple accent colors — a clean model uses restraint.

| Element | Fill | Font |
|---|---|---|
| Section headers (IS / BS / CF titles) | Dark blue `#1F4E79` | White bold |
| Column headers (FY2024A, FY2025E, etc.) | Light blue `#D9E1F2` | Black bold |
| Input cells (historicals, assumption drivers) | Light grey `#F2F2F2` or white | Blue `#0000FF` |
| Formula cells | White | Black |
| Cross-tab links | White | Green `#008000` |
| Check rows / key totals | Medium blue `#BDD7EE` | Black bold |

**That's 3 blues + 1 grey + white.** If the template has its own color scheme, follow the template instead.

Font color signals *what* a cell is (input/formula/link). Fill color signals *where* you are (header/data/check).

## Model Structure

### Identifying Template Tab Organization

Templates vary in their tab naming conventions and organization. Before populating, review all tabs to understand the template's structure. Below are common tab names and their typical contents:

| Common Tab Names | Contents to Look For |
|------------------|----------------------|
| IS, P&L, Income Statement | Income Statement |
| BS, Balance Sheet | Balance Sheet |
| CF, CFS, Cash Flow | Cash Flow Statement |
| WC, Working Capital | Working Capital Schedule |
| DA, D&A, Depreciation, PP&E | Depreciation & Amortization Schedule |
| Debt, Debt Schedule | Debt Schedule |
| NOL, Tax, DTA | Net Operating Loss Schedule |
| Assumptions, Inputs, Drivers | Driver assumptions and inputs |
| Checks, Audit, Validation | Error-checking dashboard |

**Template Review Checklist**
- Identify which tabs exist in the template (not all templates include every schedule)
- Note any template-specific tabs not listed above
- Understand tab dependencies (e.g., which schedules feed into the main statements)
- Locate input cells vs. formula cells on each tab

### Understanding Template Structure

Before populating a template, familiarize yourself with its existing layout to ensure data is entered in the correct locations and formulas remain intact.

**Identifying Row Structure**
- Locate the model title at top of each tab
- Identify section headers and their visual separation
- Find the units row indicating $ millions, %, x, etc.
- Note column headers distinguishing Actuals vs. Estimates periods
- Confirm period labels (e.g., FY2024A, FY2025E)
- Identify input cells vs. formula cells (typically distinguished by font color)

**Identifying Column Structure**
- Confirm line item labels in leftmost column
- Verify historical years precede projection years
- Note the visual border separating historical from projected periods
- Check for consistent column order across all tabs

**Working with Named Ranges (Python `openpyxl`):**
- When populating a template, inspect `wb.defined_names` programmatically (`{k: v.attr_text for k, v in wb.defined_names.items()}`) before modifying cells so you preserve every defined range.
- When building from scratch, define clean named ranges for key controls (e.g., `Scenario` / `Case` selector, `MinCash`) via `openpyxl.workbook.defined_name.DefinedName`.

### Projection Period
- Templates typically project 5 years forward from last historical year
- Verify historical (A) vs. projected (E) columns are clearly separated
- Confirm columns use fiscal year notation (e.g., FY2024A, FY2025E)

## Margin Analysis

**Note: The following margin analysis should only be performed if prompted by the user or if the template explicitly requires it. If no prompt is given, skip this section.**

Calculate and display profitability margins on the Income Statement (IS) tab to track operational efficiency and enable peer comparison.

### Core Margins to Include

| Margin | Formula | What It Measures |
|--------|---------|------------------|
| Gross Margin | Gross Profit / Revenue | Pricing power, production efficiency |
| EBITDA Margin | EBITDA / Revenue | Core operating profitability |
| EBIT Margin | EBIT / Revenue | Operating profitability after D&A |
| Net Income Margin | Net Income / Revenue | Bottom-line profitability |

### Income Statement Layout with Margins

Display margin percentages directly below each profit line item:
- Gross Margin % below Gross Profit
- EBIT Margin % below EBIT
- EBITDA Margin % below EBITDA
- Net Income Margin % below Net Income

## Credit Metrics

**Note: The following Credit analysis should only be performed if prompted by the user or if the template explicitly requires it. If no prompt is given, skip this section.**

Calculate and display credit/leverage metrics on the Balance Sheet (BS) tab to assess financial health, debt capacity, and covenant compliance.

### Core Credit Metrics to Include

| Metric | Formula | What It Measures |
|--------|---------|------------------|
| Total Debt / EBITDA | Total Debt / LTM EBITDA | Leverage multiple |
| Net Debt / EBITDA | (Total Debt - Cash) / LTM EBITDA | Leverage net of cash |
| Interest Coverage | EBITDA / Interest Expense | Ability to service debt |
| Debt / Total Cap | Total Debt / (Total Debt + Equity) | Capital structure |
| Debt / Equity | Total Debt / Total Equity | Financial leverage |
| Current Ratio | Current Assets / Current Liabilities | Short-term liquidity |
| Quick Ratio | (Current Assets - Inventory) / Current Liabilities | Immediate liquidity |

### Credit Metric Hierarchy Checks

Validate that Upside shows strongest credit profile:
- Leverage: Upside < Base < Downside (lower is better)
- Coverage: Upside > Base > Downside (higher is better)
- Liquidity: Upside > Base > Downside (higher is better)

### Covenant Compliance Tracking

If debt covenants are known, add explicit compliance checks comparing actual metrics to covenant thresholds.

## Scenario Analysis (Base / Upside / Downside)

Use a scenario toggle (dropdown) in the Assumptions tab with CHOOSE or INDEX/MATCH formulas.

| Scenario | Description |
|----------|-------------|
| Base Case | Management guidance or consensus estimates |
| Upside Case | Above-guidance growth, margin expansion |
| Downside Case | Below-trend growth, margin compression |

**Key Drivers to Sensitize**: Revenue growth, Gross margin, SG&A %, DSO/DIO/DPO, CapEx %, Interest rate, Tax rate.

**Scenario Audit Checks**: Toggle switches all statements, BS balances in all scenarios, Cash ties out, Hierarchy holds (Upside > Base > Downside for NI, EBITDA, FCF, margins).

## SEC Filings Data Extraction

If the template specifically requires pulling data from SEC filings (10-K, 10-Q), see [references/sec-filings.md](references/sec-filings.md) for detailed extraction guidance. This reference is only needed when populating templates with public company data from regulatory filings.

## Building from Scratch vs. Populating Templates (Python / `openpyxl`)

### Step 1: Analyze Structure (or Design Clean Scratch Architecture)

- **When building from scratch:** Create canonical tabs — `Assumptions` (with Scenario toggle), `Income Statement`, `Balance Sheet`, `Cash Flow Statement`, `Working Capital`, `PP&E & D&A`, `Debt & Revolver`, `Tax & NOL` (if applicable), and `Checks`. Lock all row coordinates before writing any formulas.
- **When populating a template:** Inspect all sheets, formulas, and `wb.defined_names` via `openpyxl.load_workbook(path, data_only=False)`. Map blue inputs vs. black formulas vs. green cross-sheet links, document units and sign conventions, and NEVER overwrite existing formula cells unless explicitly repairing a bug.

### Step 2: Programmatic Data Entry & Accounting Rules

- **Avoid SBC & D&A Double-Counting:** In standard US GAAP filings, D&A and Stock-Based Compensation are typically **embedded inside** functional lines (`COGS`, `R&D`, `S&M`, `G&A`). State explicitly on the `Income Statement` whether functional OpEx lines are shown **inclusive** of D&A and SBC (in which case do NOT subtract D&A or SBC again — show them as memo rows fed from schedules/CF) OR **exclusive** of D&A and SBC (in which case carve them out of historical functional OpEx before subtracting D&A and SBC separately).
- **Complete Equity Roll-Forward:** Track **Share Repurchases (Treasury Stock)** (`Prior Treasury - Buybacks`) and **Other Comprehensive Income (AOCI)** (`Prior AOCI + OCI`) alongside Retained Earnings (`Prior RE + NI - Dividends`) and Common Stock/APIC (`Prior APIC + Equity Issuance + SBC`) so Total Equity ties cleanly to the Balance Sheet.
- **Minimum Cash & Revolver Sweep:** Include a **revolver / minimum cash sweep** in the `Debt & Revolver` schedule (`Minimum Cash Balance` input; draw revolver when `Cash Before Financing < Min Cash`, pay down revolver with excess cash up to outstanding revolver balance) so Cash never goes silently negative.

### Step 3: Programmatic Formula & Linkage Validation

Inspect formula integrity in Python rather than manual GUI steps:
- **Formula consistency across columns:** Iterate across forecast columns in Python to confirm identical relative/absolute formula structure across projection years (replacing manual `Ctrl+\`).
- **Precedent / hardcode audit:** Parse formula strings to verify projection cells contain `=` formulas referencing `Assumptions` / schedule cells (no embedded hardcoded percentages or constants).
- **Recalculation & error audit:** Run `python3 scripts/recalc_xlsx.py <model.xlsx> 30 --in-place` (or `--output <recalc.xlsx>`) to populate LibreOffice-recalculated cached values on disk, and verify exit `0` (`"status": "success"`, zero `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`, zero uncached formulas).

### Step 4: Quality Checks by Sheet

Perform these validation checks on each sheet after populating the template:

**Income Statement (IS) Quality Checks**
- Revenue figures match source data for historical periods
- All expense line items sum to reported totals
- Subtotals (Gross Profit, EBIT, EBT, Net Income) calculate correctly
- Tax calculation logic is appropriate (handles losses correctly)
- Forecast drivers reference assumptions tab (no hardcodes)
- Period-over-period changes are directionally reasonable

**Balance Sheet (BS) Quality Checks**
- Assets = Liabilities + Equity for every period (primary check)
- Cash balance matches Cash Flow Statement ending cash
- Working capital accounts tie to supporting schedules (if applicable)
- Retained Earnings rolls forward correctly: Prior RE + Net Income - Dividends = Ending RE
- Common Stock / APIC rolls forward correctly: Prior Common Stock/APIC + Equity Issuance + SBC = Ending Common Stock/APIC
- Debt balances tie to debt schedule (if applicable)
- All balance sheet items have appropriate signs (assets positive, most liabilities positive)

**Cash Flow Statement (CF) Quality Checks**
- Net Income at top of CFO matches Income Statement Net Income
- Non-cash add-backs (D&A, SBC, etc.) tie to their source schedules/statements
- Working capital changes have correct signs (increase in asset = use of cash = negative)
- CapEx ties to PP&E schedule or fixed asset roll-forward
- Financing activities tie to changes in debt and equity accounts on BS
- Ending Cash matches Balance Sheet Cash
- Beginning Cash equals prior period Ending Cash

**Supporting Schedule Quality Checks**
- Opening balances equal prior period closing balances
- Roll-forward logic is complete (Beginning + Additions - Deductions = Ending)
- Schedule totals tie to main statement line items
- Assumptions used in calculations match Assumptions tab

### Step 5: Cross-Statement Integrity Checks

After validating individual sheets, confirm the three statements are properly integrated:

| Check | Formula | Expected Result |
|-------|---------|-----------------|
| Balance Sheet Balance | Assets - Liabilities - Equity | = 0 |
| Cash Tie-Out | CF Ending Cash - BS Cash | = 0 |
| Net Income Link | IS Net Income - CF Starting Net Income | = 0 |
| Retained Earnings | Prior RE + NI - Dividends - BS Ending RE | = 0 |
| Common Stock / APIC | Prior Common Stock/APIC + Equity Issuance + SBC - BS Ending Common Stock/APIC | = 0 |

### Step 6: Final Review

Before considering the model complete:
- Toggle through all scenarios (if applicable) to verify checks pass in each case
- Review all #REF!, #DIV/0!, #VALUE!, and #NAME? errors and resolve or document
- Confirm all input cells have been populated (search for placeholder values)
- Verify units are consistent across all tabs
- Save a clean version before making any additional modifications

## Model Validation and Audit

This section consolidates all validation checks and audit procedures for completed templates.

### Core Linkages (Must Always Hold)

See [references/formulas.md](references/formulas.md) for all formula details.

| Check | Formula | Expected Result |
|-------|---------|-----------------|
| Balance Sheet Balance | `Assets - Liabilities - Equity` | `= 0` |
| Cash Tie-Out | `CF Ending Cash - BS Cash` | `= 0` |
| Net Income Link | `IS Net Income - CF Starting Net Income` | `= 0` |
| Retained Earnings | `Prior RE + NI - Dividends - BS Ending RE` | `= 0` |
| Common Stock / APIC | `ΔCommon Stock/APIC (BS) - Equity Issuance (CFF) - SBC (IS)` | `= 0` |
| Treasury Stock / Buybacks | `ΔTreasury Stock (BS) + Share Repurchases (CFF)` | `= 0` |
| AOCI Roll-Forward | `Prior AOCI + OCI - BS Ending AOCI` | `= 0` |
| Revolver / Min Cash Floor | `BS Ending Cash >= Min Cash Target` | `TRUE` |
| *Monthly Roll-Up (if monthly tabs exist)* | `Closing Cash (Monthly) - Closing Cash (Annual)` | `= 0` |

### Sign Convention Reference

| Statement | Item | Sign Convention |
|-----------|------|-----------------|
| CFO | D&A, SBC | Positive (add-back) |
| CFO | ΔAR / ΔInventory (increase) | Negative (use of cash) |
| CFO | ΔAP / ΔDeferred Revenue (increase) | Positive (source of cash) |
| CFI | CapEx | Negative |
| CFF | Debt / Revolver draw | Positive |
| CFF | Debt / Revolver repayment | Negative |
| CFF | Share Repurchases / Dividends | Negative |

### Circular Reference Handling (Python / `openpyxl`)

Average-balance interest creates circularity (`Interest → Net Income → Cash → Revolver/Debt → Interest`):
- **Preferred (non-circular default):** Compute projected interest expense and interest income on **beginning-of-period debt and cash balances** — completely avoids circular references and guarantees deterministic LibreOffice / Excel recalculation.
- **If average-balance iteration is explicitly required:** Add a circuit-breaker toggle cell (`0 = Beginning Balance, 1 = Average Balance`) in `Assumptions`, AND enable workbook iteration programmatically in `openpyxl`:
  ```python
  wb.calculation.iterate = True
  wb.calculation.iterateCount = 100
  wb.calculation.iterateDelta = 0.001
  ```

### Check Categories

**Section 1: Currency & Units Consistency** — Consistent currency and scale across all tabs.

**Section 2: Balance Sheet Integrity** — `Assets - Liabilities - Equity = 0` in every historical and projected period.

**Section 3: Cash Flow Integrity** — CF Ending Cash `= BS Cash`; CF Starting NI `= IS NI`; D&A, SBC, ΔNWC, CapEx, Debt/Revolver, Buybacks, and Dividends all tie to schedules. *(Include Monthly-vs-Annual cash check only in templates with monthly tabs.)*

**Section 4: Equity Roll-Forward** — Retained Earnings (`Prior RE + NI - Dividends`), Common Stock/APIC (`Prior APIC + Equity Issuance + SBC`), Treasury Stock (`Prior Treasury - Share Repurchases`), and AOCI (`Prior AOCI + OCI`) all reconcile to BS.

**Section 5: Working Capital** — AR, Inventory, AP, and Deferred Revenue tie to BS; DSO/DIO/DPO within reasonable ranges.

**Section 6: Debt & Revolver Schedule** — Total Debt ties to BS (Current + LT + Revolver); minimum cash sweep prevents negative cash; interest ties to IS.

**Section 6b: Tax & NOL Schedule (ASC 740)**
- **Established public companies:** Anchor Beginning NOL and Valuation Allowance directly to the latest `10-K` Income Tax footnote (do NOT force Beginning NOL `= 0` for established companies with existing carryforwards!).
- **New-formation / startup models only:** Beginning NOL (Year 1 / Formation) `= 0`.
- NOL increases only when `EBT < 0`; post-2017 federal utilization `<= 80%` of taxable income; Ending NOL `>= 0`.
- `Current Tax Payable = MAX(0, Taxable Income × Tax Rate)`; GAAP `Deferred Tax Expense / (Benefit) = -ΔDTA` on IS and `-ΔDTA` on CFO (or `0` on both under a `100%` valuation allowance).

**Section 7: Scenario Hierarchy** — Upside > Base > Downside for Revenue, EBITDA, NI, FCF, margins; Upside < Base < Downside for leverage.

**Section 8: Formula Integrity** — Zero `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`; zero uncached formulas (`scripts/recalc_xlsx.py` exits `0`).

**Section 9: Credit Metric Thresholds** — Flag covenant breaches clearly.

### Master Check Formula

Aggregate all section statuses into a single master check:
- If all sections pass → "✓ ALL CHECKS PASS"
- If any section fails → "✗ ERRORS DETECTED - REVIEW BELOW"

### Quick Debug Workflow

When Master Status shows errors:
1. Scroll to find red-highlighted sections
2. Identify which check category has failures
3. Navigate to source tab to investigate
4. Fix the underlying issue
5. Return to Checks tab to verify resolution
