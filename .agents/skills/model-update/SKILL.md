---
name: model-update
description: Update financial models with new data — quarterly earnings, management guidance, macro changes, or revised assumptions. Adjusts estimates, recalculates valuation, and flags material changes. Use after earnings, guidance updates, or when assumptions need refreshing. Triggers on "update model", "plug earnings", "refresh estimates", "update numbers for [company]", "new guidance", or "revise estimates".
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/model-update/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# Model Update

## Workflow

### Step 1: Identify What Changed

Determine the update trigger:
- **Earnings release**: New quarterly actuals to plug in
- **Guidance change**: Company updated forward outlook
- **Estimate revision**: Analyst changing assumptions based on new data
- **Macro update**: Interest rates, FX, commodity prices changed
- **Event-driven**: M&A, restructuring, new product, management change

### Step 2: Plug New Data

#### After Earnings
Update the model with reported actuals:

| Line Item | Prior Estimate | Actual | Delta | Notes |
|-----------|---------------|--------|-------|-------|
| Revenue | | | | |
| Gross Margin | | | | |
| Operating Expenses | | | | |
| EBITDA | | | | |
| EPS | | | | |
| [Key metric 1] | | | | |
| [Key metric 2] | | | | |

**Segment Detail** (if applicable):
- Update each segment's revenue and margin
- Note any segment mix shifts

**Balance Sheet / Cash Flow Updates**:
- Cash and debt balances
- Share count (buybacks, dilution)
- Capex actual vs. estimate
- Working capital changes

### Step 3: Revise Forward Estimates

Based on the new data, adjust forward estimates:

| | Old FY Est | New FY Est | Change | Old Next FY | New Next FY | Change |
|---|-----------|-----------|--------|------------|------------|--------|
| Revenue | | | | | | |
| EBITDA | | | | | | |
| EPS | | | | | | |

**Key Assumption Changes:**
- What assumptions are you changing and why?
- Revenue growth rate: old → new (reason)
- Margin assumption: old → new (reason)
- Any new items (restructuring charges, one-time gains, etc.)

### Step 4: Valuation Impact

Recalculate valuation with updated estimates:

| Valuation Method | Prior | Updated | Change |
|-----------------|-------|---------|--------|
| DCF fair value | | | |
| P/E (NTM EPS × target multiple) | | | |
| EV/EBITDA (NTM EBITDA × target multiple) | | | |
| **Price Target** | | | |

### Step 5: Summary & Action

**Estimate Change Summary:**
- One paragraph: what changed, why, and what it means for the stock
- Is this a thesis-changing event or noise?

**Rating / Price Target:**
- Maintain or change rating?
- New price target (if changed) with methodology
- Upside/downside to current price

### Step 6: Output

- **Excel Feature Preservation Warning (`openpyxl` & LibreOffice):** Loading and saving an existing analyst workbook with `openpyxl` (and any subsequent LibreOffice re-export) can strip or alter Excel-only features — including **charts, pivot tables, slicers/timelines, shapes/text boxes/images, form/ActiveX controls, sparklines, VBA macros (`.xlsm`), What-If data tables, and newer-Excel conditional formatting**. Before editing an existing workbook, inspect it with read-only `python3 scripts/recalc_xlsx.py <original.xlsx>` (`--check-only`), which reports `lossy_features`, `data_tables`, and `lossy_features_inspected` without modifying the file. Never overwrite the prior-quarter source workbook:
  - **Interactive mode:** If `lossy_features` or `data_tables` are non-empty **or `lossy_features_inspected` is `false`** (e.g., legacy `.xls` or `.ods` inputs that cannot be pre-inspected), **pause and ask the user before modifying**.
  - **Subagent mode (e.g., delegated via `invoke_subagent` inside `earnings-reviewer`):** Do **not** stall waiting for interactive input. Leave the original source workbook 100% untouched, save the updated model to a new versioned canonical copy populated with cached formula values below, and prominently list the untouched source path plus any stripped/uninspected `lossy_features` (`lossy_feature_details` / `lossy_features_inspected: false`) in the final handoff summary.
- **Updated Excel model (cached canonical file on disk):** Never overwrite the prior-quarter source workbook and **never leave an uncached `openpyxl` file at the canonical filename** `research/companies/<TICKER>/models/<TICKER>_Model_YYYY-MM-DD.xlsx`:
  1. If the newly written `openpyxl` copy has `lossy_features_inspected: true`, zero `lossy_features`, and zero `data_tables`: save to the canonical versioned path and recalculate with `--in-place`:
     `python3 scripts/recalc_xlsx.py research/companies/<TICKER>/models/<TICKER>_Model_YYYY-MM-DD.xlsx 30 --in-place`
  2. If the updated copy retains a feature that causes `--in-place` to refuse (or when staging is preferred): write the uncached `openpyxl` output to a temporary staging file (e.g., `/tmp/<TICKER>_Model_YYYY-MM-DD_staged.xlsx`), write the LibreOffice-recalculated output **directly to the canonical versioned path** via `--output`:
     `python3 scripts/recalc_xlsx.py /tmp/<TICKER>_Model_YYYY-MM-DD_staged.xlsx 30 --output research/companies/<TICKER>/models/<TICKER>_Model_YYYY-MM-DD.xlsx`
     and delete the temporary staging file.
  Then confirm exit `0` (`"status": "success"`, `uncached_formulas: 0`), run `python3 .agents/skills/dcf-model/scripts/validate_dcf.py research/companies/<TICKER>/models/<TICKER>_Model_YYYY-MM-DD.xlsx` whenever a `DCF` tab is present, and read all note/summary metrics strictly from **that canonical recalculated file**. *(For read-only audits of untouched existing workbooks without `openpyxl` edits, use the default non-destructive `python3 scripts/recalc_xlsx.py <file.xlsx>` / `--check-only`.)*
- Estimate change summary & updated price target derivation (`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY`): `research/companies/<TICKER>/notes/<TICKER>_Model_Update_YYYY-MM-DD.md` (markdown or Word)

## Important Notes

- Always reconcile your estimates to the company's reported figures before projecting forward
- Always inspect existing workbooks for Excel-only features (`lossy_features`, `data_tables`, and `lossy_features_inspected`) first and never overwrite the original file — in interactive mode ask the user before modifying when `lossy_features`/`data_tables` are non-empty or `lossy_features_inspected` is `false`; in subagent mode proceed on a new versioned canonical copy and surface all untouched-source / `lossy_features` details prominently in the handoff
- Note any non-recurring items and whether your estimates are GAAP or adjusted
- Track your estimate revision history — it shows your analytical progression
- If the quarter was noisy, separate signal from noise in your estimate changes
- Check consensus after updating — how do your revised estimates compare to the Street?
- Share count matters — dilution from stock comp, converts, or buybacks can materially affect EPS
