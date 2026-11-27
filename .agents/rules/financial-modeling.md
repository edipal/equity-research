---
trigger: model_decision
description: Apply whenever creating, modifying, populating, or auditing Excel financial models (.xlsx), DCF valuations, 3-statement models, comparable company analyses, unit economics workbooks, or Python financial modeling scripts.
---

# Financial Modeling & Spreadsheet Standards

1. **Formulas over hardcodes (non-negotiable).** Every projection, margin, growth rate, subtotal, multiple, discount factor, equity bridge, and sensitivity cell MUST be a live Excel formula (`ws["D20"] = "=D19*(1+$B$8)"`). Never compute derived outputs in Python and write static numbers into calculation cells.
2. **Institutional color coding & unified fill palette.**
   - **Font colors:** Blue (`#0000FF`) = hardcoded historical actuals & assumption inputs; Black (`#000000`) = same-sheet formulas; Green (`#008000`) = cross-sheet or cross-workbook links.
   - **Fill palette (unless a user template overrides):**
     - Dark navy (`#1F4E79`) with white bold text = section headers / banners
     - Light blue (`#D9E1F2`) with black bold text = column headers / sub-headers
     - Light grey (`#F2F2F2`) = input cells / summary statistics shading
     - Medium blue (`#BDD7EE`) with bold text = key totals, checks, active scenario selector, and sensitivity center-cell highlights
3. **Document every input.** Attach cell comments on hardcoded inputs stating source, filing date, and unit (or `[ASSUMPTION]` for analyst estimates), and consolidate methodology notes on the sheet or in a `Sources & Notes` tab.
4. **Stage checkpoints on complex builds.** On 3-statement and DCF models running interactively, verify inputs, historicals, projections, balance checks, and valuation bridges incrementally before finalizing. When invoked as an autonomous subagent, tag defaults `[ASSUMPTION]` and surface the full assumption summary and audit report in the final handoff.
5. **Mandatory automated QA before delivery.**
   - Run `python3 scripts/recalc_xlsx.py <workbook.xlsx> 30` (or `--check-only` for read-only audits; `--output <PATH>` for non-destructive writes) to recalculate cached formula values via LibreOffice and scan every sheet for all 7 Excel errors (`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`). Confirm exit `0` (`"status": "success"`). If exit `2` (`"status": "unverified"`, LibreOffice unavailable), install LibreOffice or explicitly disclose unverified cached values (only passing `--allow-unrecalculated` with analyst consent).
   - For DCF workbooks, also run `python3 .agents/skills/dcf-model/scripts/validate_dcf.py <workbook.xlsx>` and confirm exit `0` (`"status": "PASS"`). You can bootstrap or compare against `.agents/skills/dcf-model/scripts/build_dcf_skeleton.py`.
   - Verify accounting integrity: Balance Sheet balances (`Assets - Liabilities - Equity = 0`) for every period, Cash Flow ending cash ties to Balance Sheet cash, gross-debt WACC weights sum to `100%` (`0%–100%`), exit-multiple TV is discounted at `(1 + WACC)^N`, and sensitivity table center cells match the base case.
6. **Reproducibility.** Save or document the Python generator script (creating parent directories via `Path(...).parent.mkdir(parents=True, exist_ok=True)`) so models can be audited and rolled forward cleanly next quarter.
