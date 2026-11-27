---
trigger: model_decision
description: Apply whenever creating, modifying, populating, or auditing Excel financial models (.xlsx), DCF valuations, 3-statement models, comparable company analyses, unit economics workbooks, or Python financial modeling scripts.
---

# Financial Modeling & Spreadsheet Standards

1. **Formulas over hardcodes (non-negotiable).** Every projection, margin, growth rate, subtotal, multiple, discount factor, equity bridge, and sensitivity cell MUST be a live Excel formula (`ws["D20"] = "=D19*(1+$B$8)"`). Never compute derived outputs in Python and write static numbers into calculation cells.
2. **Institutional color coding.**
   - **Blue font (`#0000FF`)**: Hardcoded historical actuals and explicit assumption inputs.
   - **Black font (`#000000`)**: Same-sheet formulas.
   - **Green font (`#008000`)**: Cross-sheet or cross-workbook links.
   - Keep fill palettes restrained (dark navy `#1F4E79` section headers with white bold text; soft blue `#D9E1F2` column headers; light grey `#F2F2F2` totals).
3. **Document every input.** Attach cell comments on hardcoded inputs stating source, filing date, and unit, or consolidate sources in a dedicated `Notes & Sources` block.
4. **Stage checkpoints on complex builds.** On 3-statement and DCF models, verify inputs, historicals, projections, balance checks, and valuation bridges incrementally rather than presenting an unaudited black-box workbook.
5. **Mandatory automated QA before delivery.**
   - Run `python3 scripts/recalc_xlsx.py <workbook.xlsx>` to recalculate cached formula values via LibreOffice (if installed) and scan every tab for `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, and `#NUM!`.
   - For DCF workbooks, also run `python3 .agents/skills/dcf-model/scripts/validate_dcf.py <workbook.xlsx>`.
   - Verify accounting integrity: Balance Sheet balances (`Assets - Liabilities - Equity = 0`) for every period, Cash Flow ending cash ties to Balance Sheet cash, and sensitivity table center cells match the base case.
6. **Reproducibility.** Save or document the Python generator script so models can be audited and rolled forward cleanly next quarter.
