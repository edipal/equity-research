---
name: xlsx-author
description: Build institutional-quality Excel (.xlsx) financial models and schedules programmatically on disk using Python (openpyxl) with live formulas, institutional color coding, balance checks, and LibreOffice formula verification. Use when generating or updating standalone .xlsx workbooks in Antigravity.
---

# xlsx-author

Build institutional-grade Excel (`.xlsx`) workbooks programmatically in Antigravity using Python (`openpyxl`), with verified live formulas and zero hardcoded calculation cells.

## Output Contract

- Default output location:
  - Single-company models: `research/companies/<TICKER>/models/<name>.xlsx`
  - Sector / peer comps: `research/sectors/<sector-slug>/comps/<name>.xlsx`
  - Ad-hoc outputs: `./out/<name>.xlsx` (create parent directories if needed)
- Always save the generating Python script alongside or in a clearly documented path so the model can be refreshed reproducibly.
- Return the workspace file path and a formula-audit summary in your final response.

## How to Build the Workbook

Write a clean Python script using `openpyxl` and execute it in the terminal:

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

wb = Workbook()
ws = wb.active
ws.title = "Inputs"
ws["B2"] = "Revenue"
ws["C2"] = 1_250_000_000
ws["C2"].font = Font(color="0000FF")          # Blue = hardcoded input

calc = wb.create_sheet("DCF")
calc["C5"] = "=Inputs!C2*(1+Inputs!C3)"       # Green/Black = live formula
calc["C5"].font = Font(color="008000")        # Green = cross-sheet link
wb.save("research/companies/TICKER/models/model.xlsx")
```

## Institutional Conventions (Mirror `audit-xls`)

- **Color coding**: Blue (`#0000FF`) = hardcoded input, Black (`#000000`) = same-sheet formula, Green (`#008000`) = cross-sheet link.
- **Zero hardcodes in calculation cells**: Every derived value (growth rates, margins, subtotals, multiples, discount factors, sensitivity cells) must be an Excel formula.
- **Source comments**: Attach cell comments (`openpyxl.comments.Comment`) documenting source, filing date, and unit on every blue input cell.
- **Named ranges**: Define named ranges for key summary outputs referenced in memos or slide decks.
- **Checks tab**: Include a `Checks` tab validating balance sheet balance (`Assets - L - E = 0`), cash flow tie-out, and error counts (`TRUE`/`FALSE`).
- **One model per file**: Do not overwrite or append unrelated models into an existing workbook unless explicitly asked.

## Mandatory Verification Step

After saving the `.xlsx` file, always run:

```bash
python3 scripts/recalc_xlsx.py <path-to-workbook.xlsx>
```

This recalculates cached formula values via LibreOffice (when installed) and scans every sheet for `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, and `#NUM!`. Then apply the `audit-xls` checklist before presenting the workbook.

## Optional Google Sheets Sync

If the user asks for a Google Sheets deliverable and the `gsheets` / `gdrive` plugins are connected, verify the local `.xlsx` first, then upload or sync to the requested Drive folder or spreadsheet.
