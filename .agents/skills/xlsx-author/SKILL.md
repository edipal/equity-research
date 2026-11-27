---
name: xlsx-author
description: Build institutional-quality Excel (.xlsx) financial models and schedules programmatically on disk using Python (openpyxl) with live formulas, institutional color coding, balance checks, and LibreOffice formula verification. Use when generating or updating standalone .xlsx workbooks in Antigravity.
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/xlsx-author/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# xlsx-author

Build institutional-grade Excel (`.xlsx`) workbooks programmatically in Antigravity using Python (`openpyxl`), with verified live formulas and zero hardcoded calculation cells.

## Output Contract

- Default output location:
  - Single-company models: `research/companies/<TICKER>/models/<name>.xlsx`
  - Sector / peer comps: `research/sectors/<sector-slug>/comps/<name>.xlsx`
  - Stock screens / cross-sector schedules: `research/screens/<name>.xlsx` (create parent directories if needed)
- Always save the generating Python script alongside or in a clearly documented path so the model can be refreshed reproducibly.
- Return the workspace file path and a formula-audit summary in your final response.

## How to Build the Workbook

Write a clean Python script using `openpyxl` and execute it in the terminal:

```python
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
ws = wb.active
ws.title = "Inputs"
ws["B2"] = "Revenue"
ws["C2"] = 1_250_000_000
ws["C2"].font = Font(color="0000FF")          # Blue = hardcoded input

calc = wb.create_sheet("DCF")
calc["C5"] = "=Inputs!C2*(1+Inputs!C3)"       # Green/Black = live formula
calc["C5"].font = Font(color="008000")        # Green = cross-sheet link

out_path = Path("research/companies/TICKER/models/model.xlsx")
out_path.parent.mkdir(parents=True, exist_ok=True)
wb.save(out_path)
```

## Institutional Conventions (Mirror `audit-xls`)

- **Color coding**: Blue (`#0000FF`) = hardcoded input, Black (`#000000`) = same-sheet formula, Green (`#008000`) = cross-sheet link.
- **Zero hardcodes in calculation cells**: Every derived value (growth rates, margins, subtotals, multiples, discount factors, sensitivity cells) must be an Excel formula.
- **Source comments**: Attach cell comments (`openpyxl.comments.Comment`) documenting source, filing date, and unit on every blue input cell.
- **Named ranges**: Define named ranges for key summary outputs referenced in memos or slide decks.
- **Checks tab**: Include a `Checks` tab validating balance sheet balance (`Assets - L - E = 0`), cash flow tie-out, and error counts (`TRUE`/`FALSE`).
- **One model per file**: Do not overwrite or append unrelated models into an existing workbook unless explicitly asked.

## Mandatory Verification Step

After saving a newly generated `.xlsx` file with `openpyxl`, always populate cached formula values via LibreOffice and scan for errors by running:

```bash
python3 scripts/recalc_xlsx.py <path-to-workbook.xlsx> 30 --in-place
```

*(Because `recalc_xlsx.py` is **non-destructive by default** / `--check-only`, omitting `--in-place` or `--output <PATH>` leaves `openpyxl`-written files uncached on disk and reports `on_disk_uncached_formulas > 0` with a warning. Use `--in-place` for newly Python-generated `.xlsx` files, `--output <recalc.xlsx>` when preserving an existing workbook, and confirm `on_disk_uncached_formulas == 0` before reading with `openpyxl(data_only=True)` or running `validate_dcf.py`.)*

This recalculates cached formula values via LibreOffice and scans every sheet for all 7 Excel errors (`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`):
- Exit `0` (`"status": "success"`): zero errors and verified formulas (`uncached_formulas: 0`, `on_disk_uncached_formulas: 0` when persisted via `--in-place` or `--output`).
- Exit `1` (`"status": "errors_found"` / `"error"`): formula errors or load/LibreOffice failure — fix and re-run.
- Exit `2` (`"status": "unverified"`): LibreOffice unavailable and workbook contains uncached formulas (install LibreOffice or pass `--allow-unrecalculated` only if explicitly accepted).

Then apply the `audit-xls` checklist before presenting the workbook.

## Optional Google Sheets Sync

If the user asks for a Google Sheets deliverable and the `gsheets` / `gdrive` plugins are connected, verify the local `.xlsx` first, then upload or sync to the requested Drive folder or spreadsheet.
