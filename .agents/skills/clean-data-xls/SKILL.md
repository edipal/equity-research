---
name: clean-data-xls
description: Clean up messy spreadsheet data — trim whitespace, fix inconsistent casing, convert numbers-stored-as-text, standardize dates, remove duplicates, and flag mixed-type columns. Use when data is messy, inconsistent, or needs prep before analysis. Triggers on "clean this data", "clean up this sheet", "normalize this data", "fix formatting", "dedupe", "standardize this column", "this data is messy".
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/clean-data-xls/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# Clean Data

Clean messy data in a target `.xlsx`/`.csv` file and sheet (or a specified Google Sheet range when `gsheets` is connected).

## Environment

- **Standalone `.xlsx` / `.csv` files (default in Antigravity):** Use Python (`openpyxl` / `pandas`) to inspect and write formula-based helper columns or cleaned sheets, then verify with `python3 scripts/recalc_xlsx.py <file.xlsx> --check-only` (read-only) or `--output <cleaned.xlsx>` (non-destructive). Confirm exit `0` (`"status": "success"`); if exit `2` (`"status": "unverified"`, LibreOffice absent), disclose clearly or pass `--allow-unrecalculated` only with user approval.
- **Google Sheets (when `gsheets` plugin is connected):** Export/read sheet ranges, inspect data quality, and write cleaned helper formulas cleanly.

## Workflow

### Step 1: Scope

- Confirm the target `.xlsx`/`.csv` file path and sheet name (or Google Sheet ID and tab name when `gsheets` is connected)
- If a cell range is given (e.g., `Sheet1!A1:F200`), scope to that range; otherwise use the full used range of the specified sheet
- Profile each column: detect its dominant type (text / number / date) and identify outliers

### Step 2: Detect issues

| Issue | What to look for |
|---|---|
| Whitespace | leading/trailing spaces, double spaces |
| Casing | inconsistent casing in categorical columns (`usa` / `USA` / `Usa`) |
| Number-as-text | numeric values stored as text; stray `$`, `,`, `%` in number cells |
| Dates | mixed formats in the same column (`3/8/26`, `2026-03-08`, `March 8 2026`) |
| Duplicates | exact-duplicate rows and near-duplicates (case/whitespace differences) |
| Blanks | empty cells in otherwise-populated columns |
| Mixed types | a column that's 98% numbers but has 3 text entries |
| Encoding | mojibake (`Ã©`, `â€™`), non-printing characters |
| Errors | `#REF!`, `#N/A`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#NUM!`, `#NULL!` |

### Step 3: Propose fixes

Show a summary table before changing anything:

| Column | Issue | Count | Proposed Fix |
|---|---|---|---|

### Step 4: Apply

- **Prefer formulas over hardcoded cleaned values** — where the cleaned output can be expressed as a formula (e.g. `=TRIM(A2)`, `=VALUE(SUBSTITUTE(B2,"$",""))`, `=UPPER(C2)`, `=DATEVALUE(D2)`), write the formula in an adjacent helper column rather than computing the result in Python and overwriting the original. This keeps the transformation transparent and auditable.
- Only overwrite in place with computed values when the user explicitly asks for it, or when no sensible formula equivalent exists (e.g. encoding/mojibake repair)
- For destructive operations (removing duplicates, filling blanks, overwriting originals), confirm with the user first
- After each category of fix (whitespace → casing → number conversion → dates → dedup), show the user a sample of what changed and get confirmation before moving to the next category
- Report a before/after summary of what changed
