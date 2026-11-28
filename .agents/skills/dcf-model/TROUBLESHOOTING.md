<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/dcf-model/TROUBLESHOOTING.md), Apache-2.0; modified for Google Antigravity. -->

# DCF Model Troubleshooting Guide

**When to read this file:** If `scripts/recalc_xlsx.py` reports `"errors_found"`, `"error"`, or `"unverified"`, if `validate_dcf.py` reports `"FAIL"`, if valuation results look unreasonable, or if the `Case` selector does not update projections.

## 1. Recalculation & Validation Failures

### `status: "unverified"` (Exit Code 2) or Uncached Formulas
- **Cause:** Either LibreOffice (`soffice`) is not installed on `PATH`, OR `recalc_xlsx.py` was run in its default non-destructive (`--check-only`) mode on an `openpyxl`-written file so cached values were never written back to disk before `validate_dcf.py` ran.
- **Fix:** Ensure LibreOffice is installed (`sudo apt-get install libreoffice-calc` or macOS `brew install --cask libreoffice`) and run `python3 scripts/recalc_xlsx.py <model.xlsx> 30 --in-place` (for newly generated `.xlsx` files) or `--output <recalc.xlsx>` (when preserving the original file), then run `validate_dcf.py` on that recalculated file. Only pass `--allow-unrecalculated` when the analyst explicitly accepts unverified cached values, and disclose `recalculated_with_libreoffice: false`.

### `Refusing to overwrite ... Excel What-If data table(s)` or LibreOffice-Lossy Features (Exit Code 1)
- **Cause:** `--in-place` (or `--output` pointing to the same file) was passed on an existing workbook containing Excel What-If data tables (`<f t="dataTable">`) or other Excel features that LibreOffice re-export can drop or alter (`lossy_features` + `lossy_feature_details`, such as charts, pivot tables, slicers/timelines, drawings/shapes, form/ActiveX controls, sparklines, VBA macros, or worksheet extensions).
- **Fix:** Run `python3 scripts/recalc_xlsx.py <model.xlsx>` (default non-destructive / `--check-only`) for safe read-only auditing without mutating the workbook, or pass `--output <new.xlsx>` if a re-exported copy is explicitly desired (inspect `lossy_features`, `lossy_feature_details`, and `lossy_features_inspected` first). For newly built DCF models, always use live `IF(WACC<=g,"NM", SUMPRODUCT(...))` sensitivity formulas (never Excel What-If data tables).

### `#REF!` Errors
- **Cause:** Formulas reference wrong or deleted rows (usually from inserting headers after writing formulas), or an out-of-bounds `INDEX` row offset (e.g., `Case` outside `1..3` or the stacked scenario range `C$26:C$58` shorter than the 33 rows spanning Bear/Base/Bull metrics: `(3 - 1) * BLOCK_STRIDE + 9 = 33`).
- **Fix:** Lock every section's row constants first (see `references/dcf-layout-and-formulas.md` and `scripts/build_dcf_skeleton.py`), verify `Case` is `1`, `2`, or `3` with a 33-row `INDEX` range (`C$26:C$58`), then regenerate the workbook cleanly.

### `#DIV/0!` Errors
- **Cause:** Division by zero (e.g., `WACC - g = 0`, zero shares, zero total capital, or zero EBITDA).
- **Fix:** Ensure `g < WACC`, verify diluted shares `> 0`, and guard ratio cells with `=IF(divisor=0,"NM",num/divisor)`.

### `#VALUE!`, `#NAME?`, `#NUM!`, `#N/A`, `#NULL!` Errors
- `#VALUE!`: Text string stored where a numeric input is expected (convert formatted numbers to floats).
- `#NAME?`: Misspelled Excel function name or missing `Case` defined name (`wb.defined_names["Case"] = DefinedName("Case", attr_text="DCF!$B$4")`).
- `#NUM!`: Invalid exponent or negative base (check WACC and discount periods).
- `#N/A`: Lookup key not found (`VLOOKUP`, `XLOOKUP`, or `MATCH` mismatch, or an explicit `NA()` placeholder).
- `#NULL!`: Accidental space intersection operator in a range reference (missing comma or colon).

## 2. Valuation Seems Unreasonable

### Implied Share Price Far Too High
- Check **TV `%` of EV** (`DCF!C113`): target `50%–70%` (`validate_dcf.py` warns if `< 50%` or `> 70%`). If `> 70%`, review `g` (must be `<= 3.5%` and `< WACC`) or extend explicit projection horizon to 7–10 years.
- Confirm Exit Multiple TV is discounted strictly at `(1 + WACC)^N` (NOT `N - 0.5`; only mid-year Perpetuity Growth TV uses `N - 0.5`).
- Confirm SBC is NOT added back to UFCF without modeling full dilution/repurchase offset.
- Check whether `Net Debt / EV` was accidentally used in WACC (must use gross-capitalization weights `E/(E+D+P)`, `D/(E+D+P)`, `P/(E+D+P)`).

### Implied Share Price Far Too Low
- Verify the **Enterprise-to-Equity bridge** signs (`- Gross Debt - Preferred - NCI + Cash + Equity Investments`).
- Confirm **ASC 842 operating lease liabilities** were not double-counted (do not deduct them from EV if FCF already deducts operating lease rent).
- Check WACC (`WACC!C30`) against peer range (`5%–20%`) and verify CapEx / NWC signs.

## 3. Case Selector Not Working

### Active Case Block Not Updating When Switching `Case` (`DCF!B4`)
- Confirm `DCF!B4` holds an integer (`1`, `2`, or `3`) with `1..3` `DataValidation`, and `wb.defined_names["Case"]` points to `DCF!$B$4`.
- Verify `DCF!C4` uses `=CHOOSE(INT(Case),"Bear Case","Base Case","Bull Case")`.
- Confirm Active Case formulas use `INT($B$4)` and match the exact vertical stride of the Bear/Base/Bull blocks: `=INDEX(C$26:C$58,(INT($B$4)-1)*12+offset)`.
- Test switching `--case 1`, `--case 2`, `--case 3` and recalculating via `recalc_xlsx.py <model.xlsx> 30 --in-place` to confirm Bear < Base < Bull implied share prices.
