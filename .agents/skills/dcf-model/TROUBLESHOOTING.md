<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/dcf-model/TROUBLESHOOTING.md), Apache-2.0; modified for Google Antigravity. -->

# DCF Model Troubleshooting Guide

**When to read this file:** If `scripts/recalc_xlsx.py` reports `"errors_found"`, `"error"`, or `"unverified"`, if `validate_dcf.py` reports `"FAIL"`, if valuation results look unreasonable, or if the `Case` selector does not update projections.

## 1. Recalculation & Validation Failures

### `status: "unverified"` (Exit Code 2) or Uncached Formulas
- **Cause:** LibreOffice (`soffice`) is not installed on `PATH`, so `openpyxl`-written formulas have no cached values (`val_v is None`).
- **Fix:** Install LibreOffice (`sudo apt-get install libreoffice-calc` or macOS `brew install --cask libreoffice`) and re-run `python3 scripts/recalc_xlsx.py <model.xlsx> 30`. Only pass `--allow-unrecalculated` when the analyst explicitly accepts unverified cached values, and disclose `recalculated_with_libreoffice: false`.

### `#REF!` Errors
- **Cause:** Formulas reference wrong or deleted rows (usually from inserting headers after writing formulas).
- **Fix:** Lock every section's row constants first (see `references/dcf-layout-and-formulas.md` and `scripts/build_dcf_skeleton.py`), then regenerate the workbook cleanly.

### `#DIV/0!` Errors
- **Cause:** Division by zero (e.g., `WACC - g = 0`, zero shares, zero total capital, or zero EBITDA).
- **Fix:** Ensure `g < WACC`, verify diluted shares `> 0`, and guard ratio cells with `=IF(divisor=0,"NM",num/divisor)`.

### `#VALUE!`, `#NAME?`, `#NUM!`, `#N/A`, `#NULL!` Errors
- `#VALUE!`: Text string stored where a numeric input is expected (convert formatted numbers to floats).
- `#NAME?`: Misspelled Excel function name or missing `Case` defined name (`wb.defined_names["Case"] = DefinedName("Case", attr_text="DCF!$B$4")`).
- `#NUM!`: Invalid exponent or negative base (check WACC and discount periods).
- `#N/A`: Lookup/index mismatch (verify `Case` is `1`, `2`, or `3` and `INDEX` range length matches `BLOCK_STRIDE * 3`).
- `#NULL!`: Accidental space intersection operator in a range reference (missing comma or colon).

## 2. Valuation Seems Unreasonable

### Implied Share Price Far Too High
- Check **TV `%` of EV** (`DCF!C112`): if `> 80%`, terminal assumptions dominate; review `g` (must be `<= 3.5%` and `< WACC`) or extend explicit projection horizon to 7–10 years.
- Confirm Exit Multiple TV is discounted at `(1 + WACC)^N` (NOT `N - 0.5`).
- Confirm SBC is NOT added back to UFCF without modeling full dilution/repurchase offset.
- Check whether `Net Debt / EV` was accidentally used in WACC (must use gross-debt weights `E/(E+D)` and `D/(E+D)`).

### Implied Share Price Far Too Low
- Verify the **Enterprise-to-Equity bridge** signs (`- Gross Debt - Preferred - NCI + Cash + Equity Investments`).
- Confirm **ASC 842 operating lease liabilities** were not double-counted (do not deduct them from EV if FCF already deducts operating lease rent).
- Check WACC (`WACC!C27`) against peer range (`5%–20%`) and verify CapEx / NWC signs.

## 3. Case Selector Not Working

### Active Case Block Not Updating When Switching `Case` (`DCF!B4`)
- Confirm `DCF!B4` holds an integer (`1`, `2`, or `3`) and `wb.defined_names["Case"]` points to `DCF!$B$4`.
- Verify `DCF!C4` uses `=CHOOSE(Case,"Bear Case","Base Case","Bull Case")`.
- Confirm Active Case formulas use absolute `$B$4` and match the exact vertical stride of the Bear/Base/Bull blocks: `=INDEX(C$25:C$57,($B$4-1)*12+offset)`.
- Test switching `--case 1`, `--case 2`, `--case 3` and recalculating via `recalc_xlsx.py` to confirm Bear < Base < Bull implied share prices.
