#!/usr/bin/env python3
"""
DCF Model Validation Script
Validates Excel DCF models for formula errors and common DCF mistakes
"""

import sys
import json
import re
from pathlib import Path
from typing import Optional

EXCEL_ERRORS = ('#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', '#NUM!', '#N/A')
STRING_LITERAL_RE = re.compile(r'"(?:""|[^"])*"')
SENSITIVITY_VS_RE = re.compile(r'\bvs\.?\b')


def _is_number(value) -> bool:
    """Return True for numeric values while excluding Python booleans."""
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _extract_formula_text(value) -> Optional[str]:
    """Extract formula string from standard string formulas and openpyxl ArrayFormula objects."""
    if isinstance(value, str):
        return value if value.startswith('=') else None
    if type(value).__name__ == 'ArrayFormula':
        text = getattr(value, 'text', None)
        if isinstance(text, str) and text:
            return text if text.startswith('=') else f'={text}'
    return None


def _is_sensitivity_header(text: str) -> bool:
    """Return True if a label belongs to a sensitivity table header rather than a model assumption."""
    low = text.lower()
    if 'sensitivity' in low:
        return True
    if 'wacc' in low and 'growth' in low:
        return True
    if SENSITIVITY_VS_RE.search(low) and any(
        kw in low for kw in ('wacc', 'growth', 'margin', 'beta', 'multiple', 'discount')
    ):
        return True
    return False


def _is_wacc_label(text: str) -> bool:
    low = text.lower()
    if _is_sensitivity_header(low):
        return False
    return 'wacc' in low or 'weighted average cost of capital' in low


def _is_terminal_growth_label(text: str) -> bool:
    low = text.lower()
    if _is_sensitivity_header(low):
        return False
    return 'terminal' in low and 'growth' in low


class DCFModelValidator:
    """Validates DCF models for errors and quality issues"""

    def __init__(self, excel_path: str):
        try:
            import openpyxl
        except ImportError:
            raise ImportError("openpyxl not installed. Run: pip install openpyxl")

        self.excel_path = excel_path
        self.openpyxl = openpyxl

        path_obj = Path(excel_path)
        if not path_obj.exists():
            raise FileNotFoundError(f"File not found: {excel_path}")
        if not path_obj.is_file():
            raise IsADirectoryError(f"Not a file: {excel_path}")

        self.workbook_formulas = openpyxl.load_workbook(excel_path, data_only=False)
        try:
            self.workbook_values = openpyxl.load_workbook(excel_path, data_only=True)
        except Exception:
            self.workbook_formulas.close()
            raise
        self.errors = []
        self.warnings = []
        self.info = []

    def close(self) -> None:
        """Close underlying openpyxl workbooks."""
        if getattr(self, 'workbook_formulas', None) is not None:
            self.workbook_formulas.close()
        if getattr(self, 'workbook_values', None) is not None:
            self.workbook_values.close()

    def __enter__(self) -> "DCFModelValidator":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()

    def validate_all(self) -> dict:
        """
        Run all validation checks

        Returns:
            Dict with validation results
        """
        from datetime import datetime

        self.check_sheet_structure()
        self.check_formula_errors()
        self.check_dcf_logic()

        results = {
            'file': self.excel_path,
            'validation_date': datetime.now().isoformat(),
            'status': 'PASS' if len(self.errors) == 0 else 'FAIL',
            'error_count': len(self.errors),
            'warning_count': len(self.warnings),
            'errors': self.errors,
            'warnings': self.warnings,
            'info': self.info
        }

        return results

    def check_sheet_structure(self):
        """Verify required sheets exist (2-sheet DCF + WACC spec with inline or separate sensitivity)."""
        required_sheets = ['DCF', 'WACC']
        sheet_names = self.workbook_values.sheetnames

        for sheet in required_sheets:
            if sheet not in sheet_names:
                self.warnings.append(f"Recommended sheet missing: {sheet}")
            else:
                self.info.append(f"Found sheet: {sheet}")

        has_sensitivity_sheet = any('sensitivity' in name.lower() for name in sheet_names)
        has_inline_sensitivity = False
        if not has_sensitivity_sheet and 'DCF' in sheet_names:
            dcf_ws = self.workbook_formulas['DCF']
            for row in dcf_ws.iter_rows(max_row=250, max_col=25):
                for cell in row:
                    if isinstance(cell.value, str) and _is_sensitivity_header(cell.value):
                        has_inline_sensitivity = True
                        break
                if has_inline_sensitivity:
                    break

        if has_sensitivity_sheet:
            self.info.append("Found dedicated Sensitivity sheet")
        elif has_inline_sensitivity:
            self.info.append("Found inline sensitivity analysis on DCF sheet")
        else:
            self.warnings.append(
                "Sensitivity analysis not found (expected inline on DCF sheet or on a Sensitivity sheet)"
            )

    def check_formula_errors(self):
        """Check for Excel formula errors and uncached formulas in all sheets."""
        error_details = {err: [] for err in EXCEL_ERRORS}
        total_errors = 0
        total_formulas = 0
        uncached_formulas = 0

        for sheet_name in self.workbook_values.sheetnames:
            ws_values = self.workbook_values[sheet_name]
            ws_formulas = self.workbook_formulas[sheet_name]

            for row in ws_formulas.iter_rows():
                for formula_cell in row:
                    coord = formula_cell.coordinate
                    value_cell = ws_values[coord]
                    val_f = formula_cell.value
                    val_v = value_cell.value

                    formula_text = _extract_formula_text(val_f)
                    is_formula = formula_text is not None
                    if is_formula:
                        total_formulas += 1
                        if val_v is None and getattr(value_cell, 'data_type', None) not in ('s', 'str'):
                            uncached_formulas += 1

                    # Detect Excel error without false-positiving on prose notes or string literals
                    detected_err = None
                    if getattr(value_cell, 'data_type', None) == 'e':
                        detected_err = str(val_v) if val_v is not None else '#VALUE!'
                    elif isinstance(val_v, str) and val_v in EXCEL_ERRORS:
                        detected_err = val_v
                    elif is_formula:
                        stripped_formula = STRING_LITERAL_RE.sub('', formula_text)
                        for err in EXCEL_ERRORS:
                            if err in stripped_formula:
                                detected_err = err
                                break

                    if detected_err is not None:
                        location = f"{sheet_name}!{coord}"
                        error_details.setdefault(detected_err, []).append(location)
                        total_errors += 1
                        self.errors.append(f"{detected_err} at {location}")

        if uncached_formulas > 0:
            self.errors.append(
                f"Uncached formula values detected ({uncached_formulas} of {total_formulas} formula "
                "cell(s) have no cached value). Recalculate workbook via scripts/recalc_xlsx.py first."
            )

        # Record summary metrics in self.info so self.errors holds only distinct errors
        self.info.append(f"Total formulas: {total_formulas}")
        self.info.append(f"Uncached formulas: {uncached_formulas}")
        if total_errors == 0:
            self.info.append("✓ No formula errors found")
        else:
            self.info.append(f"Total formula errors: {total_errors}")

        return error_details, total_errors

    def check_dcf_logic(self):
        """Validate DCF-specific logic and calculations"""
        self._check_terminal_growth_vs_wacc()
        self._check_wacc_range()
        self._check_terminal_value_proportion()

    def _collect_adjacent_rates(self, sheet, label_cell, allow_zero: bool = False):
        """Collect all rate values (0 < v < 1, or 0 <= v < 1) adjacent to a label cell."""
        entries = []
        for offset in range(1, 6):
            adj_cell = sheet.cell(label_cell.row, label_cell.column + offset)
            val = adj_cell.value
            if _is_number(val):
                in_range = (0 <= val < 1) if allow_zero else (0 < val < 1)
                if in_range:
                    entries.append((adj_cell.column, float(val), f"{sheet.title}!{adj_cell.coordinate}"))
        return entries

    def _extract_rate_rows(self, sheet, max_row: int = 150, max_col: int = 20):
        """Extract terminal-growth and WACC rows from a sheet, stopping before sensitivity tables."""
        tg_rows = []
        wacc_rows = []

        for row in sheet.iter_rows(max_row=max_row, max_col=max_col):
            # Stop scanning if we hit a sensitivity section header after collecting assumptions
            row_texts = [c.value.lower() for c in row if isinstance(c.value, str) and c.value.strip()]
            if any(_is_sensitivity_header(t) for t in row_texts):
                if tg_rows or wacc_rows:
                    break
                continue

            for cell in row:
                if not (cell.value and isinstance(cell.value, str)):
                    continue
                if _is_terminal_growth_label(cell.value):
                    entries = self._collect_adjacent_rates(sheet, cell, allow_zero=True)
                    if entries:
                        tg_rows.append((cell.row, entries))
                elif _is_wacc_label(cell.value):
                    entries = self._collect_adjacent_rates(sheet, cell, allow_zero=False)
                    if entries:
                        wacc_rows.append((cell.row, entries))

        return tg_rows, wacc_rows

    def _check_terminal_growth_vs_wacc(self):
        """Critical check: Terminal growth must be less than WACC across all scenarios."""
        try:
            dcf_sheet = self.workbook_values['DCF']
            tg_rows, wacc_rows = self._extract_rate_rows(dcf_sheet)

            # Fallback to WACC sheet if WACC is not explicitly displayed on DCF sheet
            if not wacc_rows and 'WACC' in self.workbook_values.sheetnames:
                _, wacc_rows = self._extract_rate_rows(self.workbook_values['WACC'])

            if not tg_rows or not wacc_rows:
                self.warnings.append("Could not locate terminal growth and WACC values")
                return

            pairs = []
            if len(tg_rows) == len(wacc_rows):
                paired_rows = list(zip(tg_rows, wacc_rows))
            else:
                paired_rows = [
                    (tg_row, min(wacc_rows, key=lambda wr: abs(tg_row[0] - wr[0])))
                    for tg_row in tg_rows
                ]

            for (_, tg_entries), (_, wacc_entries) in paired_rows:
                wacc_by_col = {col: (val, coord) for col, val, coord in wacc_entries}
                matched_any = False
                for col, tg_val, tg_coord in tg_entries:
                    if col in wacc_by_col:
                        w_val, w_coord = wacc_by_col[col]
                        pairs.append((tg_val, tg_coord, w_val, w_coord))
                        matched_any = True
                if not matched_any:
                    w_val, w_coord = wacc_entries[0][1], wacc_entries[0][2]
                    for _, tg_val, tg_coord in tg_entries:
                        pairs.append((tg_val, tg_coord, w_val, w_coord))

            for tg_val, tg_coord, w_val, w_coord in pairs:
                if tg_val >= w_val:
                    self.errors.append(
                        f"CRITICAL: Terminal growth ({tg_val:.2%} at {tg_coord}) >= "
                        f"WACC ({w_val:.2%} at {w_coord}). "
                        "This creates infinite value and is mathematically invalid."
                    )
                elif len(pairs) == 1:
                    self.info.append(f"✓ Terminal growth ({tg_val:.2%}) < WACC ({w_val:.2%})")
                else:
                    self.info.append(
                        f"✓ Terminal growth ({tg_val:.2%} at {tg_coord}) < WACC ({w_val:.2%} at {w_coord})"
                    )

        except KeyError:
            self.warnings.append("DCF sheet not found")
        except Exception as e:
            self.warnings.append(f"Could not validate terminal growth vs WACC: {str(e)}")

    def _check_wacc_range(self):
        """Check if WACC is in reasonable range"""
        try:
            wacc_sheet = (
                self.workbook_values['WACC']
                if 'WACC' in self.workbook_values.sheetnames
                else self.workbook_values['DCF']
            )
            _, wacc_rows = self._extract_rate_rows(wacc_sheet, max_row=100, max_col=20)
            wacc_values = []
            for _, entries in wacc_rows:
                for _, val, _ in entries:
                    if val not in wacc_values:
                        wacc_values.append(val)

            if wacc_values:
                for wacc in wacc_values:
                    if wacc < 0.05 or wacc > 0.20:
                        self.warnings.append(
                            f"WACC ({wacc:.2%}) is outside typical range (5%-20%). Verify calculation."
                        )
                    else:
                        self.info.append(f"✓ WACC ({wacc:.2%}) in reasonable range")
            else:
                self.warnings.append("Could not locate WACC value")

        except KeyError:
            self.warnings.append("WACC and DCF sheets not found")
        except Exception as e:
            self.warnings.append(f"Could not validate WACC range: {str(e)}")

    def _check_terminal_value_proportion(self):
        """Check if terminal value is reasonable proportion of enterprise value"""
        try:
            dcf_sheet = self.workbook_values['DCF']

            terminal_value = None
            enterprise_value = None

            for row in dcf_sheet.iter_rows(max_row=200, max_col=20):
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        cell_str = cell.value.lower()
                        if _is_sensitivity_header(cell_str):
                            continue
                        if 'terminal' in cell_str and 'value' in cell_str and 'pv' in cell_str:
                            for offset in range(1, 5):
                                adjacent = dcf_sheet.cell(cell.row, cell.column + offset).value
                                if _is_number(adjacent) and adjacent > 0:
                                    terminal_value = float(adjacent)
                                    break
                        if 'enterprise' in cell_str and 'value' in cell_str:
                            for offset in range(1, 5):
                                adjacent = dcf_sheet.cell(cell.row, cell.column + offset).value
                                if _is_number(adjacent) and adjacent > 0:
                                    enterprise_value = float(adjacent)
                                    break

            if terminal_value is not None and enterprise_value is not None and enterprise_value > 0:
                proportion = terminal_value / enterprise_value
                if proportion > 0.80:
                    self.warnings.append(
                        f"Terminal value is {proportion:.1%} of EV (typically should be 50-70%). "
                        "Model may be over-reliant on terminal assumptions."
                    )
                elif proportion < 0.40:
                    self.warnings.append(
                        f"Terminal value is {proportion:.1%} of EV (typically should be 50-70%). "
                        "Check if terminal assumptions are too conservative."
                    )
                else:
                    self.info.append(f"✓ Terminal value is {proportion:.1%} of EV")
            else:
                self.warnings.append("Could not locate terminal value and enterprise value")

        except Exception as e:
            self.warnings.append(f"Could not validate terminal value proportion: {str(e)}")


def validate_dcf_model(excel_path: str) -> dict:
    """
    Validate a DCF model Excel file

    Args:
        excel_path: Path to Excel DCF model

    Returns:
        Dict with validation results
    """
    with DCFModelValidator(excel_path) as validator:
        return validator.validate_all()


def main():
    """Command-line interface"""
    if len(sys.argv) < 2:
        print("Usage: python validate_dcf.py <excel_file> [output.json]")
        print("\nValidates DCF model for:")
        print("  - Formula errors (#REF!, #DIV/0!, etc.)")
        print("  - Terminal growth < WACC (critical)")
        print("  - WACC in reasonable range (5-20%)")
        print("  - Terminal value proportion of EV (40-80%)")
        print("\nReturns JSON with errors, warnings, and info")
        print("\nExample: python validate_dcf.py model.xlsx")
        print("Example: python validate_dcf.py model.xlsx results.json")
        sys.exit(1)

    excel_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    try:
        results = validate_dcf_model(excel_file)

        # Print results
        print(json.dumps(results, indent=2))

        # Save to file if requested
        if output_file:
            with open(output_file, 'w') as f:
                json.dump(results, f, indent=2)

        # Exit with error code if validation failed
        sys.exit(0 if results['status'] == 'PASS' else 1)

    except Exception as e:
        error_result = {
            'file': excel_file,
            'status': 'ERROR',
            'error': str(e)
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
