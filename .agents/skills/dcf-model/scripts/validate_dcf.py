#!/usr/bin/env python3
# Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/dcf-model/scripts/validate_dcf.py), Apache-2.0; modified for Google Antigravity.
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
OF_EV_RE = re.compile(r'\bof\s+ev\b')
EBIT_RE = re.compile(r'\bebit\b')
WACC_TOKEN_RE = re.compile(r'\bwacc\b')
WACC_EXCLUDE_WORDS_RE = re.compile(
    r'\b('
    r'weight|weights|weighting|factor|factors|table|tables|sensitivity|'
    r'vs|versus|spread|change|delta|assumption|assumptions|input|inputs|'
    r'target|beta|debt|equity|build|calculation|summary|schedule|note|notes|header|section'
    r')\b'
)
GENUINE_WACC_CORES = frozenset({
    'wacc',
    'weighted average cost of capital',
    'discount rate',
    'base wacc',
    'base case wacc',
    'bear wacc',
    'bear case wacc',
    'bull wacc',
    'bull case wacc',
    'calculated wacc',
    'active wacc',
    'effective wacc',
    'blended wacc',
    'implied wacc',
    'model wacc',
})
RATIO_OR_MULTIPLE_TOKENS = (
    '%',
    '/',
    'percent',
    'proportion',
    'of enterprise value',
    'multiple',
    'ebitda',
    'revenue',
    'sales',
    'per share',
)


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


SENSITIVITY_AXIS_RE = re.compile(
    r'\bwacc\s*(?:[/\\×]|by)\s*(?:terminal\s*)?(?:growth|g)\b'
    r'|\b(?:terminal\s*)?(?:growth|g)\s*(?:[/\\×]|by)\s*wacc\b'
)


def _is_sensitivity_header(text: str) -> bool:
    """Return True if a label belongs to a sensitivity table header rather than a model note or assumption."""
    low = text.lower()
    if 'sensitivity' in low:
        return True
    if SENSITIVITY_AXIS_RE.search(low):
        return True
    if SENSITIVITY_VS_RE.search(low) and any(
        kw in low for kw in ('wacc', 'growth', 'margin', 'beta', 'multiple', 'discount')
    ):
        return True
    return False


def _is_wacc_label(text: str) -> bool:
    """Match genuine WACC value labels while excluding weights, factors, headers, and sensitivity rows."""
    low = ' '.join(text.lower().strip().split())
    if not low or _is_sensitivity_header(low):
        return False

    is_full_phrase = 'weighted average cost of capital' in low
    has_wacc_token = bool(WACC_TOKEN_RE.search(low))
    if not is_full_phrase and not has_wacc_token:
        return False

    # Exclude non-WACC labels such as "Equity weight in WACC", "Debt weight in WACC",
    # "WACC discount factor", "WACC Build", "WACC Spread", etc.
    # Note: word boundary \bweight\b does NOT match "weighted" in "weighted average cost of capital".
    if WACC_EXCLUDE_WORDS_RE.search(low):
        return False

    # Strip parenthetical qualifiers like "(%)", "(WACC)", "(Discount Rate)" and trailing colon
    core = re.sub(r'\s*\([^)]*\)', '', low).rstrip(':').strip()
    if core in GENUINE_WACC_CORES:
        return True
    if is_full_phrase and len(core.split()) <= 7:
        return True
    return False


def _is_terminal_growth_label(text: str) -> bool:
    low = ' '.join(text.lower().strip().split())
    if not low or _is_sensitivity_header(low):
        return False
    if any(kw in low for kw in ('spread', 'table', 'note', 'check')):
        return False
    return 'terminal' in low and 'growth' in low


def _is_ratio_or_multiple_label(text: str) -> bool:
    """Return True if a row label represents a percentage, proportion, or valuation multiple."""
    low = text.lower()
    if any(tok in low for tok in RATIO_OR_MULTIPLE_TOKENS):
        return True
    if OF_EV_RE.search(low) or EBIT_RE.search(low):
        return True
    return False


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

    @staticmethod
    def _get_worksheet(workbook, name: str):
        """Return Worksheet by title, ignoring Chartsheet objects."""
        for ws in workbook.worksheets:
            if ws.title == name:
                return ws
        return None

    def check_sheet_structure(self):
        """Verify required sheets exist (2-sheet DCF + WACC spec with inline or separate sensitivity)."""
        required_sheets = ['DCF', 'WACC']
        sheet_names = [ws.title for ws in self.workbook_values.worksheets]

        for sheet in required_sheets:
            if sheet not in sheet_names:
                self.errors.append(f"Required sheet missing: {sheet}")
            else:
                self.info.append(f"Found sheet: {sheet}")

        has_sensitivity_sheet = any('sensitivity' in name.lower() for name in sheet_names)
        has_inline_sensitivity = False
        if not has_sensitivity_sheet:
            dcf_ws = self._get_worksheet(self.workbook_formulas, 'DCF')
            if dcf_ws is not None:
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
        """Check for Excel formula errors and uncached formulas in all worksheets."""
        error_details = {err: [] for err in EXCEL_ERRORS}
        total_errors = 0
        total_formulas = 0
        uncached_formulas = 0

        ws_values_by_title = {ws.title: ws for ws in self.workbook_values.worksheets}
        for ws_formulas in self.workbook_formulas.worksheets:
            sheet_name = ws_formulas.title
            ws_values = ws_values_by_title.get(sheet_name)
            if ws_values is None:
                continue

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
            dcf_sheet = self._get_worksheet(self.workbook_values, 'DCF')
            if dcf_sheet is None:
                # Missing DCF sheet is already recorded as an error in check_sheet_structure()
                return
            tg_rows, wacc_rows = self._extract_rate_rows(dcf_sheet)

            # Fallback to WACC sheet if WACC is not explicitly displayed on DCF sheet
            if not wacc_rows:
                wacc_sheet = self._get_worksheet(self.workbook_values, 'WACC')
                if wacc_sheet is not None:
                    _, wacc_rows = self._extract_rate_rows(wacc_sheet)

            if not tg_rows:
                self.errors.append("Could not locate terminal growth rate on DCF sheet")
            if not wacc_rows:
                self.errors.append("Could not locate WACC value on DCF or WACC sheet")
            if not tg_rows or not wacc_rows:
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

        except Exception as e:
            self.errors.append(f"Could not validate terminal growth vs WACC: {str(e)}")

    def _check_wacc_range(self):
        """Check if WACC is in reasonable range across both WACC and DCF sheets (deduplicated)."""
        try:
            wacc_sheet = self._get_worksheet(self.workbook_values, 'WACC')
            dcf_sheet = self._get_worksheet(self.workbook_values, 'DCF')
            sheets_to_check = [s for s in (wacc_sheet, dcf_sheet) if s is not None]
            if not sheets_to_check:
                # Missing sheets already recorded as errors in check_sheet_structure()
                return

            seen_rounded: set[float] = set()
            wacc_entries: list[tuple[float, str]] = []
            for sheet in sheets_to_check:
                _, wacc_rows = self._extract_rate_rows(sheet, max_row=150, max_col=20)
                for _, entries in wacc_rows:
                    for _, val, coord in entries:
                        key = round(val, 8)
                        if key not in seen_rounded:
                            seen_rounded.add(key)
                            wacc_entries.append((val, coord))

            if wacc_entries:
                for wacc, coord in wacc_entries:
                    if wacc < 0.05 or wacc > 0.20:
                        self.warnings.append(
                            f"WACC ({wacc:.2%} at {coord}) is outside typical range (5%-20%). Verify calculation."
                        )
                    else:
                        self.info.append(f"✓ WACC ({wacc:.2%}) in reasonable range")
            else:
                self.errors.append("Could not locate WACC value for range check")

        except Exception as e:
            self.errors.append(f"Could not validate WACC range: {str(e)}")

    def _check_terminal_value_proportion(self):
        """Check if terminal value is reasonable proportion of enterprise value"""
        try:
            dcf_sheet = self._get_worksheet(self.workbook_values, 'DCF')
            if dcf_sheet is None:
                # Missing DCF sheet already recorded as error in check_sheet_structure()
                return

            terminal_value = None
            enterprise_value = None

            for row in dcf_sheet.iter_rows(max_row=200, max_col=20):
                row_texts = [c.value.lower() for c in row if isinstance(c.value, str) and c.value.strip()]
                if any(_is_sensitivity_header(t) for t in row_texts):
                    if terminal_value is not None or enterprise_value is not None:
                        break
                    continue

                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        cell_str = cell.value.lower()
                        if _is_sensitivity_header(cell_str) or _is_ratio_or_multiple_label(cell_str):
                            continue
                        if (
                            terminal_value is None
                            and 'terminal' in cell_str
                            and 'value' in cell_str
                            and ('pv' in cell_str or 'present value' in cell_str)
                        ):
                            for offset in range(1, 6):
                                adjacent = dcf_sheet.cell(cell.row, cell.column + offset).value
                                if _is_number(adjacent) and adjacent > 0:
                                    terminal_value = float(adjacent)
                                    break
                        elif (
                            enterprise_value is None
                            and 'enterprise' in cell_str
                            and 'value' in cell_str
                            and 'terminal' not in cell_str
                        ):
                            for offset in range(1, 6):
                                adjacent = dcf_sheet.cell(cell.row, cell.column + offset).value
                                if _is_number(adjacent) and adjacent > 0:
                                    enterprise_value = float(adjacent)
                                    break

                if terminal_value is not None and enterprise_value is not None:
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
                self.errors.append(
                    "Could not locate PV of Terminal Value and Enterprise Value on DCF sheet"
                )

        except Exception as e:
            self.errors.append(f"Could not validate terminal value proportion: {str(e)}")


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
        print("  - Required DCF and WACC sheets")
        print("  - Formula errors (#REF!, #DIV/0!, etc.) and uncached formulas")
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
