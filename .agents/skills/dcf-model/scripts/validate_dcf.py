#!/usr/bin/env python3
# Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/dcf-model/scripts/validate_dcf.py), Apache-2.0; modified for Google Antigravity.
"""
DCF Model Validation Script
Validates Excel DCF models for formula errors and common DCF mistakes
"""

import json
import re
import sys
from pathlib import Path
from typing import Any, Optional

_REPO_SCRIPTS_DIR = str(Path(__file__).resolve().parents[4] / "scripts")
if _REPO_SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _REPO_SCRIPTS_DIR)

from xlsx_scan import EXCEL_ERRORS, extract_formula_text as _extract_formula_text, scan_workbook_formulas  # noqa: E402

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


def _cache_sheet_rows(ws, max_row: int, max_col: int) -> list[tuple[int, dict[int, tuple[Any, str]]]]:
    """Read a bounded slice of a read_only worksheet into lightweight (row_idx, {col: (val, coord)}) tuples."""
    from openpyxl.utils import get_column_letter

    ws.reset_dimensions()
    rows_out: list[tuple[int, dict[int, tuple[Any, str]]]] = []
    for row_idx, row in enumerate(ws.iter_rows(max_row=max_row, max_col=max_col), start=1):
        row_map: dict[int, tuple[Any, str]] = {}
        for col_idx, cell in enumerate(row, start=1):
            val = cell.value
            if val is not None:
                coord = getattr(cell, 'coordinate', None) or f"{get_column_letter(col_idx)}{row_idx}"
                row_map[col_idx] = (val, coord)
        if row_map:
            rows_out.append((row_idx, row_map))
    return rows_out


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

        # Single lightweight read_only=True pass for sheet structure & DCF logic checks.
        wb_values = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
        try:
            self.sheet_names: list[str] = [ws.title for ws in wb_values.worksheets]
            ws_by_title = {ws.title: ws for ws in wb_values.worksheets}
            self._dcf_rows: Optional[list[tuple[int, dict[int, tuple[Any, str]]]]] = (
                _cache_sheet_rows(ws_by_title['DCF'], max_row=250, max_col=25)
                if 'DCF' in ws_by_title
                else None
            )
            self._wacc_rows: Optional[list[tuple[int, dict[int, tuple[Any, str]]]]] = (
                _cache_sheet_rows(ws_by_title['WACC'], max_row=150, max_col=20)
                if 'WACC' in ws_by_title
                else None
            )
        finally:
            wb_values.close()

        self.errors: list[str] = []
        self.warnings: list[str] = []
        self.info: list[str] = []

    def close(self) -> None:
        """No-op cleanup retained for context-manager backwards compatibility."""

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

        for sheet in required_sheets:
            if sheet not in self.sheet_names:
                self.errors.append(f"Required sheet missing: {sheet}")
            else:
                self.info.append(f"Found sheet: {sheet}")

        has_sensitivity_sheet = any('sensitivity' in name.lower() for name in self.sheet_names)
        has_inline_sensitivity = False
        if not has_sensitivity_sheet and self._dcf_rows is not None:
            for _, row_map in self._dcf_rows:
                for val, _ in row_map.values():
                    if isinstance(val, str) and _is_sensitivity_header(val):
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
        """Check for Excel formula errors and uncached formulas in all worksheets using read_only=True."""
        scan = scan_workbook_formulas(Path(self.excel_path))
        error_details = {err: [] for err in EXCEL_ERRORS}
        total_errors = scan['total_errors']
        total_formulas = scan['total_formulas']
        uncached_formulas = scan['uncached_formulas']

        for err_tok, bucket in scan['error_summary'].items():
            locs = bucket['locations']
            error_details.setdefault(err_tok, []).extend(locs)
            for location in locs:
                self.errors.append(f"{err_tok} at {location}")

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

    @staticmethod
    def _collect_adjacent_rates(
        sheet_title: str,
        row_map: dict[int, tuple[Any, str]],
        label_col: int,
        allow_zero: bool = False,
    ):
        """Collect all rate values (0 < v < 1, or 0 <= v < 1) adjacent to a label cell."""
        entries = []
        for offset in range(1, 6):
            adj = row_map.get(label_col + offset)
            if adj is None:
                continue
            val, coord = adj
            if _is_number(val):
                in_range = (0 <= val < 1) if allow_zero else (0 < val < 1)
                if in_range:
                    entries.append((label_col + offset, float(val), f"{sheet_title}!{coord}"))
        return entries

    def _extract_rate_rows(
        self,
        sheet_title: str,
        cached_rows: list[tuple[int, dict[int, tuple[Any, str]]]],
        max_row: int = 150,
        max_col: int = 20,
    ):
        """Extract terminal-growth and WACC rows from a cached sheet grid, stopping before sensitivity tables."""
        tg_rows = []
        wacc_rows = []

        for row_idx, row_map in cached_rows:
            if row_idx > max_row:
                break
            row_texts = [
                val.lower()
                for col_idx, (val, _) in row_map.items()
                if col_idx <= max_col and isinstance(val, str) and val.strip()
            ]
            if any(_is_sensitivity_header(t) for t in row_texts):
                if tg_rows or wacc_rows:
                    break
                continue

            for col_idx in sorted(row_map):
                if col_idx > max_col:
                    break
                val, _ = row_map[col_idx]
                if not (val and isinstance(val, str)):
                    continue
                if _is_terminal_growth_label(val):
                    entries = self._collect_adjacent_rates(sheet_title, row_map, col_idx, allow_zero=True)
                    if entries:
                        tg_rows.append((row_idx, entries))
                elif _is_wacc_label(val):
                    entries = self._collect_adjacent_rates(sheet_title, row_map, col_idx, allow_zero=False)
                    if entries:
                        wacc_rows.append((row_idx, entries))

        return tg_rows, wacc_rows

    def _check_terminal_growth_vs_wacc(self):
        """Critical check: Terminal growth must be less than WACC across all scenarios."""
        try:
            if self._dcf_rows is None:
                # Missing DCF sheet is already recorded as an error in check_sheet_structure()
                return
            tg_rows, wacc_rows = self._extract_rate_rows('DCF', self._dcf_rows)

            # Fallback to WACC sheet if WACC is not explicitly displayed on DCF sheet
            if not wacc_rows and self._wacc_rows is not None:
                _, wacc_rows = self._extract_rate_rows('WACC', self._wacc_rows)

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
            sheets_to_check = [
                (title, rows)
                for title, rows in (('WACC', self._wacc_rows), ('DCF', self._dcf_rows))
                if rows is not None
            ]
            if not sheets_to_check:
                # Missing sheets already recorded as errors in check_sheet_structure()
                return

            seen_rounded: set[float] = set()
            wacc_entries: list[tuple[float, str]] = []
            for title, rows in sheets_to_check:
                _, wacc_rows = self._extract_rate_rows(title, rows, max_row=150, max_col=20)
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
        """Check if terminal value is within the institutional 50%-70% of EV target range."""
        try:
            if self._dcf_rows is None:
                # Missing DCF sheet already recorded as error in check_sheet_structure()
                return

            terminal_value = None
            enterprise_value = None

            for row_idx, row_map in self._dcf_rows:
                if row_idx > 200:
                    break
                row_texts = [
                    val.lower()
                    for col_idx, (val, _) in row_map.items()
                    if col_idx <= 20 and isinstance(val, str) and val.strip()
                ]
                if any(_is_sensitivity_header(t) for t in row_texts):
                    if terminal_value is not None or enterprise_value is not None:
                        break
                    continue

                for col_idx in sorted(row_map):
                    if col_idx > 20:
                        break
                    val, _ = row_map[col_idx]
                    if val and isinstance(val, str):
                        cell_str = val.lower()
                        if _is_sensitivity_header(cell_str) or _is_ratio_or_multiple_label(cell_str):
                            continue
                        if (
                            terminal_value is None
                            and 'terminal' in cell_str
                            and 'value' in cell_str
                            and ('pv' in cell_str or 'present value' in cell_str)
                        ):
                            for offset in range(1, 6):
                                adj = row_map.get(col_idx + offset)
                                if adj is not None and _is_number(adj[0]) and adj[0] > 0:
                                    terminal_value = float(adj[0])
                                    break
                        elif (
                            enterprise_value is None
                            and 'enterprise' in cell_str
                            and 'value' in cell_str
                            and 'terminal' not in cell_str
                        ):
                            for offset in range(1, 6):
                                adj = row_map.get(col_idx + offset)
                                if adj is not None and _is_number(adj[0]) and adj[0] > 0:
                                    enterprise_value = float(adj[0])
                                    break

                if terminal_value is not None and enterprise_value is not None:
                    break

            if terminal_value is not None and enterprise_value is not None and enterprise_value > 0:
                proportion = terminal_value / enterprise_value
                if proportion > 0.70:
                    self.warnings.append(
                        f"Terminal value is {proportion:.1%} of EV (typically should be 50-70%). "
                        "Model may be over-reliant on terminal assumptions."
                    )
                elif proportion < 0.50:
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
        print("  - Terminal value proportion of EV (50-70%)")
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
