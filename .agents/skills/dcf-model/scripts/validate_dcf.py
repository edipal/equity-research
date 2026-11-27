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

_MISSING_XLSX_SCAN_MESSAGE = (
    "Shared dependency 'scripts/xlsx_scan.py' could not be found. "
    "Run validate_dcf.py from within the equity-research repository or add the "
    "repository's scripts/ directory to PYTHONPATH."
)


def _find_repo_scripts_dir() -> Optional[Path]:
    """Search ancestor directories for a shared `scripts/xlsx_scan.py` module."""
    for parent in Path(__file__).resolve().parents:
        candidate = parent / "scripts"
        if (candidate / "xlsx_scan.py").is_file():
            return candidate
    return None


_REPO_SCRIPTS_DIR = _find_repo_scripts_dir()
if _REPO_SCRIPTS_DIR is not None:
    _repo_scripts_str = str(_REPO_SCRIPTS_DIR)
    if _repo_scripts_str not in sys.path:
        sys.path.insert(0, _repo_scripts_str)

_XLSX_SCAN_IMPORT_ERROR: Optional[str] = None
try:
    from xlsx_scan import scan_workbook_formulas  # noqa: E402
except ModuleNotFoundError:
    scan_workbook_formulas = None  # type: ignore[assignment]
    _XLSX_SCAN_IMPORT_ERROR = _MISSING_XLSX_SCAN_MESSAGE


def _missing_xlsx_scan_result(excel_path: str) -> dict:
    """Build a structured ERROR validation payload when scripts/xlsx_scan.py is unavailable."""
    from datetime import datetime

    msg = _XLSX_SCAN_IMPORT_ERROR or _MISSING_XLSX_SCAN_MESSAGE
    return {
        'file': excel_path,
        'validation_date': datetime.now().isoformat(),
        'status': 'ERROR',
        'error': msg,
        'error_count': 1,
        'warning_count': 0,
        'errors': [msg],
        'warnings': [],
        'info': [],
    }

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
    'upside wacc',
    'upside case wacc',
    'downside wacc',
    'downside case wacc',
    'worst case wacc',
    'best case wacc',
    'calculated wacc',
    'active wacc',
    'effective wacc',
    'blended wacc',
    'implied wacc',
    'model wacc',
})
RATIO_ONLY_TOKENS = (
    '%',
    '/',
    'percent',
    'proportion',
    'of enterprise value',
    'per share',
)
MULTIPLE_ONLY_TOKENS = (
    'multiple',
    'ebitda',
    'revenue',
    'sales',
)
PV_TV_EXPLICIT_RE = re.compile(
    r'\b('
    r'(?:pv|present\s+value)\s+(?:of\s+)?(?:terminal(?:\s+value)?|tv|exit\s+multiple\s+tv|perpetuity\s+tv)'
    r'|discounted\s+(?:terminal(?:\s+value)?|tv)'
    r')\b'
)
EV_SHORT_RE = re.compile(r'\b(?:implied|total|dcf)\s+ev\b')
SCENARIO_TAG_RE = re.compile(r'\b(bear|base|bull|upside|downside|worst|best)\b')
NON_SCENARIO_PHRASE_RE = re.compile(
    r'\b('
    r'base\s+(?:year|period|rate|revenue|sales|margin|value|price|cost|costs|fee|salary|ebitda|fcf|earnings|eps|growth|line|number|amount|figure|metric)'
    r'|(?:vs\.?|versus|from|of|to|above|below|over|relative\s+to|installed|customer|user|client|subscriber|revenue|cost|asset|tax|share|capital|deposit|earnings)\s+base'
    r'|(?:use|business|edge|test|legal|court|special)\s+case'
    r'|best\s+(?:practice|practices|estimate|estimates|in[\s-]class|effort|efforts)'
    r'|(?:case|scenario)\s+(?:selector|switch|toggle|picker|dropdown|input|driver)'
    r')\b'
)
SCENARIO_HEADER_PREFIX_RE = re.compile(
    r'^(?:'
    r'(?:section|part|tab|block|schedule)\s+[a-z0-9]+(?:\.[a-z0-9]+)?'
    r'|(?:scenario|case)(?:\s+[a-z0-9]+)?'
    r'|[0-9]+[a-z]?(?:\.[0-9]+[a-z]?)?'
    r'|[a-z]'
    r')\s*[:.\-\u2013\u2014)]\s*'
)
_SCENARIO_HEADER_SUFFIX = (
    r'(?:\s+(?:assumptions?|inputs?|drivers?|projections?|model|summary|valuation|'
    r'overview|analysis|estimates?|forecasts?|parameters?|outputs?|block|section))?'
)
SCENARIO_HEADER_RE = re.compile(
    r'^(bear|base|bull|upside|downside|worst|best)'
    r'(?:'
    r'[\s-]+(?:case|scenario)(?:\s+scenario)?'
    + _SCENARIO_HEADER_SUFFIX
    + r'(?:\s*[:\-\u2013\u2014]\s*[^|]{1,40})?'
    r'|\s+(?:assumptions?|inputs?|drivers?|projections?|model|summary|valuation|overview|estimates?|forecasts?|parameters?|outputs?|block|section)'
    r')?$'
)
ACTIVE_SCENARIO_HEADER_RE = re.compile(
    r'^active[\s-]+(?:case|scenario)'
    + _SCENARIO_HEADER_SUFFIX
    + r'(?:\s*[:\-\u2013\u2014]\s*[^|]{1,40})?$'
)
GENERIC_SCENARIO_HEADER_RE = re.compile(
    r'^(?:scenario|case)\s+([0-9]+[a-z]?|[a-z]|i{1,3}|iv|v|vi{0,3})'
    + _SCENARIO_HEADER_SUFFIX
    + r'$'
)

NAMED_WACC_ALIASES = ('dcf_wacc', 'wacc', 'wacc_rate')
NAMED_TG_ALIASES = ('terminalgrowth', 'terminal_growth', 'dcf_terminalgrowth')
NAMED_PV_TV_ALIASES = ('pv_terminalvalue', 'pv_tv', 'dcf_pv_tv')
NAMED_EV_ALIASES = ('enterprisevalue', 'dcf_ev', 'implied_ev')


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


def _detect_scenario_block_header(label: str, row_has_numeric: bool = False) -> Optional[str]:
    """Detect whether a non-rate row label is a genuine scenario block boundary header.

    Returns a canonical scenario tag (``'bear'``, ``'base'``, ``'bull'``, ``'upside'``,
    ``'downside'``, ``'worst'``, ``'best'``, ``'active'``, or ``'scenario_<id>'``) only for
    header/banner rows without unrelated numeric data. Explicitly rejects ordinary metric
    and prose labels such as ``'LTM / Base Year Revenue'``, ``'Revenue CAGR vs. base year'``,
    ``'Business Case'``, ``'Best Practice'``, or ``'Active Scenario Selector'``.
    """
    if row_has_numeric:
        return None
    low = ' '.join(label.lower().strip().split())
    if not low or _is_sensitivity_header(low):
        return None

    core = re.sub(r'\s*\([^)]*\)', '', low).rstrip(':').strip()
    if not core:
        return None

    if NON_SCENARIO_PHRASE_RE.search(core):
        return None

    stripped = core
    for _ in range(2):
        nxt = SCENARIO_HEADER_PREFIX_RE.sub('', stripped, count=1).strip()
        if nxt == stripped or not nxt:
            break
        stripped = nxt

    if NON_SCENARIO_PHRASE_RE.search(stripped):
        return None

    match = SCENARIO_HEADER_RE.match(stripped)
    if match:
        return match.group(1)

    if ACTIVE_SCENARIO_HEADER_RE.match(stripped):
        return 'active'

    generic_match = GENERIC_SCENARIO_HEADER_RE.match(stripped)
    if generic_match:
        return f"scenario_{generic_match.group(1)}"

    return None


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
    low = ' '.join(text.lower().strip().split())
    if any(tok in low for tok in RATIO_ONLY_TOKENS):
        return True
    if OF_EV_RE.search(low):
        return True
    # Explicit Present Value of Terminal Value rows (e.g., "PV of Terminal Value (Exit Multiple)")
    # are dollar valuation amounts, not multiples, once ratio/percentage rows are excluded above.
    if PV_TV_EXPLICIT_RE.search(low):
        return False
    core = re.sub(r'\s*\([^)]*\)', '', low).strip()
    if 'enterprise' in core and 'value' in core:
        # Allow parenthetical method qualifiers such as "Implied Enterprise Value (Exit Multiple Method)"
        # while excluding actual multiple rows like "Enterprise Value Multiple" or "Enterprise Value to EBITDA".
        if any(tok in core for tok in MULTIPLE_ONLY_TOKENS) or EBIT_RE.search(core):
            return True
        return False
    if any(tok in low for tok in MULTIPLE_ONLY_TOKENS):
        return True
    if EBIT_RE.search(low):
        return True
    return False


def _is_pv_tv_label(text: str) -> bool:
    """Match dollar Present Value of Terminal Value labels across common naming conventions."""
    low = ' '.join(text.lower().strip().split())
    if not low or _is_sensitivity_header(low) or _is_ratio_or_multiple_label(low):
        return False
    return bool(PV_TV_EXPLICIT_RE.search(low))


def _is_ev_label(text: str) -> bool:
    """Match implied/DCF Enterprise Value labels while excluding current/market EV and multiples."""
    low = ' '.join(text.lower().strip().split())
    if not low or _is_sensitivity_header(low) or _is_ratio_or_multiple_label(low):
        return False
    if 'terminal' in low or 'current' in low or 'market' in low:
        return False
    return (
        ('enterprise' in low and 'value' in low)
        or bool(EV_SHORT_RE.search(low))
        or low.rstrip(':') == 'ev'
    )


def _load_sheet_grid(ws) -> dict[int, dict[int, tuple[Any, str]]]:
    """Stream the full used range of a read_only worksheet into a sparse row -> col map.

    Only non-empty cells are stored so large sheets remain lightweight without truncating
    rows beyond 250 or columns beyond Z.
    """
    from openpyxl.utils import get_column_letter

    ws.reset_dimensions()
    grid: dict[int, dict[int, tuple[Any, str]]] = {}
    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        row_map: dict[int, tuple[Any, str]] = {}
        for col_idx, cell in enumerate(row, start=1):
            val = cell.value
            if val is not None:
                coord = getattr(cell, 'coordinate', None) or f"{get_column_letter(col_idx)}{row_idx}"
                row_map[col_idx] = (val, coord)
        if row_map:
            grid[row_idx] = row_map
    return grid


def _resolve_named_metric(
    defined_names_lower: dict[str, Any],
    aliases: tuple[str, ...],
    grids_by_sheet: dict[str, Optional[dict[int, dict[int, tuple[Any, str]]]]],
    value_predicate,
) -> Optional[tuple[float, str, str, int, int]]:
    """Resolve the first matching workbook-level named range pointing to a numeric cell on DCF or WACC."""
    from openpyxl.utils.cell import coordinate_to_tuple

    for alias in aliases:
        dn = defined_names_lower.get(alias)
        if dn is None:
            continue
        try:
            destinations = list(dn.destinations)
        except Exception:
            continue
        for sheet_title, ref in destinations:
            grid = grids_by_sheet.get(sheet_title)
            if grid is None or not ref or ':' in ref:
                continue
            ref_clean = ref.replace('$', '').strip()
            try:
                row_idx, col_idx = coordinate_to_tuple(ref_clean)
            except Exception:
                continue
            cell_entry = grid.get(row_idx, {}).get(col_idx)
            if cell_entry is None:
                continue
            val, coord = cell_entry
            if _is_number(val):
                fval = float(val)
                if value_predicate(fval):
                    return (fval, f"{sheet_title}!{coord}", sheet_title, row_idx, col_idx)
    return None


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

        # Single lightweight read_only=True pass over the full used range for DCF & WACC sheets
        # plus workbook-level defined_names resolution.
        wb_values = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
        try:
            self.sheet_names: list[str] = [ws.title for ws in wb_values.worksheets]
            ws_by_title = {ws.title: ws for ws in wb_values.worksheets}
            self._dcf_grid: Optional[dict[int, dict[int, tuple[Any, str]]]] = (
                _load_sheet_grid(ws_by_title['DCF'])
                if 'DCF' in ws_by_title
                else None
            )
            self._wacc_grid: Optional[dict[int, dict[int, tuple[Any, str]]]] = (
                _load_sheet_grid(ws_by_title['WACC'])
                if 'WACC' in ws_by_title
                else None
            )

            defined_names_lower = {
                name.lower(): dn for name, dn in wb_values.defined_names.items()
            }
            grids_by_sheet = {'DCF': self._dcf_grid, 'WACC': self._wacc_grid}
            self._named_metrics: dict[str, tuple[float, str, str, int, int]] = {}
            metric_specs = (
                ('wacc', NAMED_WACC_ALIASES, lambda v: 0 < v < 1),
                ('terminal_growth', NAMED_TG_ALIASES, lambda v: -1 < v < 1),
                ('pv_terminal_value', NAMED_PV_TV_ALIASES, lambda v: v > 0),
                ('enterprise_value', NAMED_EV_ALIASES, lambda v: v > 0),
            )
            for key, aliases, predicate in metric_specs:
                resolved = _resolve_named_metric(
                    defined_names_lower, aliases, grids_by_sheet, predicate
                )
                if resolved is not None:
                    self._named_metrics[key] = resolved
        finally:
            wb_values.close()

        self.errors: list[str] = []
        self.warnings: list[str] = []
        self.info: list[str] = []

    def validate_all(self) -> dict:
        """
        Run all validation checks

        Returns:
            Dict with validation results
        """
        from datetime import datetime

        if scan_workbook_formulas is None:
            result = _missing_xlsx_scan_result(self.excel_path)
            self.errors.append(result['error'])
            return result

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
        if not has_sensitivity_sheet and self._dcf_grid is not None:
            for row_idx in sorted(self._dcf_grid):
                for val, _ in self._dcf_grid[row_idx].values():
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

    def check_formula_errors(self) -> None:
        """Check for Excel formula errors and uncached formulas in all worksheets using read_only=True."""
        assert scan_workbook_formulas is not None
        scan = scan_workbook_formulas(Path(self.excel_path))
        total_errors = scan['total_errors']
        total_formulas = scan['total_formulas']
        uncached_formulas = scan['uncached_formulas']

        for err_tok, bucket in scan['error_summary'].items():
            for location in bucket['locations']:
                self.errors.append(f"{err_tok} at {location}")

        if uncached_formulas > 0:
            self.errors.append(
                f"Uncached formula values detected ({uncached_formulas} of {total_formulas} formula "
                "cell(s) have no cached value). Recalculate workbook via "
                "scripts/recalc_xlsx.py <file.xlsx> --in-place (or --output <recalc.xlsx>) first."
            )

        # Record summary metrics in self.info so self.errors holds only distinct errors
        self.info.append(f"Total formulas: {total_formulas}")
        self.info.append(f"Uncached formulas: {uncached_formulas}")
        if total_errors == 0:
            self.info.append("✓ No formula errors found")
        else:
            self.info.append(f"Total formula errors: {total_errors}")

    def check_dcf_logic(self):
        """Validate DCF-specific logic and calculations"""
        self._check_terminal_growth_vs_wacc()
        self._check_wacc_range()
        self._check_terminal_value_proportion()

    @staticmethod
    def _single_scenario_tag(text: str) -> Optional[str]:
        cleaned = NON_SCENARIO_PHRASE_RE.sub(' ', text.lower())
        tags = set(SCENARIO_TAG_RE.findall(cleaned))
        if len(tags) == 1:
            return next(iter(tags))
        return None

    @staticmethod
    def _collect_adjacent_rates(
        sheet_title: str,
        row_map: dict[int, tuple[Any, str]],
        label_col: int,
        allow_zero: bool = False,
        allow_negative: bool = False,
    ):
        """Collect all rate values adjacent to a label cell.

        Accepts 0 < v < 1 by default (WACC), 0 <= v < 1 when allow_zero=True,
        and -1 < v < 1 when allow_negative=True (e.g., declining terminal growth).
        """
        entries = []
        for offset in range(1, 6):
            adj = row_map.get(label_col + offset)
            if adj is None:
                continue
            val, coord = adj
            if _is_number(val):
                if allow_negative:
                    in_range = -1 < val < 1
                elif allow_zero:
                    in_range = 0 <= val < 1
                else:
                    in_range = 0 < val < 1
                if in_range:
                    entries.append((label_col + offset, float(val), f"{sheet_title}!{coord}"))
        return entries

    def _extract_rate_rows(
        self,
        sheet_title: str,
        sheet_grid: dict[int, dict[int, tuple[Any, str]]],
    ):
        """Extract scenario-scoped terminal-growth and WACC rows from the full sparse sheet grid.

        Tracks scenario block boundaries (blank separator rows, scenario headers such as
        Bear/Base/Bull/Upside/Downside/Case, and explicit per-row scenario tags) so per-scenario
        rate rows never cross scenarios. Stops before inline sensitivity tables.
        """
        tg_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]] = []
        wacc_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]] = []

        block_id = 0
        block_tag: Optional[str] = None
        prev_row_idx: Optional[int] = None

        for row_idx in sorted(sheet_grid):
            row_map = sheet_grid[row_idx]
            if prev_row_idx is not None and row_idx > prev_row_idx + 1:
                block_id += 1
                block_tag = None
            prev_row_idx = row_idx

            row_texts = [
                val.lower()
                for col_idx in sorted(row_map)
                for val, _ in (row_map[col_idx],)
                if isinstance(val, str) and val.strip()
            ]
            row_has_numeric = any(_is_number(val) for val, _ in row_map.values())
            if any(_is_sensitivity_header(t) for t in row_texts):
                if tg_rows or wacc_rows:
                    break
                continue

            row_matches: list[tuple[str, str, list[tuple[int, float, str]]]] = []
            for col_idx in sorted(row_map):
                val, _ = row_map[col_idx]
                if not (val and isinstance(val, str)):
                    continue
                if _is_terminal_growth_label(val):
                    entries = self._collect_adjacent_rates(
                        sheet_title,
                        row_map,
                        col_idx,
                        allow_zero=True,
                        allow_negative=True,
                    )
                    if entries:
                        row_matches.append(('tg', val, entries))
                elif _is_wacc_label(val):
                    entries = self._collect_adjacent_rates(
                        sheet_title,
                        row_map,
                        col_idx,
                        allow_zero=False,
                        allow_negative=False,
                    )
                    if entries:
                        row_matches.append(('wacc', val, entries))

            if not row_matches:
                detected_tags = {
                    tag
                    for text in row_texts
                    if (tag := _detect_scenario_block_header(text, row_has_numeric)) is not None
                }
                if detected_tags:
                    block_id += 1
                    block_tag = next(iter(detected_tags)) if len(detected_tags) == 1 else None
                continue

            for kind, label_text, entries in row_matches:
                explicit_tag = self._single_scenario_tag(label_text)
                if explicit_tag is not None and explicit_tag != block_tag:
                    block_id += 1
                    block_tag = explicit_tag
                effective_tag = explicit_tag or block_tag
                if kind == 'tg':
                    tg_rows.append((row_idx, block_id, effective_tag, entries))
                else:
                    wacc_rows.append((row_idx, block_id, effective_tag, entries))

        return tg_rows, wacc_rows

    @staticmethod
    def _pair_entries_by_column(
        tg_entries: list[tuple[int, float, str]],
        wacc_entries: list[tuple[int, float, str]],
    ) -> list[tuple[float, str, float, str]]:
        """Pair rate entries on matching columns first (horizontal scenario tables), else first WACC."""
        pairs: list[tuple[float, str, float, str]] = []
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
        return pairs

    def _check_terminal_growth_vs_wacc(self):
        """Critical check: Terminal growth must be less than WACC across all scenarios."""
        try:
            if self._dcf_grid is None:
                # Missing DCF sheet is already recorded as an error in check_sheet_structure()
                return

            tg_rows, wacc_rows = self._extract_rate_rows('DCF', self._dcf_grid)
            wacc_sheet_rows = (
                self._extract_rate_rows('WACC', self._wacc_grid)[1]
                if self._wacc_grid is not None
                else []
            )
            named_tg = self._named_metrics.get('terminal_growth')
            named_wacc = self._named_metrics.get('wacc')

            # Collect all discovered WACCs (DCF, WACC sheet, and named range)
            all_wacc_entries: list[tuple[float, str]] = []
            seen_wacc_coords: set[str] = set()
            if named_wacc is not None:
                all_wacc_entries.append((named_wacc[0], named_wacc[1]))
                seen_wacc_coords.add(named_wacc[1])
            for _, _, _, entries in wacc_rows + wacc_sheet_rows:
                for _, w_val, w_coord in entries:
                    if w_coord not in seen_wacc_coords:
                        seen_wacc_coords.add(w_coord)
                        all_wacc_entries.append((w_val, w_coord))

            has_any_tg = bool(tg_rows) or named_tg is not None
            has_any_wacc = bool(all_wacc_entries)
            if not has_any_tg:
                self.errors.append("Could not locate terminal growth rate on DCF sheet")
            if not has_any_wacc:
                self.errors.append("Could not locate WACC value on DCF or WACC sheet")
            if not has_any_tg or not has_any_wacc:
                return

            distinct_wacc_vals = {round(v, 8) for v, _ in all_wacc_entries}
            single_global_wacc: Optional[tuple[float, str]] = None
            if len(distinct_wacc_vals) == 1:
                if named_wacc is not None:
                    single_global_wacc = (named_wacc[0], named_wacc[1])
                elif wacc_rows:
                    first = wacc_rows[0][3][0]
                    single_global_wacc = (first[1], first[2])
                else:
                    single_global_wacc = all_wacc_entries[0]

            min_wacc = min(all_wacc_entries, key=lambda item: item[0])

            pairs: list[tuple[float, str, float, str]] = []
            seen_pairs: set[tuple[str, str]] = set()

            def _add_pair(tg_val: float, tg_coord: str, w_val: float, w_coord: str) -> None:
                key = (tg_coord, w_coord)
                if key not in seen_pairs:
                    seen_pairs.add(key)
                    pairs.append((tg_val, tg_coord, w_val, w_coord))

            for tg_row_idx, tg_block_id, tg_tag, tg_entries in tg_rows:
                same_block_waccs = [wr for wr in wacc_rows if wr[1] == tg_block_id]
                same_tag_waccs = (
                    [wr for wr in wacc_rows if wr[2] == tg_tag]
                    if (not same_block_waccs and tg_tag is not None)
                    else []
                )
                if same_block_waccs:
                    closest = min(same_block_waccs, key=lambda wr: abs(tg_row_idx - wr[0]))
                    for pair in self._pair_entries_by_column(tg_entries, closest[3]):
                        _add_pair(*pair)
                elif same_tag_waccs:
                    closest = min(same_tag_waccs, key=lambda wr: abs(tg_row_idx - wr[0]))
                    for pair in self._pair_entries_by_column(tg_entries, closest[3]):
                        _add_pair(*pair)
                elif single_global_wacc is not None:
                    w_val, w_coord = single_global_wacc
                    for _, tg_val, tg_coord in tg_entries:
                        _add_pair(tg_val, tg_coord, w_val, w_coord)
                else:
                    # Multiple distinct scenario WACCs exist and this TG row has no WACC
                    # in its own scenario block; use the conservative minimum WACC.
                    w_val, w_coord = min_wacc
                    for _, tg_val, tg_coord in tg_entries:
                        _add_pair(tg_val, tg_coord, w_val, w_coord)

            # Ensure named TerminalGrowth is always validated even with non-standard row labels
            if named_tg is not None:
                ntg_val, ntg_coord, _, _, _ = named_tg
                already_checked_tg = {p[1] for p in pairs}
                if ntg_coord not in already_checked_tg:
                    if named_wacc is not None:
                        w_val, w_coord = named_wacc[0], named_wacc[1]
                    elif single_global_wacc is not None:
                        w_val, w_coord = single_global_wacc
                    else:
                        w_val, w_coord = min_wacc
                    _add_pair(ntg_val, ntg_coord, w_val, w_coord)

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
        """Check if WACC is in reasonable range across named ranges and both WACC and DCF sheets."""
        try:
            sheets_to_check = [
                (title, grid)
                for title, grid in (('WACC', self._wacc_grid), ('DCF', self._dcf_grid))
                if grid is not None
            ]
            if not sheets_to_check:
                # Missing sheets already recorded as errors in check_sheet_structure()
                return

            seen_rounded: set[float] = set()
            wacc_entries: list[tuple[float, str]] = []

            named_wacc = self._named_metrics.get('wacc')
            if named_wacc is not None:
                key = round(named_wacc[0], 8)
                seen_rounded.add(key)
                wacc_entries.append((named_wacc[0], named_wacc[1]))

            for title, grid in sheets_to_check:
                _, wacc_rows = self._extract_rate_rows(title, grid)
                for _, _, _, entries in wacc_rows:
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
            if self._dcf_grid is None:
                # Missing DCF sheet already recorded as error in check_sheet_structure()
                return

            named_tv = self._named_metrics.get('pv_terminal_value')
            named_ev = self._named_metrics.get('enterprise_value')
            terminal_value: Optional[float] = named_tv[0] if named_tv is not None else None
            enterprise_value: Optional[float] = named_ev[0] if named_ev is not None else None

            if terminal_value is None or enterprise_value is None:
                for row_idx in sorted(self._dcf_grid):
                    row_map = self._dcf_grid[row_idx]
                    row_texts = [
                        val.lower()
                        for col_idx in sorted(row_map)
                        for val, _ in (row_map[col_idx],)
                        if isinstance(val, str) and val.strip()
                    ]
                    if any(_is_sensitivity_header(t) for t in row_texts):
                        if terminal_value is not None or enterprise_value is not None:
                            break
                        continue

                    for col_idx in sorted(row_map):
                        val, _ = row_map[col_idx]
                        if not (val and isinstance(val, str)):
                            continue
                        if terminal_value is None and _is_pv_tv_label(val):
                            for offset in range(1, 6):
                                adj = row_map.get(col_idx + offset)
                                if adj is not None and _is_number(adj[0]) and adj[0] > 0:
                                    terminal_value = float(adj[0])
                                    break
                        elif enterprise_value is None and _is_ev_label(val):
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
    if scan_workbook_formulas is None:
        return _missing_xlsx_scan_result(excel_path)
    validator = DCFModelValidator(excel_path)
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
            with open(output_file, 'w', encoding='utf-8') as f:
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
