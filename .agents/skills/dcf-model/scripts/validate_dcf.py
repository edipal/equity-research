#!/usr/bin/env python3
# Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/dcf-model/scripts/validate_dcf.py), Apache-2.0; modified for Google Antigravity.
"""
DCF Model Validation Script
Validates Excel DCF models for formula errors and common DCF mistakes
"""

import json
import math
import re
import sys
from collections.abc import Iterable
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

SCENARIO_NAMES = ('bear', 'base', 'bull')
NAMED_SCENARIO_TG_ALIASES = {
    'bear': ('terminalgrowth_bear', 'terminal_growth_bear', 'dcf_terminalgrowth_bear'),
    'base': ('terminalgrowth_base', 'terminal_growth_base', 'dcf_terminalgrowth_base'),
    'bull': ('terminalgrowth_bull', 'terminal_growth_bull', 'dcf_terminalgrowth_bull'),
}
NAMED_SCENARIO_WACC_ALIASES = {
    'bear': ('wacc_bear', 'dcf_wacc_bear'),
    'base': ('wacc_base', 'dcf_wacc_base'),
    'bull': ('wacc_bull', 'dcf_wacc_bull'),
}


def _is_number(value) -> bool:
    """Return True for finite numeric values while excluding Python booleans."""
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def _dedupe_by_val(entries: Iterable[tuple[float, str]]) -> list[tuple[float, str]]:
    """Deduplicate (value, coord) pairs by rounded numeric value while preserving order."""
    out: list[tuple[float, str]] = []
    seen: set[float] = set()
    for val, coord in entries:
        rkey = round(val, 8)
        if rkey not in seen:
            seen.add(rkey)
            out.append((val, coord))
    return out


_SENS_AXIS_TERM = (
    r'(?:'
    r'wacc'
    r'|discount\s+rate'
    r'|(?:terminal|perpetuity)\s+(?:growth(?:\s+rate)?|g)'
    r'|(?:exit|terminal)\s+(?:ev\s*/\s*ebitda\s+|ebitda\s+)?multiple'
    r')'
)
_SENS_SEP = r'(?:\s*(?:vs\.?|versus|[/\\×]|by)\s*|\s+x\s+)'
SENSITIVITY_PAIR_RE = re.compile(
    rf'\b{_SENS_AXIS_TERM}\b{_SENS_SEP}\b{_SENS_AXIS_TERM}\b'
    rf'|\b(?:wacc|discount\s+rate)\b{_SENS_SEP}\bg\b'
    rf'|\bg\b{_SENS_SEP}\b(?:wacc|discount\s+rate)\b'
)


def _is_sensitivity_header(text: str) -> bool:
    """Return True only for genuine sensitivity table banners or two-variable valuation axis headers."""
    low = ' '.join(text.lower().strip().split())
    if not low:
        return False
    if 'sensitivity' in low:
        return True
    if SENSITIVITY_PAIR_RE.search(low):
        return True
    return False


def _primary_row_text(row_map: dict[int, tuple[Any, str]]) -> Optional[str]:
    """Return the leftmost non-empty string cell in a sparse row map, or None."""
    for col_idx in sorted(row_map):
        val, _ = row_map[col_idx]
        if isinstance(val, str) and val.strip():
            return val.strip()
    return None


def _is_sensitivity_stop_row(primary_text: Optional[str], row_has_numeric: bool) -> bool:
    """Return True when the row's primary label marks the start of an inline sensitivity table.

    Checks only the primary (leftmost non-empty string) cell so commentary notes in right-hand
    columns (e.g., ``DCF!H39 = "Margin vs. peer median"``) can never trigger an early stop.
    Requires either no numeric data on the row (a section banner) or an explicit two-variable
    sensitivity grid corner header (e.g., ``"WACC vs Terminal Growth"`` or ``"WACC \\ Terminal g"``)
    that may sit on the same row as numeric axis percentages.
    """
    if not primary_text or not _is_sensitivity_header(primary_text):
        return False
    if not row_has_numeric:
        return True
    return bool(SENSITIVITY_PAIR_RE.search(' '.join(primary_text.lower().split())))


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


def _lookup_sheet_grid(
    grids_by_sheet: dict[str, Optional[dict[int, dict[int, tuple[Any, str]]]]],
    sheet_title: str,
) -> tuple[str, Optional[dict[int, dict[int, tuple[Any, str]]]]]:
    """Look up a sheet grid by canonical title, falling back to unescaped/case-insensitive match."""
    grid = grids_by_sheet.get(sheet_title)
    if grid is not None:
        return sheet_title, grid
    unescaped = sheet_title.replace("''", "'")
    targets = {sheet_title.lower(), unescaped.lower()}
    for canon_title, candidate_grid in grids_by_sheet.items():
        if candidate_grid is not None and canon_title.lower() in targets:
            return canon_title, candidate_grid
    return unescaped, None


def _resolve_named_metric(
    defined_names_lower: dict[str, tuple[str, int, int]],
    aliases: tuple[str, ...],
    grids_by_sheet: dict[str, Optional[dict[int, dict[int, tuple[Any, str]]]]],
    value_predicate=None,
) -> Optional[tuple[float, str, str, int, int]]:
    """Resolve the first matching workbook-level named range pointing to a finite numeric cell."""
    for alias in aliases:
        target = defined_names_lower.get(alias)
        if target is None:
            continue
        sheet_title, row_idx, col_idx = target
        canon_title, grid = _lookup_sheet_grid(grids_by_sheet, sheet_title)
        if grid is None:
            continue
        cell_entry = grid.get(row_idx, {}).get(col_idx)
        if cell_entry is None:
            continue
        val, coord = cell_entry
        if _is_number(val):
            fval = float(val)
            if value_predicate is None or value_predicate(fval):
                return (fval, f"{canon_title}!{coord}", canon_title, row_idx, col_idx)
    return None


def _find_invalid_named_metric(
    defined_names_lower: dict[str, tuple[str, int, int]],
    aliases: tuple[str, ...],
    grids_by_sheet: dict[str, Optional[dict[int, dict[int, tuple[Any, str]]]]],
) -> Optional[tuple[Any, str]]:
    """Return (raw_value, 'Sheet!Coord') if an explicitly named cell is empty or non-numeric."""
    for alias in aliases:
        target = defined_names_lower.get(alias)
        if target is None:
            continue
        sheet_title, row_idx, col_idx = target
        canon_title, grid = _lookup_sheet_grid(grids_by_sheet, sheet_title)
        if grid is None:
            continue
        cell_entry = grid.get(row_idx, {}).get(col_idx)
        if cell_entry is None:
            from openpyxl.utils import get_column_letter

            return (None, f"{canon_title}!{get_column_letter(col_idx)}{row_idx}")
        val, coord = cell_entry
        if not _is_number(val):
            return (val, f"{canon_title}!{coord}")
    return None


def _bind_named(
    defined_names_lower: dict[str, tuple[str, int, int]],
    aliases: tuple[str, ...],
    grids_by_sheet: dict[str, Optional[dict[int, dict[int, tuple[Any, str]]]]],
) -> tuple[Optional[tuple[float, str, str, int, int]], Optional[tuple[Any, str]]]:
    """Resolve a named metric to a finite numeric entry or return its invalid diagnostic."""
    resolved = _resolve_named_metric(defined_names_lower, aliases, grids_by_sheet)
    if resolved is not None:
        return resolved, None
    return None, _find_invalid_named_metric(defined_names_lower, aliases, grids_by_sheet)


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

        self._scan: Optional[dict[str, Any]] = None
        self.uncached_formulas: int = 0
        self._formula_error_coords: set[str] = set()
        self.sheet_names: list[str] = []
        self._dcf_grid: Optional[dict[int, dict[int, tuple[Any, str]]]] = None
        self._wacc_grid: Optional[dict[int, dict[int, tuple[Any, str]]]] = None
        self._named_metrics: dict[str, tuple[float, str, str, int, int]] = {}
        self._named_scenario_tg: dict[str, tuple[float, str, str, int, int]] = {}
        self._named_scenario_wacc: dict[str, tuple[float, str, str, int, int]] = {}
        self._invalid_named_metrics: dict[str, tuple[Any, str]] = {}
        self._invalid_named_scenario_tg: dict[str, tuple[Any, str]] = {}
        self._invalid_named_scenario_wacc: dict[str, tuple[Any, str]] = {}

        if scan_workbook_formulas is not None:
            # Single lockstep read_only=True pass collecting formulas, errors, sheet grids, and defined names
            self._scan = scan_workbook_formulas(
                path_obj,
                capture_sheet_grids=True,
                grid_sheet_names=("DCF", "WACC"),
                capture_defined_names=True,
            )
            self.uncached_formulas = int(self._scan.get('uncached_formulas', 0))
            for bucket in self._scan.get('error_summary', {}).values():
                self._formula_error_coords.update(bucket.get('locations', ()))
            self.sheet_names = list(self._scan.get('sheetnames', []))
            sheet_grids: dict[str, dict[int, dict[int, tuple[Any, str]]]] = self._scan.get(
                'sheet_grids', {}
            )
            self._dcf_grid = sheet_grids.get('DCF')
            self._wacc_grid = sheet_grids.get('WACC')

            raw_defined_names: dict[str, tuple[str, int, int]] = self._scan.get('defined_names', {})
            defined_names_lower = {name.lower(): target for name, target in raw_defined_names.items()}
            grids_by_sheet: dict[str, Optional[dict[int, dict[int, tuple[Any, str]]]]] = dict(sheet_grids)

            # Accept any finite numeric value for explicitly bound workbook named ranges so
            # out-of-bounds mistakes (e.g., typing 2.5 for 2.5% or negative WACC/TV) are
            # caught explicitly rather than silently dropped.
            metric_specs = (
                ('wacc', NAMED_WACC_ALIASES),
                ('terminal_growth', NAMED_TG_ALIASES),
                ('pv_terminal_value', NAMED_PV_TV_ALIASES),
                ('enterprise_value', NAMED_EV_ALIASES),
            )
            def _store_named(
                valid_store: dict[str, tuple[float, str, str, int, int]],
                invalid_store: dict[str, tuple[Any, str]],
                key: str,
                aliases: tuple[str, ...],
            ) -> None:
                resolved, invalid = _bind_named(defined_names_lower, aliases, grids_by_sheet)
                if resolved is not None:
                    valid_store[key] = resolved
                elif invalid is not None:
                    invalid_store[key] = invalid

            for key, aliases in metric_specs:
                _store_named(self._named_metrics, self._invalid_named_metrics, key, aliases)

            for scenario in SCENARIO_NAMES:
                _store_named(
                    self._named_scenario_tg,
                    self._invalid_named_scenario_tg,
                    scenario,
                    NAMED_SCENARIO_TG_ALIASES[scenario],
                )
                _store_named(
                    self._named_scenario_wacc,
                    self._invalid_named_scenario_wacc,
                    scenario,
                    NAMED_SCENARIO_WACC_ALIASES[scenario],
                )

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
                row_map = self._dcf_grid[row_idx]
                primary_text = _primary_row_text(row_map)
                row_has_numeric = any(_is_number(val) for val, _ in row_map.values())
                if _is_sensitivity_stop_row(primary_text, row_has_numeric):
                    has_inline_sensitivity = True
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
        """Check for Excel formula errors and uncached formulas using the single-pass workbook scan."""
        assert scan_workbook_formulas is not None
        scan = self._scan if self._scan is not None else scan_workbook_formulas(Path(self.excel_path))
        total_errors = scan['total_errors']
        total_formulas = scan['total_formulas']
        uncached_formulas = scan['uncached_formulas']
        self.uncached_formulas = int(uncached_formulas)

        for err_tok, bucket in scan['error_summary'].items():
            for location in bucket['locations']:
                self._formula_error_coords.add(location)
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

    def _all_scenario_tg_named(self) -> bool:
        return all(
            s in self._named_scenario_tg or s in self._invalid_named_scenario_tg
            for s in SCENARIO_NAMES
        )

    def _all_scenario_wacc_named(self) -> bool:
        return all(
            s in self._named_scenario_wacc or s in self._invalid_named_scenario_wacc
            for s in SCENARIO_NAMES
        )

    def _has_named_wacc_coverage(self) -> bool:
        return (
            'wacc' in self._named_metrics
            or 'wacc' in self._invalid_named_metrics
            or self._all_scenario_wacc_named()
        )

    def _report_invalid_named(
        self,
        items: list[tuple[str, Any, str]],
        *,
        severity: str,
        empty_template: str,
        non_numeric_template: str,
        skip_empty_when_uncached: bool = True,
    ) -> None:
        """Emit deduplicated error/warning messages for empty or non-numeric named cells."""
        target_list = self.errors if severity == 'error' else self.warnings
        seen_coords: set[str] = set()
        for label, raw_val, coord in items:
            if coord in seen_coords:
                continue
            seen_coords.add(coord)
            if coord in self._formula_error_coords:
                continue
            if raw_val is None:
                if not skip_empty_when_uncached or self.uncached_formulas == 0:
                    target_list.append(empty_template.format(label=label, coord=coord))
            else:
                target_list.append(
                    non_numeric_template.format(label=label, coord=coord, raw_val=raw_val)
                )

    def _owned_named_coords(
        self,
        metric_key: str,
        scenario_valid: dict[str, tuple[float, str, str, int, int]],
        scenario_invalid: dict[str, tuple[Any, str]],
    ) -> set[str]:
        """Return all canonical coordinates owned by valid or invalid named ranges for a metric."""
        coords: set[str] = set()
        if metric_key in self._named_metrics:
            coords.add(self._named_metrics[metric_key][1])
        if metric_key in self._invalid_named_metrics:
            coords.add(self._invalid_named_metrics[metric_key][1])
        for scenario in SCENARIO_NAMES:
            if scenario in scenario_valid:
                coords.add(scenario_valid[scenario][1])
            if scenario in scenario_invalid:
                coords.add(scenario_invalid[scenario][1])
        return coords

    @staticmethod
    def _filter_rows_excluding_coords(
        rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]],
        excluded_coords: set[str],
    ) -> list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]]:
        """Filter extracted rate rows to drop any cell coordinates owned by excluded_coords."""
        if not excluded_coords:
            return rows
        filtered: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]] = []
        for row_idx, block_id, tag, entries in rows:
            kept = [entry for entry in entries if entry[2] not in excluded_coords]
            if kept:
                filtered.append((row_idx, block_id, tag, kept))
        return filtered

    def check_dcf_logic(self):
        """Validate DCF-specific logic by always merging workbook named ranges with label-scanned rates."""
        dcf_tg_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]] = []
        dcf_wacc_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]] = []
        wacc_sheet_wacc_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]] = []

        if self._dcf_grid is not None:
            dcf_tg_rows, dcf_wacc_rows = self._extract_rate_rows('DCF', self._dcf_grid)
        if self._wacc_grid is not None:
            _, wacc_sheet_wacc_rows = self._extract_rate_rows('WACC', self._wacc_grid)

        self._check_terminal_growth_vs_wacc(
            dcf_tg_rows=dcf_tg_rows,
            dcf_wacc_rows=dcf_wacc_rows,
            wacc_sheet_wacc_rows=wacc_sheet_wacc_rows,
        )
        self._check_wacc_range(
            dcf_wacc_rows=dcf_wacc_rows,
            wacc_sheet_wacc_rows=wacc_sheet_wacc_rows,
        )
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
            primary_text = _primary_row_text(row_map)
            if _is_sensitivity_stop_row(primary_text, row_has_numeric):
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

    def _build_wacc_pools(
        self,
        unnamed_dcf_wacc_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]],
        unnamed_wacc_sheet_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]],
    ) -> tuple[
        list[tuple[float, str]],
        dict[str, list[tuple[float, str]]],
        dict[str, list[tuple[float, str]]],
        list[tuple[float, str]],
    ]:
        """Build deduplicated all, scenario, non-DCF-scenario, and global WACC pools."""
        named_wacc = self._named_metrics.get('wacc')
        all_waccs: list[tuple[float, str]] = []
        seen_coords: set[str] = set()

        def _record_all(val: float, coord: str) -> None:
            if coord not in seen_coords:
                seen_coords.add(coord)
                all_waccs.append((val, coord))

        raw_scenario: dict[str, list[tuple[float, str]]] = {}
        non_dcf_scenario: dict[str, list[tuple[float, str]]] = {}
        raw_global: list[tuple[float, str]] = []

        if named_wacc is not None:
            _record_all(named_wacc[0], named_wacc[1])
            raw_global.append((named_wacc[0], named_wacc[1]))

        for scenario in SCENARIO_NAMES:
            sw = self._named_scenario_wacc.get(scenario)
            if sw is not None:
                _record_all(sw[0], sw[1])
                raw_scenario.setdefault(scenario, []).append((sw[0], sw[1]))
                non_dcf_scenario.setdefault(scenario, []).append((sw[0], sw[1]))

        for is_wacc_sheet, rows in ((False, unnamed_dcf_wacc_rows), (True, unnamed_wacc_sheet_rows)):
            for _, _, tag, entries in rows:
                for _, w_val, w_coord in entries:
                    _record_all(w_val, w_coord)
                    if tag is None or tag == 'active':
                        raw_global.append((w_val, w_coord))
                    else:
                        raw_scenario.setdefault(tag, []).append((w_val, w_coord))
                        if is_wacc_sheet:
                            non_dcf_scenario.setdefault(tag, []).append((w_val, w_coord))

        scenario_waccs = {tag: _dedupe_by_val(items) for tag, items in raw_scenario.items()}
        global_waccs = _dedupe_by_val(raw_global)
        return all_waccs, scenario_waccs, non_dcf_scenario, global_waccs

    def _check_terminal_growth_vs_wacc(
        self,
        dcf_tg_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]],
        dcf_wacc_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]],
        wacc_sheet_wacc_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]],
    ):
        """Critical check: Terminal growth must be less than WACC across all scenarios."""
        try:
            if self._dcf_grid is None:
                # Missing DCF sheet is already recorded as an error in check_sheet_structure()
                return

            all_scenario_tg_named = self._all_scenario_tg_named()
            all_scenario_wacc_named = self._all_scenario_wacc_named()
            has_full_named_coverage = all_scenario_tg_named and self._has_named_wacc_coverage()

            owned_tg_coords = self._owned_named_coords(
                'terminal_growth', self._named_scenario_tg, self._invalid_named_scenario_tg
            )
            owned_wacc_coords = self._owned_named_coords(
                'wacc', self._named_scenario_wacc, self._invalid_named_scenario_wacc
            )
            invalid_wacc_coords: set[str] = {
                coord for _, coord in self._invalid_named_scenario_wacc.values()
            }
            if 'wacc' in self._invalid_named_metrics:
                invalid_wacc_coords.add(self._invalid_named_metrics['wacc'][1])

            # Named ranges win on exact coordinates (including invalid/empty/non-numeric named coords)
            unnamed_tg_rows = self._filter_rows_excluding_coords(dcf_tg_rows, owned_tg_coords)
            unnamed_dcf_wacc_rows = self._filter_rows_excluding_coords(dcf_wacc_rows, owned_wacc_coords)
            unnamed_wacc_sheet_rows = self._filter_rows_excluding_coords(
                wacc_sheet_wacc_rows, owned_wacc_coords
            )
            # Keep valid named DCF WACCs available for same-block/same-tag 1-to-1 column pairing
            valid_dcf_wacc_rows = self._filter_rows_excluding_coords(dcf_wacc_rows, invalid_wacc_coords)

            all_waccs, scenario_waccs, non_dcf_scenario_waccs, global_waccs = self._build_wacc_pools(
                unnamed_dcf_wacc_rows, unnamed_wacc_sheet_rows
            )
            named_tg = self._named_metrics.get('terminal_growth')
            named_wacc = self._named_metrics.get('wacc')

            # Report any explicitly named terminal growth cells that are empty or non-numeric
            invalid_tg_items: list[tuple[str, Any, str]] = [
                (scenario.title(), *self._invalid_named_scenario_tg[scenario])
                for scenario in SCENARIO_NAMES
                if scenario in self._invalid_named_scenario_tg
            ]
            if 'terminal_growth' in self._invalid_named_metrics:
                invalid_tg_items.append(('Active', *self._invalid_named_metrics['terminal_growth']))
            self._report_invalid_named(
                invalid_tg_items,
                severity='error',
                empty_template=(
                    "CRITICAL: Named terminal growth ({label} at {coord}) is empty — "
                    "enter a valid terminal growth rate."
                ),
                non_numeric_template=(
                    "CRITICAL: Named terminal growth ({label} at {coord}) has non-numeric value ({raw_val!r})."
                ),
                skip_empty_when_uncached=True,
            )

            has_any_tg = bool(self._named_scenario_tg) or named_tg is not None or bool(unnamed_tg_rows)
            has_any_wacc = bool(all_waccs)
            has_invalid_named_tg = bool(invalid_tg_items)
            has_invalid_named_wacc = (
                'wacc' in self._invalid_named_metrics or bool(self._invalid_named_scenario_wacc)
            )
            if not has_any_tg and not has_invalid_named_tg:
                self.errors.append("Could not locate terminal growth rate on DCF sheet")
            if not has_any_wacc and not has_invalid_named_wacc:
                self.errors.append("Could not locate WACC value on DCF or WACC sheet")
            if not has_any_tg or not has_any_wacc:
                return

            min_wacc = min(all_waccs, key=lambda item: item[0])

            def _resolve_waccs(
                tag: Optional[str], *, allow_unscoped_global: bool = True
            ) -> list[tuple[float, str]]:
                if tag is not None and scenario_waccs.get(tag):
                    return scenario_waccs[tag]
                if global_waccs and (
                    allow_unscoped_global or named_wacc is not None or not scenario_waccs
                ):
                    return global_waccs
                return [min_wacc]

            pairs: list[tuple[float, str, float, str]] = []
            seen_pairs: set[tuple[str, str]] = set()

            def _add_pair(tg_val: float, tg_coord: str, w_val: float, w_coord: str) -> None:
                key = (tg_coord, w_coord)
                if key not in seen_pairs:
                    seen_pairs.add(key)
                    pairs.append((tg_val, tg_coord, w_val, w_coord))

            # 1. Named TG cells (per-scenario TerminalGrowth_<Scenario> and active TerminalGrowth)
            for scenario in SCENARIO_NAMES:
                stg = self._named_scenario_tg.get(scenario)
                if stg is None or scenario in self._invalid_named_scenario_wacc:
                    continue
                for w_val, w_coord in _resolve_waccs(scenario):
                    _add_pair(stg[0], stg[1], w_val, w_coord)

            if (
                named_tg is not None
                and 'wacc' not in self._invalid_named_metrics
                and (bool(global_waccs) or not (all_scenario_tg_named and all_scenario_wacc_named))
                and named_tg[1] not in {p[1] for p in pairs}
            ):
                for w_val, w_coord in _resolve_waccs(None):
                    _add_pair(named_tg[0], named_tg[1], w_val, w_coord)

            if has_full_named_coverage:
                if not has_invalid_named_tg and not has_invalid_named_wacc:
                    self.info.append(
                        "✓ Validated scenario Terminal Growth and WACC via workbook named ranges"
                    )
            else:
                self.info.append(
                    "Using label-scan fallback for rate validation "
                    "(define DCF_WACC and TerminalGrowth_Bear/Base/Bull named ranges for exact cell binding)"
                )

            # 2. Unnamed DCF TG rows (1-to-1 column alignment with local DCF WACC row, else _resolve_waccs)
            for tg_row_idx, tg_block_id, tg_tag, tg_entries in unnamed_tg_rows:
                if tg_tag in self._invalid_named_scenario_wacc:
                    continue
                local_candidates = [wr for wr in valid_dcf_wacc_rows if wr[1] == tg_block_id]
                if not local_candidates and tg_tag is not None:
                    local_candidates = [wr for wr in valid_dcf_wacc_rows if wr[2] == tg_tag]
                if local_candidates:
                    local_dcf_wacc = min(local_candidates, key=lambda wr: abs(tg_row_idx - wr[0]))
                    for pair in self._pair_entries_by_column(tg_entries, local_dcf_wacc[3]):
                        _add_pair(*pair)
                    local_coords = {w_coord for _, _, w_coord in local_dcf_wacc[3]}
                    target_waccs = (
                        _dedupe_by_val(
                            (w_val, w_coord)
                            for w_val, w_coord in non_dcf_scenario_waccs.get(tg_tag, ())
                            if w_coord not in local_coords
                        )
                        if tg_tag is not None
                        else []
                    )
                else:
                    target_waccs = _resolve_waccs(tg_tag, allow_unscoped_global=False)
                for _, tg_val, tg_coord in tg_entries:
                    for w_val, w_coord in target_waccs:
                        _add_pair(tg_val, tg_coord, w_val, w_coord)

            self._emit_tg_vs_wacc_results(pairs)

        except Exception as e:
            self.errors.append(f"Could not validate terminal growth vs WACC: {str(e)}")

    def _emit_tg_vs_wacc_results(self, pairs: list[tuple[float, str, float, str]]) -> None:
        """Record errors and pass checkmarks for all resolved (terminal_growth, WACC) pairs."""
        for tg_val, tg_coord, w_val, w_coord in pairs:
            if tg_val >= w_val:
                self.errors.append(
                    f"CRITICAL: Terminal growth ({tg_val:.2%} at {tg_coord}) >= "
                    f"WACC ({w_val:.2%} at {w_coord}). "
                    "This creates infinite value and is mathematically invalid."
                )
            elif tg_val <= -1.0:
                self.errors.append(
                    f"CRITICAL: Terminal growth ({tg_val:.2%} at {tg_coord}) is <= -100% — "
                    "verify decimal entry (e.g., enter -0.02 for -2%)."
                )
            elif w_val <= 0 or w_val >= 1.0:
                # Non-positive or >= 100% WACC is flagged as CRITICAL in _check_wacc_range();
                # avoid emitting a misleading pass checkmark against an invalid discount rate.
                continue
            elif len(pairs) == 1:
                self.info.append(f"✓ Terminal growth ({tg_val:.2%}) < WACC ({w_val:.2%})")
            else:
                self.info.append(
                    f"✓ Terminal growth ({tg_val:.2%} at {tg_coord}) < WACC ({w_val:.2%} at {w_coord})"
                )

    def _check_wacc_range(
        self,
        dcf_wacc_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]],
        wacc_sheet_wacc_rows: list[tuple[int, int, Optional[str], list[tuple[int, float, str]]]],
    ):
        """Check if WACC is in reasonable range across named ranges and both WACC and DCF sheets."""
        try:
            if self._wacc_grid is None and self._dcf_grid is None:
                # Missing sheets already recorded as errors in check_sheet_structure()
                return

            owned_wacc_coords = self._owned_named_coords(
                'wacc', self._named_scenario_wacc, self._invalid_named_scenario_wacc
            )
            seen_rounded: set[float] = set()
            wacc_entries: list[tuple[float, str]] = []

            named_wacc = self._named_metrics.get('wacc')
            if named_wacc is not None:
                key = round(named_wacc[0], 8)
                seen_rounded.add(key)
                wacc_entries.append((named_wacc[0], named_wacc[1]))

            for scenario in SCENARIO_NAMES:
                sw = self._named_scenario_wacc.get(scenario)
                if sw is not None:
                    key = round(sw[0], 8)
                    if key not in seen_rounded:
                        seen_rounded.add(key)
                        wacc_entries.append((sw[0], sw[1]))

            # Always merge unnamed scanned WACCs from both WACC and DCF sheets (named wins on coord)
            for _, _, _, entries in wacc_sheet_wacc_rows + dcf_wacc_rows:
                for _, val, coord in entries:
                    if coord in owned_wacc_coords:
                        continue
                    key = round(val, 8)
                    if key not in seen_rounded:
                        seen_rounded.add(key)
                        wacc_entries.append((val, coord))

            invalid_wacc_items: list[tuple[str, Any, str]] = []
            if 'wacc' in self._invalid_named_metrics:
                raw_val, coord = self._invalid_named_metrics['wacc']
                invalid_wacc_items.append(('DCF_WACC', raw_val, coord))
            for scenario in SCENARIO_NAMES:
                if scenario in self._invalid_named_scenario_wacc:
                    raw_val, coord = self._invalid_named_scenario_wacc[scenario]
                    invalid_wacc_items.append((scenario.title(), raw_val, coord))
            self._report_invalid_named(
                invalid_wacc_items,
                severity='error',
                empty_template=(
                    "CRITICAL: Named WACC ({label} at {coord}) is empty — "
                    "enter a valid discount rate."
                ),
                non_numeric_template=(
                    "CRITICAL: Named WACC ({label} at {coord}) has non-numeric value ({raw_val!r})."
                ),
                skip_empty_when_uncached=True,
            )

            if wacc_entries:
                for wacc, coord in wacc_entries:
                    if wacc <= 0:
                        self.errors.append(
                            f"CRITICAL: Non-positive WACC ({wacc:.2%} at {coord}) — "
                            "discount rate must be positive."
                        )
                    elif wacc >= 1.0:
                        self.errors.append(
                            f"CRITICAL: WACC ({wacc:.2%} at {coord}) is >= 100% — "
                            "verify decimal entry (e.g., enter 0.10 for 10%)."
                        )
                    elif wacc < 0.05 or wacc > 0.20:
                        self.warnings.append(
                            f"WACC ({wacc:.2%} at {coord}) is outside typical range (5%-20%). Verify calculation."
                        )
                    else:
                        self.info.append(f"✓ WACC ({wacc:.2%}) in reasonable range")
            elif not invalid_wacc_items:
                self.errors.append("Could not locate WACC value for range check")

        except Exception as e:
            self.errors.append(f"Could not validate WACC range: {str(e)}")

    def _check_terminal_value_proportion(self):
        """Check if terminal value is within the institutional 50%-70% of EV target range."""
        try:
            if self._dcf_grid is None:
                # Missing DCF sheet already recorded as error in check_sheet_structure()
                return

            invalid_tv = self._invalid_named_metrics.get('pv_terminal_value')
            invalid_ev = self._invalid_named_metrics.get('enterprise_value')
            if invalid_tv is not None or invalid_ev is not None:
                invalid_tv_ev_items: list[tuple[str, Any, str]] = []
                for metric_name, invalid_item in (
                    ('PV_TerminalValue', invalid_tv),
                    ('EnterpriseValue', invalid_ev),
                ):
                    if invalid_item is not None:
                        raw_val, coord = invalid_item
                        invalid_tv_ev_items.append((metric_name, raw_val, coord))
                self._report_invalid_named(
                    invalid_tv_ev_items,
                    severity='warning',
                    empty_template=(
                        "Named {label} at {coord} is empty or uncached — "
                        "skipping TV % of EV ratio check."
                    ),
                    non_numeric_template=(
                        "Named {label} at {coord} has non-numeric value ({raw_val!r}) — "
                        "skipping TV % of EV ratio check."
                    ),
                    skip_empty_when_uncached=False,
                )
                return

            named_tv = self._named_metrics.get('pv_terminal_value')
            named_ev = self._named_metrics.get('enterprise_value')
            terminal_value: Optional[float] = named_tv[0] if named_tv is not None else None
            tv_row: Optional[int] = (
                named_tv[3] if (named_tv is not None and named_tv[2].lower() == 'dcf') else None
            )
            enterprise_value: Optional[float] = named_ev[0] if named_ev is not None else None

            # When explicitly bound named ranges are non-positive, warn directly instead of
            # falling through to label-scan unrelated rows (e.g., Exit Multiple TV).
            if enterprise_value is not None and enterprise_value <= 0:
                ev_loc = f" at {named_ev[1]}" if named_ev is not None else ""
                self.warnings.append(
                    f"Enterprise Value ({enterprise_value:,.2f}{ev_loc}) is non-positive — "
                    "skipping TV % of EV ratio check (verify cash flows, WACC, and terminal growth)."
                )
                return
            if terminal_value is not None and terminal_value <= 0:
                tv_loc = f" at {named_tv[1]}" if named_tv is not None else ""
                self.warnings.append(
                    f"PV of Terminal Value ({terminal_value:,.2f}{tv_loc}) is non-positive — "
                    "check terminal value assumptions (e.g., terminal growth >= WACC or negative terminal FCF)."
                )
                return

            if terminal_value is None or enterprise_value is None:
                pv_tv_candidates: list[tuple[int, float]] = []
                ev_candidates: list[tuple[int, float]] = []

                for row_idx in sorted(self._dcf_grid):
                    row_map = self._dcf_grid[row_idx]
                    row_has_numeric = any(_is_number(val) for val, _ in row_map.values())
                    primary_text = _primary_row_text(row_map)
                    if _is_sensitivity_stop_row(primary_text, row_has_numeric):
                        if pv_tv_candidates or ev_candidates or named_tv is not None or named_ev is not None:
                            break
                        continue

                    for col_idx in sorted(row_map):
                        val, _ = row_map[col_idx]
                        if not (val and isinstance(val, str)):
                            continue
                        if _is_pv_tv_label(val):
                            for offset in range(1, 6):
                                adj = row_map.get(col_idx + offset)
                                if adj is not None and _is_number(adj[0]) and adj[0] > 0:
                                    pv_tv_candidates.append((row_idx, float(adj[0])))
                                    break
                        elif _is_ev_label(val):
                            for offset in range(1, 6):
                                adj = row_map.get(col_idx + offset)
                                if adj is not None and _is_number(adj[0]) and adj[0] > 0:
                                    ev_candidates.append((row_idx, float(adj[0])))
                                    break

                if terminal_value is None and pv_tv_candidates:
                    tv_row, terminal_value = pv_tv_candidates[0]

                if enterprise_value is None and ev_candidates:
                    if tv_row is not None:
                        after_tv = [c for c in ev_candidates if c[0] >= tv_row]
                        if after_tv:
                            enterprise_value = after_tv[0][1]
                        else:
                            preceding = [c for c in ev_candidates if 0 < tv_row - c[0] <= 15]
                            if preceding:
                                enterprise_value = min(preceding, key=lambda c: tv_row - c[0])[1]
                    else:
                        enterprise_value = ev_candidates[0][1]

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
                self.warnings.append(
                    "Could not locate PV of Terminal Value and Enterprise Value on DCF sheet; "
                    "define PV_TerminalValue and EnterpriseValue named ranges or use standard labels "
                    "to verify TV % of EV."
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
