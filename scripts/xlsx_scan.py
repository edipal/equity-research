#!/usr/bin/env python3
# Independent implementation for Google Antigravity, licensed under the Apache License, Version 2.0.
"""Shared low-memory Excel formula and error scanning helpers.

Used by both ``scripts/recalc_xlsx.py`` and
``.agents/skills/dcf-model/scripts/validate_dcf.py`` so formula detection,
What-If data table detection, uncached-value checks, and error classification
stay strictly consistent across all QA tools.
"""

from __future__ import annotations

import itertools
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

EXCEL_ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")
OOXML_ZIP_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm")
STRING_LITERAL_RE = re.compile(r'"(?:""|[^"])*"')


class WorkbookLoadError(RuntimeError):
    """Raised when an Excel workbook cannot be opened or parsed by openpyxl."""


def is_data_table_formula(val) -> bool:
    """Return True if *val* is an openpyxl DataTableFormula (<f t="dataTable">)."""
    return type(val).__name__ == "DataTableFormula" or getattr(val, "t", None) == "dataTable"


def extract_formula_text(val_f: Any, data_type: str | None = None) -> str | None:
    """Extract formula text from standard strings, ArrayFormula, and DataTableFormula objects.

    When *data_type* is supplied (e.g. ``cell.data_type`` from openpyxl), plain string values
    starting with ``"="`` are only treated as formulas when ``data_type == "f"``, preventing
    text cells (``data_type`` in ``("s", "str", "inlineStr", ...)``) from being misclassified
    as formulas.
    """
    if isinstance(val_f, str):
        if data_type is not None and data_type != "f":
            return None
        return val_f if val_f.startswith("=") else None
    if type(val_f).__name__ == "ArrayFormula":
        text = getattr(val_f, "text", None)
        if isinstance(text, str) and text:
            return text if text.startswith("=") else f"={text}"
        return None
    if is_data_table_formula(val_f):
        r1 = getattr(val_f, "r1", None) or ""
        r2 = getattr(val_f, "r2", None) or ""
        return f"=TABLE({r1},{r2})"
    return None


def detect_cell_error(
    cell_v,
    val_f,
    formula_data_type: str | None = None,
) -> str | None:
    """Return Excel error token for a cell without false-positiving on prose, text cells, or quoted literals.

    Checks cached error values first:
    - Any cell with Excel error type (``data_type == "e"``) is always flagged.
    - For formula cells (``formula_text is not None``), a cached string matching an exact
      Excel error token is also flagged.
    - Plain non-formula text cells (``formula_text is None`` and ``data_type != "e"``)
      containing literal strings such as ``"#N/A"`` or ``"== Notes: old link was #REF! =="``
      are NOT flagged as formula errors.

    When the cached result is not an error, only scans stripped formula text for ``#REF!``
    (a broken reference is always a static formula error, whereas ``#N/A`` or other error
    tokens inside live branches such as ``=IF(ISBLANK(B5),#N/A,B5)`` are valid when the
    evaluated cached value is non-error).
    """
    val_v = cell_v.value if cell_v is not None else None
    formula_text = extract_formula_text(val_f, data_type=formula_data_type)

    # 1. Cached error value (Excel error data_type, or exact error token cached on a formula cell)
    if getattr(cell_v, "data_type", None) == "e":
        return str(val_v) if val_v is not None else "#ERROR!"
    if formula_text is not None and isinstance(val_v, str) and val_v in EXCEL_ERRORS:
        return val_v

    # 2. Static broken reference (#REF!) inside formula text (ignoring double-quoted string literals)
    if formula_text is not None:
        stripped_formula = STRING_LITERAL_RE.sub("", formula_text)
        if "#REF!" in stripped_formula:
            return "#REF!"

    return None


def is_uncached_formula(cell_v, formula_text: str | None) -> bool:
    """Return True when a formula cell has no cached result (preserving cached empty strings)."""
    if formula_text is None:
        return False
    if cell_v is None:
        return True
    val_v = cell_v.value
    # Excel/LibreOffice store cached empty-string formula outputs as <c t="str"><v/></c>,
    # which openpyxl loads as value=None with data_type='str'. Error-typed cells (data_type='e')
    # also carry a cached error type rather than a missing cache.
    return val_v is None and getattr(cell_v, "data_type", None) not in ("s", "str", "e")


def _local_tag(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _sheet_titles_by_xml_path(zf: zipfile.ZipFile) -> dict[str, str]:
    """Map OOXML worksheet zip paths (e.g. 'xl/worksheets/sheet1.xml') to sheet titles."""
    mapping: dict[str, str] = {}
    if "xl/workbook.xml" not in zf.namelist() or "xl/_rels/workbook.xml.rels" not in zf.namelist():
        return mapping
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_targets: dict[str, str] = {}
    for rel in rels_root:
        if _local_tag(rel.tag) != "Relationship":
            continue
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            norm = posixpath.normpath(posixpath.join("xl", target.lstrip("/")))
            if target.startswith("/"):
                norm = target.lstrip("/")
            rel_targets[rid] = norm

    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    for elem in wb_root.iter():
        if _local_tag(elem.tag) != "sheet":
            continue
        name = elem.attrib.get("name")
        rid = None
        for attr_k, attr_v in elem.attrib.items():
            if _local_tag(attr_k) == "id":
                rid = attr_v
                break
        if name and rid and rid in rel_targets:
            mapping[rel_targets[rid]] = name
    return mapping


def _detect_manifest_lossy_features(namelist: list[str]) -> set[str]:
    """Identify Excel features in the OOXML archive manifest that LibreOffice re-export can drop or alter.

    Legacy comment VML files (``xl/drawings/commentsDrawing*.vml`` and
    ``xl/drawings/vmlDrawing*.vml``) are intentionally excluded from ``drawings_or_shapes``
    so workbooks with standard blue-input source cell comments still support ``--in-place``.
    """
    features: set[str] = set()
    for name in namelist:
        low = name.lower()
        if (
            low.startswith("xl/charts/") and not low.startswith("xl/charts/_rels/")
        ) or (
            low.startswith("xl/chartsheets/") and not low.startswith("xl/chartsheets/_rels/")
        ):
            features.add("charts")
        elif low.startswith(("xl/pivottables/", "xl/pivotcache/")):
            features.add("pivot_tables")
        elif low.startswith(
            ("xl/slicers/", "xl/slicercaches/", "xl/timelines/", "xl/timelinecaches/")
        ):
            features.add("slicers_or_timelines")
        elif (
            low.startswith("xl/drawings/")
            and not low.startswith("xl/drawings/_rels/")
            and low.endswith(".xml")
        ):
            features.add("drawings_or_shapes")
        elif low == "xl/vbaproject.bin" or low.endswith("/vbaproject.bin"):
            features.add("vba_macros")
        elif low.startswith(("xl/ctrlprops/", "xl/activex/")):
            features.add("form_or_activex_controls")
    return features


def inspect_workbook_xml(xlsx_path: Path) -> dict[str, Any]:
    """Inspect OOXML worksheets and archive manifest in a single low-memory streaming XML pass.

    Returns a dictionary with:
        - ``"data_tables"``: list of Excel What-If data table locations (``Sheet!ref``).
        - ``"uncached_formulas"``: count of formula ``<c>`` cells lacking a cached
          ``<v>`` value (treating ``<c t="str|s"><f>...</f><v/></c>`` as a cached empty
          string; matches ``is_uncached_formula()`` for standard Excel/LibreOffice/openpyxl
          OOXML output, while conservatively counting malformed ``t="str"``/``t="e"``
          formula cells with no ``<v>`` element as uncached), or ``None`` for non-OOXML
          legacy formats such as ``.xls``.
        - ``"lossy_features"``: sorted list of Excel-only feature categories detected in
          the archive/XML that LibreOffice re-export can drop or alter (``charts``,
          ``drawings_or_shapes``, ``form_or_activex_controls``, ``pivot_tables``,
          ``slicers_or_timelines``, ``sparklines``, ``vba_macros``).

    Fails closed by raising ``WorkbookLoadError`` on corrupt OOXML archives or malformed XML.
    """
    if xlsx_path.suffix.lower() not in OOXML_ZIP_SUFFIXES:
        return {"data_tables": [], "uncached_formulas": None, "lossy_features": []}
    try:
        locations: list[str] = []
        uncached_formulas = 0
        with zipfile.ZipFile(xlsx_path, "r") as zf:
            namelist = zf.namelist()
            lossy_features = _detect_manifest_lossy_features(namelist)
            titles = _sheet_titles_by_xml_path(zf)
            sheet_entries = sorted(
                name
                for name in namelist
                if name.startswith("xl/worksheets/") and name.endswith(".xml")
            )
            for entry in sheet_entries:
                sheet_label = titles.get(entry, Path(entry).stem)
                with zf.open(entry, "r") as stream:
                    in_cell = False
                    current_cell_ref: str | None = None
                    current_cell_type = ""
                    cell_has_formula = False
                    cell_has_v = False
                    cell_v_text: str | None = None
                    for event, elem in ET.iterparse(stream, events=("start", "end")):
                        local = _local_tag(elem.tag)
                        if event == "start":
                            if local == "c":
                                in_cell = True
                                current_cell_ref = elem.attrib.get("r")
                                current_cell_type = elem.attrib.get("t", "")
                                cell_has_formula = False
                                cell_has_v = False
                                cell_v_text = None
                            elif in_cell and local == "f":
                                cell_has_formula = True
                                if elem.attrib.get("t") == "dataTable":
                                    ref = elem.attrib.get("ref") or current_cell_ref or "dataTable"
                                    locations.append(f"{sheet_label}!{ref}")
                            elif local == "sparklineGroups":
                                lossy_features.add("sparklines")
                        else:
                            if in_cell and local == "v":
                                if not cell_has_v:
                                    cell_has_v = True
                                    cell_v_text = elem.text
                            elif local == "c":
                                if cell_has_formula:
                                    if cell_has_v and cell_v_text not in (None, ""):
                                        pass  # cached value present
                                    elif current_cell_type in ("str", "s") and cell_has_v:
                                        pass  # cached empty string ""
                                    else:
                                        uncached_formulas += 1
                                in_cell = False
                                current_cell_ref = None
                                current_cell_type = ""
                                cell_has_formula = False
                                cell_has_v = False
                                cell_v_text = None
                            elem.clear()
        return {
            "data_tables": locations,
            "uncached_formulas": uncached_formulas,
            "lossy_features": sorted(lossy_features),
        }
    except (zipfile.BadZipFile, ET.ParseError, KeyError, OSError) as exc:
        raise WorkbookLoadError(
            f"Failed to inspect workbook '{xlsx_path}' XML (What-If data tables / uncached formulas): {exc}"
        ) from exc


def workbook_has_data_tables(xlsx_path: Path) -> list[str]:
    """Return list of Excel What-If data table locations (<f t="dataTable">) in an OOXML workbook.

    Thin wrapper around ``inspect_workbook_xml`` so existing callers share the same
    fail-closed streaming XML inspection.
    """
    return inspect_workbook_xml(xlsx_path)["data_tables"]


_DEFAULT_CAPTURED_GRID_SHEETS = frozenset({"dcf", "wacc", "sensitivity"})


def scan_workbook_formulas(
    xlsx_path: Path,
    *,
    capture_sheet_grids: bool = False,
    capture_defined_names: bool = False,
) -> dict[str, Any]:
    """Scan all worksheets in lockstep using openpyxl read_only=True streaming mode.

    Returns a dictionary with:
        total_formulas, uncached_formulas, data_table_formulas, total_errors, error_summary.

    When *capture_sheet_grids* is True, also includes ``sheetnames`` (worksheet titles only,
    excluding chartsheets) and ``sheet_grids`` (sparse ``row -> col -> (value, coord)`` maps
    for DCF / WACC / Sensitivity worksheets and any sheet targeted by captured defined names)
    collected during the same single streaming pass.

    When *capture_defined_names* is True, also includes ``defined_names`` mapping workbook
    defined name strings to resolved single-cell ``(sheet_name, row, col)`` tuples.
    """
    try:
        import openpyxl
        from openpyxl.utils import get_column_letter
        from openpyxl.utils.cell import coordinate_to_tuple
        from openpyxl.utils.exceptions import InvalidFileException
    except ImportError as exc:
        raise WorkbookLoadError(
            f"openpyxl is required: {exc}. Install with `pip install -r requirements.txt`."
        ) from exc

    wb_formulas = None
    wb_values = None
    try:
        wb_formulas = openpyxl.load_workbook(xlsx_path, data_only=False, read_only=True)
        wb_values = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    except (
        zipfile.BadZipFile,
        InvalidFileException,
        ET.ParseError,
        KeyError,
        OSError,
        ValueError,
        AttributeError,
        TypeError,
    ) as exc:
        if wb_formulas is not None:
            wb_formulas.close()
        if wb_values is not None:
            wb_values.close()
        raise WorkbookLoadError(f"Failed to open workbook '{xlsx_path}': {exc}") from exc

    total_formulas = 0
    uncached_formulas = 0
    data_table_formulas = 0
    error_summary: dict[str, dict] = {}
    total_errors = 0
    sheetnames: list[str] = []
    sheet_grids: dict[str, dict[int, dict[int, tuple[Any, str]]]] = {}
    defined_names_map: dict[str, tuple[str, int, int]] = {}

    try:
        ws_v_by_title = {ws.title: ws for ws in wb_values.worksheets}
        title_by_lower = {ws.title.lower(): ws.title for ws in wb_values.worksheets}
        if capture_sheet_grids or capture_defined_names:
            sheetnames = [ws.title for ws in wb_values.worksheets]

        extra_grid_sheets_lower: set[str] = set()
        if capture_defined_names:
            for dn_name, dn in wb_values.defined_names.items():
                try:
                    destinations = list(dn.destinations)
                except Exception:
                    continue
                for sheet_title, ref in destinations:
                    if not ref or ":" in ref:
                        continue
                    ref_clean = ref.replace("$", "").strip()
                    try:
                        r_idx, c_idx = coordinate_to_tuple(ref_clean)
                    except Exception:
                        continue
                    unescaped_title = sheet_title.replace("''", "'")
                    canon_title = (
                        title_by_lower.get(unescaped_title.lower())
                        or title_by_lower.get(sheet_title.lower())
                        or unescaped_title
                    )
                    defined_names_map[dn_name] = (canon_title, r_idx, c_idx)
                    extra_grid_sheets_lower.add(canon_title.lower())
                    break

        target_grid_sheets_lower = _DEFAULT_CAPTURED_GRID_SHEETS | extra_grid_sheets_lower

        for ws_f in wb_formulas.worksheets:
            sheet_name = ws_f.title
            ws_v = ws_v_by_title.get(sheet_name)
            if ws_v is None:
                continue

            # Reset cached <dimension> bounds so stale metadata never hides cells in read_only mode.
            ws_f.reset_dimensions()
            ws_v.reset_dimensions()

            capture_this_grid = (
                capture_sheet_grids and sheet_name.lower() in target_grid_sheets_lower
            )
            current_grid: dict[int, dict[int, tuple[Any, str]]] | None = (
                {} if capture_this_grid else None
            )

            for row_idx, (row_f, row_v) in enumerate(
                itertools.zip_longest(ws_f.iter_rows(), ws_v.iter_rows(), fillvalue=()),
                start=1,
            ):
                row_map: dict[int, tuple[Any, str]] | None = {} if capture_this_grid else None
                for col_idx, (cell_f, cell_v) in enumerate(
                    itertools.zip_longest(row_f, row_v, fillvalue=None),
                    start=1,
                ):
                    val_f = cell_f.value if cell_f is not None else None
                    data_type_f = getattr(cell_f, "data_type", None) if cell_f is not None else None
                    formula_text = extract_formula_text(val_f, data_type=data_type_f)
                    if formula_text is not None:
                        total_formulas += 1
                        if is_data_table_formula(val_f):
                            data_table_formulas += 1
                        if is_uncached_formula(cell_v, formula_text):
                            uncached_formulas += 1

                    err = detect_cell_error(cell_v, val_f, formula_data_type=data_type_f)
                    if err is not None:
                        coord = (
                            getattr(cell_f, "coordinate", None)
                            or getattr(cell_v, "coordinate", None)
                            or f"{get_column_letter(col_idx)}{row_idx}"
                        )
                        total_errors += 1
                        bucket = error_summary.setdefault(err, {"count": 0, "locations": []})
                        bucket["count"] += 1
                        bucket["locations"].append(f"{sheet_name}!{coord}")

                    if row_map is not None and cell_v is not None and cell_v.value is not None:
                        coord_v = (
                            getattr(cell_v, "coordinate", None)
                            or f"{get_column_letter(col_idx)}{row_idx}"
                        )
                        row_map[col_idx] = (cell_v.value, coord_v)

                if current_grid is not None and row_map:
                    current_grid[row_idx] = row_map

            if current_grid is not None:
                sheet_grids[sheet_name] = current_grid
    except (
        zipfile.BadZipFile,
        ET.ParseError,
        KeyError,
        OSError,
        ValueError,
        AttributeError,
        TypeError,
    ) as exc:
        raise WorkbookLoadError(f"Failed to scan workbook '{xlsx_path}': {exc}") from exc
    finally:
        wb_formulas.close()
        wb_values.close()

    result: dict[str, Any] = {
        "total_formulas": total_formulas,
        "uncached_formulas": uncached_formulas,
        "data_table_formulas": data_table_formulas,
        "total_errors": total_errors,
        "error_summary": error_summary,
    }
    if capture_sheet_grids:
        result["sheetnames"] = sheetnames
        result["sheet_grids"] = sheet_grids
    if capture_defined_names:
        result["defined_names"] = defined_names_map
    return result
