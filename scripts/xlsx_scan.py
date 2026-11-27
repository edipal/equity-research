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


def extract_formula_text(val_f) -> str | None:
    """Extract formula text from standard strings, ArrayFormula, and DataTableFormula objects."""
    if isinstance(val_f, str):
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


def detect_cell_error(cell_v, val_f) -> str | None:
    """Return Excel error token for a cell without false-positiving on prose, text cells, or quoted literals.

    Checks cached error values first:
    - Any cell with Excel error type (``data_type == "e"``) is always flagged.
    - For formula cells (``formula_text is not None``), a cached string matching an exact
      Excel error token is also flagged.
    - Plain non-formula text cells (``formula_text is None`` and ``data_type != "e"``)
      containing literal strings such as ``"#N/A"`` are NOT flagged as formula errors.

    When the cached result is not an error, only scans stripped formula text for ``#REF!``
    (a broken reference is always a static formula error, whereas ``#N/A`` or other error
    tokens inside live branches such as ``=IF(ISBLANK(B5),#N/A,B5)`` are valid when the
    evaluated cached value is non-error).
    """
    val_v = cell_v.value if cell_v is not None else None
    formula_text = extract_formula_text(val_f)

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


def inspect_workbook_xml(xlsx_path: Path) -> dict[str, Any]:
    """Inspect OOXML worksheets in a single low-memory streaming XML pass.

    Returns a dictionary with:
        - ``"data_tables"``: list of Excel What-If data table locations (``Sheet!ref``).
        - ``"uncached_formulas"``: count of formula ``<c>`` cells lacking a cached
          ``<v>`` value (treating ``<c t="str|s"><f>...</f><v/></c>`` as a cached empty
          string; matches ``is_uncached_formula()`` for standard Excel/LibreOffice/openpyxl
          OOXML output, while conservatively counting malformed ``t="str"``/``t="e"``
          formula cells with no ``<v>`` element as uncached), or ``None`` for non-OOXML
          legacy formats such as ``.xls``.

    Fails closed by raising ``WorkbookLoadError`` on corrupt OOXML archives or malformed XML.
    """
    if xlsx_path.suffix.lower() not in OOXML_ZIP_SUFFIXES:
        return {"data_tables": [], "uncached_formulas": None}
    try:
        locations: list[str] = []
        uncached_formulas = 0
        with zipfile.ZipFile(xlsx_path, "r") as zf:
            titles = _sheet_titles_by_xml_path(zf)
            sheet_entries = sorted(
                name
                for name in zf.namelist()
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


def scan_workbook_formulas(xlsx_path: Path) -> dict:
    """Scan all worksheets in lockstep using openpyxl read_only=True streaming mode.

    Returns a dictionary with:
        total_formulas, uncached_formulas, data_table_formulas, total_errors, error_summary.
    """
    import openpyxl
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.exceptions import InvalidFileException

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

    try:
        ws_v_by_title = {ws.title: ws for ws in wb_values.worksheets}
        for ws_f in wb_formulas.worksheets:
            sheet_name = ws_f.title
            ws_v = ws_v_by_title.get(sheet_name)
            if ws_v is None:
                continue

            # Reset cached <dimension> bounds so stale metadata never hides cells in read_only mode.
            ws_f.reset_dimensions()
            ws_v.reset_dimensions()

            for row_idx, (row_f, row_v) in enumerate(
                itertools.zip_longest(ws_f.iter_rows(), ws_v.iter_rows(), fillvalue=()),
                start=1,
            ):
                for col_idx, (cell_f, cell_v) in enumerate(
                    itertools.zip_longest(row_f, row_v, fillvalue=None),
                    start=1,
                ):
                    val_f = cell_f.value if cell_f is not None else None
                    formula_text = extract_formula_text(val_f)
                    if formula_text is not None:
                        total_formulas += 1
                        if is_data_table_formula(val_f):
                            data_table_formulas += 1
                        if is_uncached_formula(cell_v, formula_text):
                            uncached_formulas += 1

                    err = detect_cell_error(cell_v, val_f)
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

    return {
        "total_formulas": total_formulas,
        "uncached_formulas": uncached_formulas,
        "data_table_formulas": data_table_formulas,
        "total_errors": total_errors,
        "error_summary": error_summary,
    }
