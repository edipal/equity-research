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
from collections.abc import Iterable
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

EXCEL_ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")
OOXML_ZIP_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm")
STRING_LITERAL_RE = re.compile(r'"(?:""|[^"])*"')
SINGLE_CELL_REF_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$")
UNQUOTED_SHEET_REF_RE = re.compile(r"^[^\s!'\"()+\-*/=<>&^%,:;\[\]?\\]+$")


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


_SAFE_EXACT_PARTS = frozenset({
    "[content_types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/styles.xml",
    "xl/sharedstrings.xml",
    "xl/calcchain.xml",
})

_SAFE_PREFIX_PARTS = (
    "docprops/",
    "xl/theme/",
    "xl/printersettings/",
)

_SAFE_REGEX_PARTS = (
    re.compile(r"^xl/worksheets/sheet[^/]+\.xml$"),
    re.compile(r"^xl/worksheets/_rels/sheet[^/]+\.xml\.rels$"),
    # Standard Excel / LibreOffice legacy cell notes (xl/comments1.xml) and openpyxl (xl/comments/comment1.xml)
    re.compile(r"^xl/comments\d*\.xml$"),
    re.compile(r"^xl/_rels/comments\d*\.xml\.rels$"),
    re.compile(r"^xl/comments/comment\d*\.xml$"),
    re.compile(r"^xl/comments/_rels/comment\d*\.xml\.rels$"),
    # Legacy comment VML emitted by openpyxl / Excel / LibreOffice
    re.compile(r"^xl/drawings/(?:vml|comments)drawing\d*\.vml$"),
    re.compile(r"^xl/drawings/_rels/(?:vml|comments)drawing\d*\.vml\.rels$"),
    # Standard Excel tables (query-backed tables still have xl/queryTables/ & xl/connections.xml)
    re.compile(r"^xl/tables/table\d+\.xml$"),
    re.compile(r"^xl/tables/_rels/table\d+\.xml\.rels$"),
)

_LOSSY_WORKSHEET_EXT_URIS: dict[str, str] = {
    "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}": "sparklines",
    "{78C0D931-6437-407D-A8EE-F0AAD7539E65}": "x14_conditional_formatting",
    "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}": "x14_data_validation",
    "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}": "slicers_or_timelines",
    "{7E03D99C-DC04-49D9-9315-930204A7B6E9}": "slicers_or_timelines",
}


_DRAWING_XML_PART_RE = re.compile(r"^xl/drawings/[^/]+\.xml$")
_DRAWING_RELS_PART_RE = re.compile(r"^xl/drawings/_rels/[^/]+\.xml\.rels$")


def _is_safe_ooxml_part(low: str) -> bool:
    """Return True if a normalized lowercase OOXML zip member is known-safe for LibreOffice re-export."""
    if low in _SAFE_EXACT_PARTS:
        return True
    if low.startswith(_SAFE_PREFIX_PARTS):
        return True
    return any(pattern.match(low) for pattern in _SAFE_REGEX_PARTS)


def _is_empty_ooxml_root(stream, expected_root: str) -> bool:
    """Return True iff the XML stream parses cleanly with *expected_root* and zero child elements.

    Streams with ``ET.iterparse`` so large drawing/rels parts never buffer whole trees,
    and lets ``ET.ParseError`` propagate so callers fail closed on malformed XML.
    """
    root_seen = False
    is_empty = True
    for event, elem in ET.iterparse(stream, events=("start", "end")):
        if event == "start":
            if not root_seen:
                root_seen = True
                if _local_tag(elem.tag) != expected_root:
                    is_empty = False
            else:
                is_empty = False
        else:
            if (elem.text or "").strip():
                is_empty = False
            elem.clear()
    return root_seen and is_empty


def _detect_manifest_lossy_features(
    namelist: list[str],
) -> tuple[dict[str, set[str]], list[str], list[str]]:
    """Identify non-allowlisted OOXML archive parts that LibreOffice re-export can drop or alter.

    Uses an allowlist of known-safe core OOXML parts (worksheets, styles, sharedStrings,
    calcChain, themes, docProps, printerSettings, standard Excel tables, and legacy comment
    XML/VML) so unknown or newer Excel parts fail closed as ``unsupported_ooxml_parts`` instead
    of slipping past a denylist. Returns a mapping of ``category -> set[zip_member_path]``
    alongside non-VML ``xl/drawings/*.xml`` and ``xl/drawings/_rels/*.xml.rels`` entries for
    content inspection so empty LibreOffice AutoFilter ``<xdr:wsDr/>`` stubs do not trigger
    false positives while non-empty drawings still fail closed.
    """
    feature_details: dict[str, set[str]] = {}
    drawing_xml_entries: list[str] = []
    drawing_rels_entries: list[str] = []
    for name in namelist:
        if not name or name.endswith("/"):
            continue
        low = name.replace("\\", "/").lstrip("/").lower()
        if not low or _is_safe_ooxml_part(low):
            continue
        if _DRAWING_XML_PART_RE.match(low):
            drawing_xml_entries.append(name)
            continue
        if _DRAWING_RELS_PART_RE.match(low):
            drawing_rels_entries.append(name)
            continue
        if low.startswith(("xl/charts/", "xl/chartsheets/")):
            category = "charts"
        elif low.startswith(("xl/pivottables/", "xl/pivotcache/")):
            category = "pivot_tables"
        elif low.startswith(
            ("xl/slicers/", "xl/slicercaches/", "xl/timelines/", "xl/timelinecaches/")
        ):
            category = "slicers_or_timelines"
        elif low.startswith("xl/drawings/"):
            category = "drawings_or_shapes"
        elif low.startswith(("xl/ctrlprops/", "xl/activex/")):
            category = "form_or_activex_controls"
        elif low.startswith(("xl/threadedcomments/", "xl/persons/")):
            category = "threaded_comments"
        elif (
            low in ("xl/connections.xml", "xl/_rels/connections.xml.rels")
            or low.startswith(("xl/querytables/", "xl/model/", "customxml/"))
        ):
            category = "connections_or_power_query"
        elif low.startswith(("xl/media/", "xl/embeddings/")):
            category = "embedded_media_or_objects"
        elif (
            low in ("xl/metadata.xml", "xl/_rels/metadata.xml.rels")
            or low.startswith("xl/richdata/")
        ):
            category = "excel_metadata_or_rich_data"
        elif low.startswith("xl/externallinks/"):
            category = "external_links"
        elif (
            low in ("xl/vbaproject.bin", "xl/vbadata.xml")
            or low.endswith(("/vbaproject.bin", "/vbadata.xml"))
            or low.startswith(("xl/vbaproject", "xl/vbadata", "xl/macrosheets/", "xl/dialogsheets/"))
        ):
            category = "vba_macros"
        else:
            category = "unsupported_ooxml_parts"
        feature_details.setdefault(category, set()).add(name)
    return feature_details, drawing_xml_entries, drawing_rels_entries


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
        - ``"lossy_features"``: sorted list of non-allowlisted OOXML parts and worksheet
          ``<extLst>`` extension categories that LibreOffice re-export can drop or alter
          (e.g., ``charts``, ``connections_or_power_query``, ``drawings_or_shapes``,
          ``embedded_media_or_objects``, ``excel_metadata_or_rich_data``,
          ``external_links``, ``form_or_activex_controls``, ``pivot_tables``,
          ``slicers_or_timelines``, ``sparklines``, ``threaded_comments``,
          ``unsupported_ooxml_parts``, ``vba_macros``, ``worksheet_extensions``,
          ``x14_conditional_formatting``, ``x14_data_validation``).
        - ``"lossy_feature_details"``: dict mapping each flagged category in
          ``lossy_features`` to the sorted list of exact offending zip member paths or
          ``<sheet_xml>:<ext_uri>`` / ``<sheet_xml>:sparklineGroups`` /
          ``<sheet_xml>:<missing-uri>`` identifiers.
        - ``"lossy_features_inspected"``: ``True`` when the workbook is an OOXML archive
          whose manifest and worksheet XML were pre-inspected, or ``False`` for non-OOXML
          formats (e.g., ``.xls``, ``.ods``) that cannot be pre-inspected.

    Fails closed by raising ``WorkbookLoadError`` on corrupt OOXML archives or malformed XML.
    """
    if xlsx_path.suffix.lower() not in OOXML_ZIP_SUFFIXES:
        return {
            "data_tables": [],
            "uncached_formulas": None,
            "lossy_features": [],
            "lossy_feature_details": {},
            "lossy_features_inspected": False,
        }
    try:
        locations: list[str] = []
        uncached_formulas = 0
        with zipfile.ZipFile(xlsx_path, "r") as zf:
            namelist = zf.namelist()
            lossy_details, drawing_xml_entries, drawing_rels_entries = _detect_manifest_lossy_features(
                namelist
            )
            for entry in drawing_xml_entries:
                with zf.open(entry, "r") as stream:
                    if not _is_empty_ooxml_root(stream, "wsDr"):
                        lossy_details.setdefault("drawings_or_shapes", set()).add(entry)
            for entry in drawing_rels_entries:
                with zf.open(entry, "r") as stream:
                    if not _is_empty_ooxml_root(stream, "Relationships"):
                        lossy_details.setdefault("drawings_or_shapes", set()).add(entry)
            titles = _sheet_titles_by_xml_path(zf)
            sheet_entries = sorted(
                name
                for name in namelist
                if name.lower().startswith("xl/worksheets/") and name.lower().endswith(".xml")
            )
            for entry in sheet_entries:
                sheet_label = titles.get(entry, Path(entry).stem)
                with zf.open(entry, "r") as stream:
                    in_cell = False
                    ext_lst_depth = 0
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
                            elif local == "extLst":
                                ext_lst_depth += 1
                            elif ext_lst_depth > 0 and local == "ext":
                                uri = (elem.attrib.get("uri") or "").strip().upper()
                                category = _LOSSY_WORKSHEET_EXT_URIS.get(
                                    uri, "worksheet_extensions"
                                )
                                uri_label = uri or "<missing-uri>"
                                lossy_details.setdefault(category, set()).add(
                                    f"{entry}:{uri_label}"
                                )
                            elif local == "sparklineGroups":
                                lossy_details.setdefault("sparklines", set()).add(
                                    f"{entry}:sparklineGroups"
                                )
                        else:
                            if local == "extLst" and ext_lst_depth > 0:
                                ext_lst_depth -= 1
                            elif in_cell and local == "v":
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
            "lossy_features": sorted(lossy_details),
            "lossy_feature_details": {
                category: sorted(details)
                for category, details in sorted(lossy_details.items())
            },
            "lossy_features_inspected": True,
        }
    except (zipfile.BadZipFile, ET.ParseError, KeyError, OSError) as exc:
        raise WorkbookLoadError(
            f"Failed to inspect workbook '{xlsx_path}' XML (What-If data tables / uncached formulas): {exc}"
        ) from exc


def scan_workbook_formulas(
    xlsx_path: Path,
    *,
    capture_sheet_grids: bool = False,
    grid_sheet_names: Iterable[str] | None = None,
    capture_defined_names: bool = False,
) -> dict[str, Any]:
    """Scan all worksheets in lockstep using openpyxl read_only=True streaming mode.

    Returns a dictionary with:
        total_formulas, uncached_formulas, data_table_formulas, total_errors, error_summary.

    When *capture_sheet_grids* is True, also includes ``sheetnames`` (worksheet titles only,
    excluding chartsheets) and ``sheet_grids`` (sparse ``row -> col -> (value, coord)`` maps)
    collected during the same single streaming pass. Full grids are captured for all sheets
    when *grid_sheet_names* is ``None``, or only for sheets matching *grid_sheet_names*
    case-insensitively; non-target sheets referenced by captured single-cell defined names are
    point-captured (storing only the targeted ``(row, col)`` cells, or ``{}`` when blank).

    When *capture_defined_names* is True, also includes ``defined_names`` mapping workbook-
    and worksheet-scoped defined name strings to resolved single-cell ``(sheet_name, row, col)``
    tuples (precedence: workbook-global > DCF/WACC sheets > remaining sheets).
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

        full_grid_sheets_lower: set[str] | None = (
            None if grid_sheet_names is None else {name.lower() for name in grid_sheet_names}
        )
        named_coords_by_sheet_lower: dict[str, set[tuple[int, int]]] = {}
        if capture_defined_names:
            resolved_names_lower: set[str] = set()

            def _collect_defined_names(dn_container, owner_sheet_title: str | None = None) -> None:
                if dn_container is None:
                    return
                for dn_name, dn in dn_container.items():
                    dn_lower = dn_name.lower()
                    if dn_lower in resolved_names_lower:
                        continue
                    destinations: list[tuple[str, str]] = []
                    try:
                        destinations = list(dn.destinations)
                    except Exception:
                        destinations = []
                    if not destinations:
                        attr_text = getattr(dn, "attr_text", None)
                        if isinstance(attr_text, str):
                            raw_ref = attr_text.strip()
                            if raw_ref.startswith("="):
                                raw_ref = raw_ref[1:].strip()
                            if raw_ref and ":" not in raw_ref:
                                if "!" in raw_ref:
                                    sheet_part, cell_part = raw_ref.rsplit("!", 1)
                                    sheet_part = sheet_part.strip()
                                    cell_part = cell_part.strip()
                                    if SINGLE_CELL_REF_RE.fullmatch(cell_part):
                                        if (
                                            len(sheet_part) >= 2
                                            and sheet_part[0] == "'"
                                            and sheet_part[-1] == "'"
                                        ):
                                            inner_sheet = sheet_part[1:-1]
                                            if (
                                                inner_sheet
                                                and "[" not in inner_sheet
                                                and "'" not in inner_sheet.replace("''", "")
                                            ):
                                                destinations = [(inner_sheet, cell_part)]
                                        elif (
                                            sheet_part
                                            and "[" not in sheet_part
                                            and UNQUOTED_SHEET_REF_RE.fullmatch(sheet_part)
                                        ):
                                            destinations = [(sheet_part, cell_part)]
                                elif owner_sheet_title and SINGLE_CELL_REF_RE.fullmatch(raw_ref):
                                    destinations = [(owner_sheet_title, raw_ref)]
                    for sheet_title, ref in destinations:
                        if not ref or ":" in ref or "[" in sheet_title:
                            continue
                        ref_stripped = ref.strip()
                        if not SINGLE_CELL_REF_RE.fullmatch(ref_stripped):
                            continue
                        ref_clean = ref_stripped.replace("$", "")
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
                        resolved_names_lower.add(dn_lower)
                        canon_lower = canon_title.lower()
                        if full_grid_sheets_lower is not None and canon_lower not in full_grid_sheets_lower:
                            named_coords_by_sheet_lower.setdefault(canon_lower, set()).add((r_idx, c_idx))
                        break

            # Precedence: workbook-global > DCF/WACC sheet-scoped > other sheet-scoped
            _collect_defined_names(wb_values.defined_names)
            priority_sheets = [
                ws for ws in wb_values.worksheets if ws.title.lower() in ("dcf", "wacc")
            ]
            other_sheets = [
                ws for ws in wb_values.worksheets if ws.title.lower() not in ("dcf", "wacc")
            ]
            for ws in priority_sheets + other_sheets:
                _collect_defined_names(getattr(ws, "defined_names", None), owner_sheet_title=ws.title)

        for ws_f in wb_formulas.worksheets:
            sheet_name = ws_f.title
            ws_v = ws_v_by_title.get(sheet_name)
            if ws_v is None:
                continue

            # Reset cached <dimension> bounds so stale metadata never hides cells in read_only mode.
            ws_f.reset_dimensions()
            ws_v.reset_dimensions()

            sheet_lower = sheet_name.lower()
            capture_full_grid = capture_sheet_grids and (
                full_grid_sheets_lower is None or sheet_lower in full_grid_sheets_lower
            )
            target_coords = (
                named_coords_by_sheet_lower.get(sheet_lower)
                if (capture_sheet_grids and not capture_full_grid)
                else None
            )
            target_cols_by_row: dict[int, set[int]] | None = None
            if target_coords is not None:
                target_cols_by_row = {}
                for r_i, c_i in target_coords:
                    target_cols_by_row.setdefault(r_i, set()).add(c_i)

            current_grid: dict[int, dict[int, tuple[Any, str]]] | None = (
                {} if (capture_full_grid or target_coords is not None) else None
            )

            for row_idx, (row_f, row_v) in enumerate(
                itertools.zip_longest(ws_f.iter_rows(), ws_v.iter_rows(), fillvalue=()),
                start=1,
            ):
                if capture_full_grid:
                    row_target_cols: set[int] | None = None
                    row_map: dict[int, tuple[Any, str]] | None = {}
                elif target_cols_by_row is not None and row_idx in target_cols_by_row:
                    row_target_cols = target_cols_by_row[row_idx]
                    row_map = {}
                else:
                    row_target_cols = None
                    row_map = None

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

                    if (
                        row_map is not None
                        and (row_target_cols is None or col_idx in row_target_cols)
                        and cell_v is not None
                        and cell_v.value is not None
                    ):
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
