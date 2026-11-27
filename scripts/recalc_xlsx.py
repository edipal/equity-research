#!/usr/bin/env python3
"""Recalculate Excel workbook formulas via LibreOffice and scan for formula errors.

Usage:
    python3 scripts/recalc_xlsx.py <path-to-workbook.xlsx> [timeout_seconds] [--allow-unrecalculated]

Outputs a JSON report to stdout and exits:
    0 - status "success" (zero errors, all formula cells verified or --allow-unrecalculated)
    1 - status "errors_found" or "error" (formula errors, missing file, LibreOffice failure/timeout)
    2 - status "unverified" or CLI usage error (uncached formulas without LibreOffice recalculation)
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

EXCEL_ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")
STRING_LITERAL_RE = re.compile(r'"(?:""|[^"])*"')


class WorkbookLoadError(RuntimeError):
    """Raised when an Excel workbook cannot be opened or parsed by openpyxl."""


class JsonArgumentParser(argparse.ArgumentParser):
    """ArgumentParser that emits structured JSON errors instead of raw tracebacks."""

    def error(self, message: str) -> None:  # pragma: no cover - exercised via CLI tests
        print(json.dumps({"status": "error", "message": message}, indent=2))
        raise SystemExit(2)


def positive_int(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"timeout must be an integer, got '{value}'") from exc
    if parsed <= 0:
        raise argparse.ArgumentTypeError(f"timeout must be > 0, got {parsed}")
    return parsed


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(
        prog="recalc_xlsx.py",
        description="Recalculate Excel workbook formulas via LibreOffice and scan for formula errors.",
    )
    parser.add_argument("path", help="Path to Excel workbook (.xlsx, .xlsm, .xls)")
    parser.add_argument(
        "timeout",
        nargs="?",
        default=30,
        type=positive_int,
        help="LibreOffice recalculation timeout in seconds (default: 30)",
    )
    parser.add_argument(
        "--allow-unrecalculated",
        action="store_true",
        help="Allow exit 0 when LibreOffice is unavailable and formulas lack cached values.",
    )
    return parser


def find_soffice() -> str | None:
    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    return None


def atomic_copy(src: Path, dst: Path) -> None:
    """Atomically replace dst with a copy of src via a temp file in dst.parent."""
    fd, tmp_name = tempfile.mkstemp(prefix=".recalc_", suffix=dst.suffix or ".xlsx", dir=dst.parent)
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        shutil.copy2(src, tmp_path)
        os.replace(tmp_path, dst)
    except Exception:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
        raise


def recalculate_with_libreoffice(
    xlsx_path: Path, timeout: int
) -> tuple[bool, str, Path, str | None]:
    """Recalculate workbook with LibreOffice in an isolated user profile.

    Returns:
        (recalculated, lo_status, target_path, message)
        where lo_status is one of: "ok", "missing", "failed".
    """
    soffice = find_soffice()
    if not soffice:
        return (
            False,
            "missing",
            xlsx_path,
            "LibreOffice (soffice) not found on PATH; scanning existing cached values only.",
        )

    with tempfile.TemporaryDirectory(prefix="xlsx_recalc_") as tmpdir:
        tmp_path = Path(tmpdir)
        profile_uri = (tmp_path / "lo_profile").as_uri()
        cmd = [
            soffice,
            f"-env:UserInstallation={profile_uri}",
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--convert-to",
            "xlsx",
            "--outdir",
            tmpdir,
            str(xlsx_path),
        ]
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired:
            return False, "failed", xlsx_path, f"LibreOffice timed out after {timeout}s."
        except OSError as exc:
            return False, "failed", xlsx_path, f"Failed to execute LibreOffice: {exc}"

        recalculated = tmp_path / f"{xlsx_path.stem}.xlsx"
        if proc.returncode != 0 or not recalculated.exists():
            stderr = (proc.stderr or proc.stdout or "").strip()
            return (
                False,
                "failed",
                xlsx_path,
                f"LibreOffice conversion failed (rc={proc.returncode}): {stderr}",
            )

        output_path = (
            xlsx_path if xlsx_path.suffix.lower() == ".xlsx" else xlsx_path.with_suffix(".xlsx")
        )
        try:
            atomic_copy(recalculated, output_path)
        except OSError as exc:
            return (
                False,
                "failed",
                output_path,
                f"Failed to write recalculated workbook to '{output_path}': {exc}",
            )
        note = None
        if output_path != xlsx_path:
            note = (
                f"Converted non-.xlsx input '{xlsx_path.name}' to '{output_path.name}' "
                "without overwriting the original file."
            )
        return True, "ok", output_path, note


def extract_formula_text(val_f) -> str | None:
    """Extract formula text from standard string formulas and openpyxl ArrayFormula objects."""
    if isinstance(val_f, str):
        return val_f if val_f.startswith("=") else None
    if type(val_f).__name__ == "ArrayFormula":
        text = getattr(val_f, "text", None)
        if isinstance(text, str) and text:
            return text if text.startswith("=") else f"={text}"
    return None


def detect_cell_error(cell_v, val_f) -> str | None:
    """Return Excel error token for a cell without false-positiving on prose or quoted string literals."""
    val_v = cell_v.value

    # 1. Cached error value (openpyxl error data_type or exact Excel error token)
    if getattr(cell_v, "data_type", None) == "e":
        return str(val_v) if val_v is not None else "#ERROR!"
    if isinstance(val_v, str) and val_v in EXCEL_ERRORS:
        return val_v

    # 2. Embedded error token inside a formula string (ignoring double-quoted string literals)
    formula_text = extract_formula_text(val_f)
    if formula_text is not None:
        stripped_formula = STRING_LITERAL_RE.sub("", formula_text)
        for err in EXCEL_ERRORS:
            if err in stripped_formula:
                return err

    return None


def scan_workbook(xlsx_path: Path) -> dict:
    try:
        import openpyxl
        from openpyxl.utils.exceptions import InvalidFileException
    except ImportError as exc:
        print(
            json.dumps(
                {
                    "status": "error",
                    "message": f"openpyxl is required: {exc}. Install with `pip install -r requirements.txt`.",
                },
                indent=2,
            )
        )
        raise SystemExit(1)

    wb_formulas = None
    wb_values = None
    try:
        wb_formulas = openpyxl.load_workbook(xlsx_path, data_only=False)
        wb_values = openpyxl.load_workbook(xlsx_path, data_only=True)
    except (
        zipfile.BadZipFile,
        InvalidFileException,
        ET.ParseError,
        KeyError,
        OSError,
        ValueError,
        AttributeError,
    ) as exc:
        if wb_formulas is not None:
            wb_formulas.close()
        if wb_values is not None:
            wb_values.close()
        raise WorkbookLoadError(f"Failed to open workbook '{xlsx_path}': {exc}") from exc

    total_formulas = 0
    uncached_formulas = 0
    error_summary: dict[str, dict] = {}
    total_errors = 0

    try:
        ws_v_by_title = {ws.title: ws for ws in wb_values.worksheets}
        for ws_f in wb_formulas.worksheets:
            sheet_name = ws_f.title
            ws_v = ws_v_by_title.get(sheet_name)
            if ws_v is None:
                continue

            for row in ws_f.iter_rows():
                for cell_f in row:
                    val_f = cell_f.value
                    cell_v = ws_v[cell_f.coordinate]
                    val_v = cell_v.value

                    is_formula = extract_formula_text(val_f) is not None
                    if is_formula:
                        total_formulas += 1
                        if val_v is None and getattr(cell_v, "data_type", None) not in ("s", "str"):
                            uncached_formulas += 1

                    err = detect_cell_error(cell_v, val_f)
                    if err is not None:
                        total_errors += 1
                        bucket = error_summary.setdefault(err, {"count": 0, "locations": []})
                        bucket["count"] += 1
                        bucket["locations"].append(f"{sheet_name}!{cell_f.coordinate}")
    finally:
        wb_formulas.close()
        wb_values.close()

    return {
        "total_formulas": total_formulas,
        "uncached_formulas": uncached_formulas,
        "total_errors": total_errors,
        "error_summary": error_summary,
    }


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    xlsx_path = Path(args.path).resolve()
    if not xlsx_path.exists():
        print(json.dumps({"status": "error", "message": f"File not found: {xlsx_path}"}, indent=2))
        return 1
    if not xlsx_path.is_file():
        print(json.dumps({"status": "error", "message": f"Not a file: {xlsx_path}"}, indent=2))
        return 1

    recalculated, lo_status, target_path, note = recalculate_with_libreoffice(
        xlsx_path, args.timeout
    )

    if lo_status == "failed":
        result: dict = {
            "status": "error",
            "file": str(target_path),
            "recalculated_with_libreoffice": False,
            "message": note,
        }
        print(json.dumps(result, indent=2))
        return 1

    if target_path.suffix.lower() not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        print(
            json.dumps(
                {
                    "status": "error",
                    "file": str(target_path),
                    "recalculated_with_libreoffice": False,
                    "message": (
                        f"Cannot scan '{target_path.suffix}' directly with openpyxl when "
                        "LibreOffice is unavailable. Convert to .xlsx first."
                    ),
                },
                indent=2,
            )
        )
        return 1

    try:
        scan = scan_workbook(target_path)
    except WorkbookLoadError as exc:
        print(
            json.dumps(
                {
                    "status": "error",
                    "file": str(target_path),
                    "recalculated_with_libreoffice": recalculated,
                    "message": str(exc),
                },
                indent=2,
            )
        )
        return 1

    if scan["total_errors"] > 0:
        status = "errors_found"
        exit_code = 1
    elif not recalculated and scan["uncached_formulas"] > 0 and not args.allow_unrecalculated:
        status = "unverified"
        exit_code = 2
    else:
        status = "success"
        exit_code = 0

    result = {
        "status": status,
        "file": str(target_path),
        "recalculated_with_libreoffice": recalculated,
        "total_errors": scan["total_errors"],
        "total_formulas": scan["total_formulas"],
        "uncached_formulas": scan["uncached_formulas"],
        "error_summary": scan["error_summary"],
    }
    if target_path != xlsx_path:
        result["source_file"] = str(xlsx_path)
    if note:
        result["warning"] = note
    if status == "unverified":
        result["message"] = (
            f"Found {scan['uncached_formulas']} uncached formula(s) and LibreOffice "
            "recalculation did not run. Install LibreOffice or pass --allow-unrecalculated "
            "only if you accept unverified cached values."
        )

    print(json.dumps(result, indent=2))
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
