#!/usr/bin/env python3
# Independent implementation for Google Antigravity, licensed under the Apache License, Version 2.0.
"""Recalculate Excel workbook formulas via LibreOffice and scan for formula errors.

Usage:
    python3 scripts/recalc_xlsx.py <path-to-workbook.xlsx> [timeout_seconds]
        [--allow-unrecalculated] [--check-only | --output PATH]

Outputs a JSON report to stdout and exits:
    0 - status "success" (zero errors, all formula cells verified or --allow-unrecalculated)
    1 - status "errors_found" or "error" (formula errors, missing/corrupt file, LibreOffice failure/timeout)
    2 - status "unverified" or CLI usage error (uncached formulas without LibreOffice recalculation)
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import signal
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

EXCEL_ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")
OOXML_ZIP_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm")
STRING_LITERAL_RE = re.compile(r'"(?:""|[^"])*"')

REGISTRY_MODIFICATIONS_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <item oor:path="/org.openoffice.Office.Calc/Formula/Load">
    <prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop>
  </item>
  <item oor:path="/org.openoffice.Office.Calc/Formula/Load">
    <prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop>
  </item>
</oor:items>
"""


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
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        "--check-only",
        action="store_true",
        help="Recalculate in a temporary directory and scan for errors without modifying or writing any file.",
    )
    mode_group.add_argument(
        "--output",
        metavar="PATH",
        help="Write the recalculated workbook to PATH without modifying the input file.",
    )
    return parser


def find_soffice() -> str | None:
    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    return None


def write_libreoffice_recalc_profile(profile_dir: Path) -> Path:
    """Create an isolated LibreOffice profile that forces full recalculation on load."""
    user_dir = profile_dir / "user"
    user_dir.mkdir(parents=True, exist_ok=True)
    if os.name != "nt":
        os.chmod(profile_dir, 0o700)
        os.chmod(user_dir, 0o700)
    xcu_path = user_dir / "registrymodifications.xcu"
    xcu_path.write_text(REGISTRY_MODIFICATIONS_XCU, encoding="utf-8")
    if os.name != "nt":
        os.chmod(xcu_path, 0o600)
    return xcu_path


def preflight_validate_input(xlsx_path: Path) -> str | None:
    """Validate that input exists, is a non-empty file, and is a valid zip for OOXML formats."""
    if not xlsx_path.exists():
        return f"File not found: {xlsx_path}"
    if not xlsx_path.is_file():
        return f"Not a file: {xlsx_path}"
    try:
        size = xlsx_path.stat().st_size
    except OSError as exc:
        return f"Failed to stat workbook '{xlsx_path}': {exc}"
    if size <= 0:
        return f"Failed to open workbook '{xlsx_path}': file is empty (0 bytes)."
    if xlsx_path.suffix.lower() in OOXML_ZIP_SUFFIXES and not zipfile.is_zipfile(xlsx_path):
        return (
            f"Failed to open workbook '{xlsx_path}': file is not a valid OOXML zip archive."
        )
    return None


def atomic_copy(src: Path, dst: Path, preserve_mode_from: Path | None = None) -> None:
    """Atomically replace dst with a copy of src via a temp file in dst.parent.

    Preserves permission bits from dst (when overwriting in place) or preserve_mode_from
    without loosening permissions via LibreOffice's temp file mode.
    """
    mode_source: Path | None = None
    if dst.exists():
        mode_source = dst
    elif preserve_mode_from is not None and preserve_mode_from.exists():
        mode_source = preserve_mode_from

    orig_mode: int | None = None
    if mode_source is not None and os.name != "nt":
        orig_mode = os.stat(mode_source).st_mode & 0o777

    fd, tmp_name = tempfile.mkstemp(prefix=".recalc_", suffix=dst.suffix or ".xlsx", dir=dst.parent)
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        # copyfile copies data only, keeping mkstemp's 0o600 permissions until explicit chmod
        shutil.copyfile(src, tmp_path)
        if orig_mode is not None:
            os.chmod(tmp_path, orig_mode)
        os.replace(tmp_path, dst)
    except Exception:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
        raise


def _kill_process_group(proc: subprocess.Popen) -> None:
    """Terminate a spawned LibreOffice process and all of its children."""
    if os.name == "nt":
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                capture_output=True,
                check=False,
            )
        except OSError:
            proc.kill()
    else:
        try:
            os.killpg(proc.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        except OSError:
            try:
                proc.kill()
            except OSError:
                pass


def run_libreoffice_convert(
    soffice: str,
    xlsx_path: Path,
    outdir: Path,
    profile_dir: Path,
    timeout: int,
) -> tuple[bool, Path | None, str | None]:
    """Run headless LibreOffice conversion with forced recalculation and process-group cleanup."""
    write_libreoffice_recalc_profile(profile_dir)
    profile_uri = profile_dir.as_uri()
    cmd = [
        soffice,
        f"-env:UserInstallation={profile_uri}",
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--convert-to",
        "xlsx",
        "--outdir",
        str(outdir),
        str(xlsx_path),
    ]

    popen_kwargs: dict = {
        "stdout": subprocess.PIPE,
        "stderr": subprocess.PIPE,
        "text": True,
    }
    if os.name == "nt":
        popen_kwargs["creationflags"] = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    else:
        popen_kwargs["start_new_session"] = True

    try:
        proc = subprocess.Popen(cmd, **popen_kwargs)
    except OSError as exc:
        return False, None, f"Failed to execute LibreOffice: {exc}"

    try:
        stdout, stderr = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        _kill_process_group(proc)
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            pass
        for stream in (proc.stdout, proc.stderr):
            if stream is not None:
                try:
                    stream.close()
                except OSError:
                    pass
        return False, None, f"LibreOffice timed out after {timeout}s."
    except Exception:
        if proc.poll() is None:
            _kill_process_group(proc)
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                pass
        raise

    recalculated = outdir / f"{xlsx_path.stem}.xlsx"
    if proc.returncode != 0 or not recalculated.exists():
        detail = (stderr or stdout or "").strip()
        return (
            False,
            None,
            f"LibreOffice conversion failed (rc={proc.returncode}): {detail}",
        )

    return True, recalculated, None


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
    preflight_err = preflight_validate_input(xlsx_path)
    if preflight_err is not None:
        print(
            json.dumps(
                {
                    "status": "error",
                    "file": str(xlsx_path),
                    "recalculated_with_libreoffice": False,
                    "message": preflight_err,
                },
                indent=2,
            )
        )
        return 1

    # Determine output destination and enforce sibling protection before running LibreOffice
    dest_path: Path | None
    if args.check_only:
        dest_path = None
    elif args.output:
        dest_path = Path(args.output).resolve()
        if not dest_path.parent.exists() or not dest_path.parent.is_dir():
            print(
                json.dumps(
                    {
                        "status": "error",
                        "file": str(dest_path),
                        "recalculated_with_libreoffice": False,
                        "message": f"Output directory does not exist: {dest_path.parent}",
                    },
                    indent=2,
                )
            )
            return 1
    elif xlsx_path.suffix.lower() == ".xlsx":
        dest_path = xlsx_path
    else:
        sibling = xlsx_path.with_suffix(".xlsx")
        if sibling.exists():
            print(
                json.dumps(
                    {
                        "status": "error",
                        "file": str(sibling),
                        "source_file": str(xlsx_path),
                        "recalculated_with_libreoffice": False,
                        "message": (
                            f"Refusing to overwrite existing sibling workbook '{sibling}' when "
                            f"converting '{xlsx_path.name}'. Pass --output <path> or --check-only."
                        ),
                    },
                    indent=2,
                )
            )
            return 1
        dest_path = sibling

    soffice = find_soffice()
    recalculated = False
    note: str | None = None

    with tempfile.TemporaryDirectory(prefix="xlsx_recalc_") as tmpdir:
        tmp_path = Path(tmpdir)
        if os.name != "nt":
            os.chmod(tmp_path, 0o700)

        if soffice is not None:
            ok, lo_output, lo_err = run_libreoffice_convert(
                soffice=soffice,
                xlsx_path=xlsx_path,
                outdir=tmp_path,
                profile_dir=tmp_path / "lo_profile",
                timeout=args.timeout,
            )
            if not ok or lo_output is None:
                print(
                    json.dumps(
                        {
                            "status": "error",
                            "file": str(dest_path or xlsx_path),
                            "recalculated_with_libreoffice": False,
                            "message": lo_err or "LibreOffice recalculation failed.",
                        },
                        indent=2,
                    )
                )
                return 1

            recalculated = True
            if args.check_only:
                scan_target = lo_output
                reported_file = xlsx_path
            else:
                assert dest_path is not None
                try:
                    atomic_copy(lo_output, dest_path, preserve_mode_from=xlsx_path)
                except OSError as exc:
                    print(
                        json.dumps(
                            {
                                "status": "error",
                                "file": str(dest_path),
                                "recalculated_with_libreoffice": False,
                                "message": f"Failed to write recalculated workbook to '{dest_path}': {exc}",
                            },
                            indent=2,
                        )
                    )
                    return 1
                scan_target = dest_path
                reported_file = dest_path
                if dest_path != xlsx_path and not args.output:
                    note = (
                        f"Converted non-.xlsx input '{xlsx_path.name}' to '{dest_path.name}' "
                        "without overwriting the original file."
                    )
        else:
            note = "LibreOffice (soffice) not found on PATH; scanning existing cached values only."
            if xlsx_path.suffix.lower() not in OOXML_ZIP_SUFFIXES:
                print(
                    json.dumps(
                        {
                            "status": "error",
                            "file": str(xlsx_path),
                            "recalculated_with_libreoffice": False,
                            "message": (
                                f"Cannot scan '{xlsx_path.suffix}' directly with openpyxl when "
                                "LibreOffice is unavailable. Convert to .xlsx first."
                            ),
                        },
                        indent=2,
                    )
                )
                return 1
            scan_target = xlsx_path
            reported_file = xlsx_path

        try:
            scan = scan_workbook(scan_target)
        except WorkbookLoadError as exc:
            print(
                json.dumps(
                    {
                        "status": "error",
                        "file": str(reported_file),
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

    result: dict = {
        "status": status,
        "file": str(reported_file),
        "recalculated_with_libreoffice": recalculated,
        "total_errors": scan["total_errors"],
        "total_formulas": scan["total_formulas"],
        "uncached_formulas": scan["uncached_formulas"],
        "error_summary": scan["error_summary"],
    }
    if args.check_only:
        result["check_only"] = True
    if reported_file != xlsx_path:
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
