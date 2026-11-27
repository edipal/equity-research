#!/usr/bin/env python3
# Independent implementation for Google Antigravity, licensed under the Apache License, Version 2.0.
"""Recalculate Excel workbook formulas via LibreOffice and scan for formula errors.

Usage:
    python3 scripts/recalc_xlsx.py <path-to-workbook.xlsx> [timeout_seconds]
        [--allow-unrecalculated] [--check-only | --in-place | --output PATH]

By default (and with --check-only), recalculates in an isolated temporary directory
and scans for formula errors without modifying the input file on disk. Pass --in-place
to atomically replace an .xlsx file with its LibreOffice-recalculated copy, or
--output PATH to write the recalculated .xlsx to a separate path.

Outputs a JSON report to stdout and exits:
    0 - status "success" (zero errors, all formula cells verified or --allow-unrecalculated)
    1 - status "errors_found" or "error" (formula errors, missing/corrupt file, LibreOffice failure/timeout)
    2 - status "unverified" or CLI usage error (uncached formulas without LibreOffice recalculation)
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

_SCRIPTS_DIR = str(Path(__file__).resolve().parent)
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from xlsx_scan import (  # noqa: E402
    OOXML_ZIP_SUFFIXES,
    WorkbookLoadError,
    inspect_workbook_xml,
    scan_workbook_formulas,
)

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
        help=(
            "Explicit alias for the default non-destructive mode: recalculate in a temporary "
            "directory and scan for errors without modifying or writing any file."
        ),
    )
    mode_group.add_argument(
        "--in-place",
        action="store_true",
        help=(
            "Atomically overwrite the input .xlsx file with the LibreOffice-recalculated copy "
            "(refuses non-.xlsx files, What-If data tables, and workbooks containing Excel features "
            "that LibreOffice re-export can drop or alter)."
        ),
    )
    mode_group.add_argument(
        "--output",
        metavar="PATH",
        help="Write the recalculated .xlsx workbook to PATH without modifying the input file.",
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


def _is_same_file(a: Path, b: Path) -> bool:
    """Return True if two paths refer to the same file (handles case-insensitive FS & hardlinks)."""
    if a == b:
        return True
    try:
        return a.exists() and b.exists() and os.path.samefile(a, b)
    except OSError:
        return False


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


def check_openpyxl_available() -> str | None:
    """Return an actionable error message if openpyxl is not importable, else None."""
    try:
        import openpyxl  # noqa: F401
    except ImportError as exc:
        return f"openpyxl is required: {exc}. Install with `pip install -r requirements.txt`."
    return None


def _fail(
    file_path: Path,
    message: str,
    *,
    source_file: Path | None = None,
    recalculated: bool = False,
    extra: dict | None = None,
    exit_code: int = 1,
) -> int:
    """Emit a structured JSON error payload to stdout and return the process exit code."""
    payload: dict = {
        "status": "error",
        "file": str(file_path),
    }
    if source_file is not None:
        payload["source_file"] = str(source_file)
    payload["recalculated_with_libreoffice"] = recalculated
    if extra:
        payload.update(extra)
    payload["message"] = message
    print(json.dumps(payload, indent=2))
    return exit_code


def scan_workbook(xlsx_path: Path) -> dict:
    return scan_workbook_formulas(xlsx_path)


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    xlsx_path = Path(args.path).resolve()

    # Check openpyxl upfront before spawning LibreOffice or performing any file write
    openpyxl_err = check_openpyxl_available()
    if openpyxl_err is not None:
        return _fail(xlsx_path, openpyxl_err)

    preflight_err = preflight_validate_input(xlsx_path)
    if preflight_err is not None:
        return _fail(xlsx_path, preflight_err)

    # Determine execution mode: default (and explicit --check-only) is non-destructive
    if args.in_place:
        mode = "in_place"
    elif args.output:
        mode = "output"
    else:
        mode = "check_only"
    is_check_only = mode == "check_only"

    # Validate --in-place extension before invoking LibreOffice
    if args.in_place and xlsx_path.suffix.lower() != ".xlsx":
        return _fail(
            xlsx_path,
            (
                f"Refusing --in-place on '{xlsx_path.name}' (extension '{xlsx_path.suffix or '(none)'}'): "
                "LibreOffice --convert-to xlsx produces macro-free .xlsx bytes. "
                "Use the default non-destructive check (--check-only) for read-only auditing, "
                "or pass --output <new.xlsx> to write a converted .xlsx copy."
            ),
        )

    # Validate --output path (extension and parent directory) before invoking LibreOffice
    output_path: Path | None = None
    if args.output:
        output_path = Path(args.output).resolve()
        if output_path.suffix.lower() != ".xlsx":
            return _fail(
                output_path,
                (
                    f"Invalid --output extension '{output_path.suffix or '(none)'}': "
                    "LibreOffice --convert-to xlsx produces macro-free .xlsx bytes, "
                    "so --output must end in '.xlsx'."
                ),
                source_file=xlsx_path,
            )
        if not output_path.parent.exists() or not output_path.parent.is_dir():
            return _fail(
                output_path,
                f"Output directory does not exist: {output_path.parent}",
                source_file=xlsx_path,
            )

    # Detect LibreOffice and enforce its presence for write modes (--in-place / --output)
    soffice = find_soffice()
    if soffice is None and args.in_place:
        return _fail(
            xlsx_path,
            (
                "Cannot update workbook --in-place because "
                "LibreOffice (soffice) is not installed on PATH."
            ),
        )
    if soffice is None and args.output:
        assert output_path is not None
        return _fail(
            output_path,
            (
                "Cannot write recalculated --output workbook because "
                "LibreOffice (soffice) is not installed on PATH."
            ),
            source_file=xlsx_path,
        )

    dest_path: Path | None
    if is_check_only:
        dest_path = None
    elif output_path is not None:
        dest_path = output_path
    else:
        dest_path = xlsx_path

    is_in_place = dest_path is not None and _is_same_file(dest_path, xlsx_path)
    err_source_file = xlsx_path if args.output else None

    # Single streaming OOXML pass to (1) detect Excel What-If data tables (<f t="dataTable">)
    # and LibreOffice-lossy Excel features before any in-place overwrite, and (2) count on-disk
    # uncached formulas without re-opening the input workbook twice in openpyxl after LibreOffice
    # runs in check-only mode. Fail closed on corrupt OOXML zip/XML when an in-place overwrite
    # is requested; in non-destructive read-only or separate --output modes, warn and proceed.
    warnings: list[str] = []
    data_table_locs: list[str] = []
    lossy_features: list[str] = []
    pre_scan_uncached: int | None = None
    try:
        xml_info = inspect_workbook_xml(xlsx_path)
        data_table_locs = xml_info["data_tables"]
        pre_scan_uncached = xml_info["uncached_formulas"]
        lossy_features = xml_info.get("lossy_features", [])
    except WorkbookLoadError as exc:
        if is_in_place:
            return _fail(xlsx_path, str(exc), source_file=err_source_file)
        warnings.append(
            f"Could not inspect input XML for Excel What-If data tables or uncached formulas: {exc}"
        )

    if data_table_locs:
        locs_preview = ", ".join(data_table_locs[:5])
        if len(data_table_locs) > 5:
            locs_preview += f" (+{len(data_table_locs) - 5} more)"
        if is_in_place:
            extra_dt: dict = {"data_tables": data_table_locs}
            if lossy_features:
                extra_dt["lossy_features"] = lossy_features
            return _fail(
                xlsx_path,
                (
                    f"Refusing to overwrite '{xlsx_path.name}' in place because it contains "
                    f"Excel What-If data table(s) ({locs_preview}). LibreOffice re-export "
                    "converts Excel What-If data tables into incompatible TABLE(...) formulas. "
                    "Use the default non-destructive check (--check-only) for safe read-only "
                    "auditing, or pass --output <new.xlsx> if a re-exported copy is explicitly desired."
                ),
                source_file=err_source_file,
                extra=extra_dt,
            )
        if args.output:
            warnings.append(
                f"Input workbook contains Excel What-If data table(s) ({locs_preview}); "
                "LibreOffice re-export converts What-If data tables into Calc TABLE(...) formulas "
                "in the --output copy (the original input file was left untouched)."
            )
        elif soffice is not None and is_check_only:
            warnings.append(
                f"Input workbook contains Excel What-If data table(s) ({locs_preview}); "
                "audited in a temporary copy without modifying the input file."
            )

    if lossy_features:
        features_str = ", ".join(lossy_features)
        if is_in_place:
            return _fail(
                xlsx_path,
                (
                    f"Refusing to overwrite '{xlsx_path.name}' in place because it contains "
                    f"Excel feature(s) that LibreOffice re-export can drop or alter ({features_str}). "
                    "Use the default non-destructive check (--check-only) for safe read-only "
                    "auditing, or pass --output <new.xlsx> to preserve the original file."
                ),
                source_file=err_source_file,
                extra={"lossy_features": lossy_features},
            )
        if args.output:
            warnings.append(
                f"Input workbook contains Excel feature(s) that LibreOffice re-export may drop or "
                f"alter ({features_str}) in the --output copy (the original input file was left untouched)."
            )
        elif soffice is not None and is_check_only:
            warnings.append(
                f"Input workbook contains Excel feature(s) ({features_str}) that LibreOffice "
                "re-export can drop or alter; audited in a temporary copy without modifying the input file."
            )

    recalculated = False
    modified_file = False

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
                return _fail(
                    dest_path or xlsx_path,
                    lo_err or "LibreOffice recalculation failed.",
                    source_file=err_source_file,
                )

            recalculated = True
            if is_check_only:
                scan_target = lo_output
                reported_file = xlsx_path
            else:
                assert dest_path is not None
                try:
                    atomic_copy(lo_output, dest_path, preserve_mode_from=xlsx_path)
                except OSError as exc:
                    return _fail(
                        dest_path,
                        f"Failed to write recalculated workbook to '{dest_path}': {exc}",
                        source_file=err_source_file,
                    )
                modified_file = True
                scan_target = dest_path
                reported_file = dest_path
        else:
            warnings.append("LibreOffice (soffice) not found on PATH; scanning existing cached values only.")
            if xlsx_path.suffix.lower() not in OOXML_ZIP_SUFFIXES:
                return _fail(
                    xlsx_path,
                    (
                        f"Cannot scan '{xlsx_path.suffix}' directly with openpyxl when "
                        "LibreOffice is unavailable. Convert to .xlsx first."
                    ),
                )
            scan_target = xlsx_path
            reported_file = xlsx_path

        try:
            scan = scan_workbook(scan_target)
        except WorkbookLoadError as exc:
            return _fail(
                reported_file,
                str(exc),
                source_file=err_source_file,
                recalculated=recalculated,
            )

    on_disk_uncached: int | None
    if is_check_only and recalculated:
        on_disk_uncached = pre_scan_uncached
    else:
        on_disk_uncached = scan["uncached_formulas"]

    if is_check_only and recalculated and on_disk_uncached is not None and on_disk_uncached > 0:
        warnings.append(
            f"Input file on disk still has {on_disk_uncached} uncached formula(s) "
            "(audited in temporary copy only); run with --in-place or --output <recalc.xlsx> "
            "before reading with openpyxl(data_only=True) or validate_dcf.py."
        )

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
        "mode": mode,
        "modified_file": modified_file,
        "recalculated_with_libreoffice": recalculated,
        "total_errors": scan["total_errors"],
        "total_formulas": scan["total_formulas"],
        "uncached_formulas": scan["uncached_formulas"],
        "on_disk_uncached_formulas": on_disk_uncached,
        "data_table_formulas": max(scan["data_table_formulas"], len(data_table_locs)),
        "error_summary": scan["error_summary"],
    }
    if data_table_locs:
        result["data_tables"] = data_table_locs
    if lossy_features:
        result["lossy_features"] = lossy_features
    if is_check_only:
        result["check_only"] = True
    if not _is_same_file(reported_file, xlsx_path):
        result["source_file"] = str(xlsx_path)
    if warnings:
        result["warning"] = " ".join(warnings)
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
