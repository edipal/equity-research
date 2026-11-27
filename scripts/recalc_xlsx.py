#!/usr/bin/env python3
"""Recalculate Excel workbook formulas via LibreOffice and scan for formula errors.

Usage:
    python3 scripts/recalc_xlsx.py <path-to-workbook.xlsx> [timeout_seconds]

Outputs a JSON report to stdout and exits 0 if zero formula errors are found,
or non-zero if formula errors or execution failures occur.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

EXCEL_ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")


def find_soffice() -> str | None:
    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    return None


def recalculate_with_libreoffice(xlsx_path: Path, timeout: int) -> tuple[bool, str | None]:
    soffice = find_soffice()
    if not soffice:
        return False, "LibreOffice (soffice) not found on PATH; scanning existing cached values only."

    with tempfile.TemporaryDirectory(prefix="xlsx_recalc_") as tmpdir:
        cmd = [
            soffice,
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
            return False, f"LibreOffice timed out after {timeout}s."

        recalculated = Path(tmpdir) / xlsx_path.name
        if proc.returncode != 0 or not recalculated.exists():
            stderr = (proc.stderr or proc.stdout or "").strip()
            return False, f"LibreOffice conversion failed (rc={proc.returncode}): {stderr}"

        shutil.copy2(recalculated, xlsx_path)
        return True, None


def scan_workbook(xlsx_path: Path) -> dict:
    try:
        import openpyxl
    except ImportError as exc:
        raise SystemExit(
            json.dumps(
                {
                    "status": "error",
                    "message": f"openpyxl is required: {exc}. Install with `pip install -r requirements.txt`.",
                },
                indent=2,
            )
        )

    wb_formulas = openpyxl.load_workbook(xlsx_path, data_only=False)
    wb_values = openpyxl.load_workbook(xlsx_path, data_only=True)

    total_formulas = 0
    error_summary: dict[str, dict] = {}
    total_errors = 0

    for sheet_name in wb_formulas.sheetnames:
        ws_f = wb_formulas[sheet_name]
        ws_v = wb_values[sheet_name]

        for row in ws_f.iter_rows():
            for cell_f in row:
                val_f = cell_f.value
                val_v = ws_v[cell_f.coordinate].value

                if isinstance(val_f, str) and val_f.startswith("="):
                    total_formulas += 1

                for candidate in (val_v, val_f):
                    if isinstance(candidate, str) and candidate in EXCEL_ERRORS:
                        total_errors += 1
                        bucket = error_summary.setdefault(
                            candidate, {"count": 0, "locations": []}
                        )
                        bucket["count"] += 1
                        bucket["locations"].append(f"{sheet_name}!{cell_f.coordinate}")
                        break

    return {
        "total_formulas": total_formulas,
        "total_errors": total_errors,
        "error_summary": error_summary,
    }


def main() -> int:
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print(__doc__.strip())
        return 0 if len(sys.argv) >= 2 else 2

    xlsx_path = Path(sys.argv[1]).resolve()
    timeout = int(sys.argv[2]) if len(sys.argv) >= 3 else 30

    if not xlsx_path.exists():
        print(json.dumps({"status": "error", "message": f"File not found: {xlsx_path}"}, indent=2))
        return 1

    recalculated, warning = recalculate_with_libreoffice(xlsx_path, timeout)
    scan = scan_workbook(xlsx_path)

    result = {
        "status": "success" if scan["total_errors"] == 0 else "errors_found",
        "file": str(xlsx_path),
        "recalculated_with_libreoffice": recalculated,
        "total_errors": scan["total_errors"],
        "total_formulas": scan["total_formulas"],
        "error_summary": scan["error_summary"],
    }
    if warning:
        result["warning"] = warning

    print(json.dumps(result, indent=2))
    return 0 if scan["total_errors"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
