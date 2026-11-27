#!/usr/bin/env python3
"""Regression test suite for Antigravity equity research harness scripts.

Run with:
    python3 -m unittest discover -s tests -v
"""

from __future__ import annotations

import importlib.util
import io
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

try:
    import openpyxl
except ImportError:
    openpyxl = None

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = ROOT / "scripts"
VALIDATE_HARNESS_PATH = SCRIPTS_DIR / "validate_harness.py"
RECALC_XLSX_PATH = SCRIPTS_DIR / "recalc_xlsx.py"
VALIDATE_DCF_PATH = ROOT / ".agents" / "skills" / "dcf-model" / "scripts" / "validate_dcf.py"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


validate_harness_mod = load_module("validate_harness", VALIDATE_HARNESS_PATH)
recalc_xlsx_mod = load_module("recalc_xlsx", RECALC_XLSX_PATH)
validate_dcf_mod = load_module("validate_dcf", VALIDATE_DCF_PATH)


def create_minimal_valid_harness(base: Path) -> None:
    """Scaffold a minimal valid Antigravity harness in a temp directory."""
    (base / "AGENTS.md").write_text("# Workspace Guide\n\nSee [docs](docs/guide.md).\n", encoding="utf-8")
    (base / "README.md").write_text("# Readme\n", encoding="utf-8")
    (base / "docs").mkdir(parents=True, exist_ok=True)
    (base / "docs" / "guide.md").write_text("# Guide\n", encoding="utf-8")

    rules_dir = base / ".agents" / "rules"
    rules_dir.mkdir(parents=True, exist_ok=True)
    (rules_dir / "integrity.md").write_text(
        "---\ntrigger: always_on\ndescription: Core rules\n---\n\n# Integrity\n",
        encoding="utf-8",
    )

    skill_dir = base / ".agents" / "skills" / "sample-skill"
    skill_dir.mkdir(parents=True, exist_ok=True)
    (skill_dir / "SKILL.md").write_text(
        "---\nname: sample-skill\ndescription: A valid sample skill.\n---\n\n# Sample Skill\n",
        encoding="utf-8",
    )

    agents_dir = base / ".agents" / "agents"
    agents_dir.mkdir(parents=True, exist_ok=True)
    (agents_dir / "sample-agent.md").write_text(
        "---\nname: sample-agent\ndescription: Sample agent.\nmainAgent: true\n"
        "subagent: true\nmodel: inherit\ncommandExecutionPolicy: sandbox\n---\n\n# Prompt\n",
        encoding="utf-8",
    )

    config_dir = base / "config"
    config_dir.mkdir(parents=True, exist_ok=True)
    (config_dir / "mcp_config.example.json").write_text(
        json.dumps({"mcpServers": {"vendor": {"serverUrl": "https://mcp.example.com"}}}),
        encoding="utf-8",
    )


class TestValidateHarness(unittest.TestCase):
    def test_repo_passes_cleanly(self) -> None:
        errors, warnings, counts = validate_harness_mod.validate_harness(ROOT)
        self.assertEqual(errors, [], f"Unexpected harness errors: {errors}")
        self.assertEqual(counts["rules"], 3)
        self.assertEqual(counts["skills"], 18)
        self.assertEqual(counts["agents"], 4)

    def test_cli_exits_zero_on_repo(self) -> None:
        proc = subprocess.run(
            [sys.executable, str(VALIDATE_HARNESS_PATH)],
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("PASSED with 0 errors", proc.stdout)

    def test_rejects_leading_digit_skill_name(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            bad_skill = base / ".agents" / "skills" / "3-statement-model"
            bad_skill.mkdir(parents=True)
            (bad_skill / "SKILL.md").write_text(
                "---\nname: 3-statement-model\ndescription: Bad leading digit.\n---\n\n# Body\n",
                encoding="utf-8",
            )
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(
                any("invalid skill name '3-statement-model'" in e for e in errors),
                f"Expected leading-digit skill rejection, got: {errors}",
            )

    def test_rejects_stale_non_antigravity_patterns(self) -> None:
        stale_samples = [
            "Use ask_user_question to confirm.",
            "Open the file with Read tool.",
            "Generate the report via DOCX skill.",
            "Build tables using XLSX skill.",
            "This works in the PowerPoint add-in.",
            "Audit the currently selected range.",
            "Clean the active sheet.",
        ]
        for sample in stale_samples:
            with self.subTest(sample=sample), tempfile.TemporaryDirectory() as tmp:
                base = Path(tmp)
                create_minimal_valid_harness(base)
                skill_md = base / ".agents" / "skills" / "sample-skill" / "SKILL.md"
                skill_md.write_text(
                    f"---\nname: sample-skill\ndescription: Valid.\n---\n\n{sample}\n",
                    encoding="utf-8",
                )
                errors, _, _ = validate_harness_mod.validate_harness(base)
                self.assertTrue(
                    any("stale non-Antigravity reference" in e for e in errors),
                    f"Expected stale pattern error for {sample!r}, got: {errors}",
                )

    def test_rejects_unsupported_agent_keys_and_invalid_model(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            agent_md = base / ".agents" / "agents" / "sample-agent.md"
            agent_md.write_text(
                "---\nname: sample-agent\ndescription: Sample.\nmodel: flash_lite\n"
                "inheritMcp: true\n---\n\n# Prompt\n",
                encoding="utf-8",
            )
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(any("unsupported custom agent frontmatter key" in e for e in errors), errors)
            self.assertTrue(any("invalid model 'flash_lite'" in e for e in errors), errors)

    def test_rejects_broken_relative_markdown_link(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            (base / "README.md").write_text("See [missing](docs/does-not-exist.md).\n", encoding="utf-8")
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(any("broken relative link" in e for e in errors), errors)

    def test_rejects_mcp_server_missing_endpoint(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            (base / "config" / "mcp_config.example.json").write_text(
                json.dumps({"mcpServers": {"broken": {"headers": {"X-Key": "123"}}}}),
                encoding="utf-8",
            )
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(any("must define 'serverUrl', 'url', or 'command'" in e for e in errors), errors)

    def test_rejects_unquoted_colon_yaml_frontmatter(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            skill_md = base / ".agents" / "skills" / "sample-skill" / "SKILL.md"
            skill_md.write_text(
                "---\nname: sample-skill\ndescription: Bad: unquoted colon\n---\n\n# Body\n",
                encoding="utf-8",
            )
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(len(errors) > 0, "Expected YAML parse error for unquoted colon")


@unittest.skipIf(openpyxl is None, "openpyxl is required for Excel script tests")
class TestRecalcXlsx(unittest.TestCase):
    def _run_main(self, argv: list[str]) -> tuple[int, dict]:
        buf = io.StringIO()
        with redirect_stdout(buf):
            try:
                rc = recalc_xlsx_mod.main(argv)
            except SystemExit as exc:
                rc = int(exc.code) if isinstance(exc.code, int) else 1
        output = buf.getvalue().strip()
        payload = json.loads(output)
        return rc, payload

    def test_invalid_timeout_emits_json_error_without_traceback(self) -> None:
        proc = subprocess.run(
            [sys.executable, str(RECALC_XLSX_PATH), "dummy.xlsx", "notanint"],
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertNotEqual(proc.returncode, 0)
        self.assertNotIn("Traceback", proc.stdout + proc.stderr)
        payload = json.loads(proc.stdout)
        self.assertEqual(payload["status"], "error")
        self.assertIn("timeout", payload["message"])

    def test_detects_cached_excel_errors(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "cached_errors.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Model"
            ws["A1"] = "#DIV/0!"
            ws["B2"] = "#REF!"
            wb.save(path)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc, payload = self._run_main([str(path)])

            self.assertEqual(rc, 1)
            self.assertEqual(payload["status"], "errors_found")
            self.assertEqual(payload["total_errors"], 2)
            self.assertIn("#DIV/0!", payload["error_summary"])
            self.assertIn("#REF!", payload["error_summary"])

    def test_detects_embedded_error_token_in_formula_string(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "embedded_ref.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "DCF"
            ws["A1"] = 100
            ws["A2"] = "=SUM(A1,#REF!)"
            wb.save(path)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc, payload = self._run_main([str(path)])

            self.assertEqual(rc, 1)
            self.assertEqual(payload["status"], "errors_found")
            self.assertEqual(payload["total_errors"], 1)
            self.assertEqual(payload["error_summary"]["#REF!"]["locations"], ["DCF!A2"])

    def test_prose_note_mentioning_ref_does_not_false_positive(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "prose_note.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Notes"
            ws["A1"] = "Audit reminder: ensure there are no #REF! or #DIV/0! errors in DCF."
            ws["A2"] = 42.5
            wb.save(path)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc, payload = self._run_main([str(path)])

            self.assertEqual(rc, 0)
            self.assertEqual(payload["status"], "success")
            self.assertEqual(payload["total_errors"], 0)
            self.assertEqual(payload["uncached_formulas"], 0)

    def test_uncached_formulas_without_libreoffice_fail_unverified_unless_allowed(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "uncached.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Model"
            ws["A1"] = 10
            ws["A2"] = "=A1*2"
            wb.save(path)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_strict, payload_strict = self._run_main([str(path)])
                rc_allow, payload_allow = self._run_main([str(path), "--allow-unrecalculated"])

            self.assertEqual(rc_strict, 2)
            self.assertEqual(payload_strict["status"], "unverified")
            self.assertEqual(payload_strict["total_formulas"], 1)
            self.assertEqual(payload_strict["uncached_formulas"], 1)

            self.assertEqual(rc_allow, 0)
            self.assertEqual(payload_allow["status"], "success")
            self.assertEqual(payload_allow["uncached_formulas"], 1)

    def test_libreoffice_failure_or_timeout_exits_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "timeout.xlsx"
            wb = openpyxl.Workbook()
            wb.save(path)

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(
                    recalc_xlsx_mod.subprocess,
                    "run",
                    side_effect=subprocess.TimeoutExpired(cmd="soffice", timeout=5),
                ),
            ):
                rc, payload = self._run_main([str(path), "5", "--allow-unrecalculated"])

            self.assertEqual(rc, 1)
            self.assertEqual(payload["status"], "error")
            self.assertIn("timed out", payload["message"])

    def test_libreoffice_uses_isolated_profile_and_handles_non_xlsx_conversion(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            src_path = Path(tmp) / "legacy_model.xlsm"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = 100
            wb.save(src_path)

            captured_cmd: list[str] = []

            def fake_run(cmd, **kwargs):
                captured_cmd.extend(cmd)
                outdir = Path(cmd[cmd.index("--outdir") + 1])
                converted = outdir / f"{src_path.stem}.xlsx"
                shutil.copy2(src_path, converted)
                return subprocess.CompletedProcess(cmd, 0, "", "")

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "run", side_effect=fake_run),
            ):
                rc, payload = self._run_main([str(src_path)])

            self.assertEqual(rc, 0)
            self.assertEqual(payload["status"], "success")
            self.assertTrue(payload["recalculated_with_libreoffice"])
            self.assertTrue(
                any(arg.startswith("-env:UserInstallation=file://") and arg.endswith("/lo_profile") for arg in captured_cmd),
                f"Expected isolated LibreOffice profile flag in {captured_cmd}",
            )
            converted_target = src_path.with_suffix(".xlsx")
            self.assertEqual(Path(payload["file"]), converted_target.resolve())
            self.assertTrue(converted_target.exists())


@unittest.skipIf(openpyxl is None, "openpyxl is required for DCF validation tests")
class TestValidateDcf(unittest.TestCase):
    def _build_compliant_two_sheet_dcf(self, path: Path, with_formula: bool = False) -> None:
        wb = openpyxl.Workbook()
        ws_dcf = wb.active
        ws_dcf.title = "DCF"
        ws_dcf["A1"] = "DCF Valuation Model"
        ws_dcf["A3"] = "WACC"
        ws_dcf["B3"] = 0.09
        ws_dcf["A4"] = "Terminal Growth Rate"
        ws_dcf["B4"] = 0.025
        ws_dcf["A6"] = "PV of Terminal Value"
        ws_dcf["B6"] = 600.0
        ws_dcf["A7"] = "Enterprise Value"
        ws_dcf["B7"] = 1000.0
        ws_dcf["A10"] = "Sensitivity Analysis (WACC vs Terminal Growth)"
        ws_dcf["A11"] = "Note: verify zero #REF! or #DIV/0! errors before delivery."
        if with_formula:
            ws_dcf["B8"] = "=B6/B7"

        ws_wacc = wb.create_sheet("WACC")
        ws_wacc["A1"] = "WACC Build"
        ws_wacc["A3"] = "WACC"
        ws_wacc["B3"] = 0.09
        wb.save(path)

    def test_compliant_two_sheet_dcf_passes_with_zero_warnings_and_no_prose_false_positive(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "clean_dcf.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=False)

            results = validate_dcf_mod.validate_dcf_model(str(path))

            self.assertEqual(results["status"], "PASS", results)
            self.assertEqual(results["error_count"], 0, results)
            self.assertEqual(results["warning_count"], 0, f"Unexpected warnings: {results['warnings']}")
            self.assertTrue(any("WACC (9.00%) in reasonable range" in i for i in results["info"]))
            self.assertTrue(any("inline sensitivity" in i for i in results["info"]))

    def test_wacc_range_fallback_to_dcf_sheet_without_attribute_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "no_wacc_sheet.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A2"] = "WACC"
            ws_dcf["B2"] = 0.10
            ws_dcf["A3"] = "Terminal Growth"
            ws_dcf["B3"] = 0.03
            ws_dcf["A4"] = "PV of Terminal Value"
            ws_dcf["B4"] = 600.0
            ws_dcf["A5"] = "Enterprise Value"
            ws_dcf["B5"] = 1000.0
            ws_dcf["A8"] = "Sensitivity Analysis"
            wb.save(path)

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "PASS", results)
            self.assertFalse(
                any("AttributeError" in w or "has no attribute" in w for w in results["warnings"]),
                results["warnings"],
            )
            self.assertTrue(any("WACC (10.00%) in reasonable range" in i for i in results["info"]))

    def test_uncached_formulas_cause_fail(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "uncached_dcf.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=True)

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "FAIL", results)
            self.assertTrue(
                any("Uncached formula values detected" in e for e in results["errors"]),
                results["errors"],
            )

    def test_catches_real_formula_errors_with_accurate_error_count(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "broken_dcf.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=False)
            wb = openpyxl.load_workbook(path)
            ws = wb["DCF"]
            ws["C5"] = "#DIV/0!"
            ws["C6"] = "=SUM(B6,#REF!)"
            wb.save(path)

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "FAIL", results)
            # 2 formula/cached errors + 1 uncached formula error for C6 = 3 exact errors
            # (summary "Total formula errors: 2" must be in info, NOT inflating self.errors)
            self.assertEqual(results["error_count"], 3, results["errors"])
            self.assertTrue(any("#DIV/0! at DCF!C5" in e for e in results["errors"]))
            self.assertTrue(any("#REF! at DCF!C6" in e for e in results["errors"]))
            self.assertTrue(any("Total formula errors: 2" in i for i in results["info"]))
            self.assertFalse(any("Total formula errors:" in e for e in results["errors"]))


if __name__ == "__main__":
    unittest.main()
