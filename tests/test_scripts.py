#!/usr/bin/env python3
"""Regression test suite for Antigravity equity research harness scripts.

Run with:
    python3 -m unittest discover -s tests -v
"""

from __future__ import annotations

import importlib.util
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
import zipfile
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
BUILD_DCF_SKELETON_PATH = ROOT / ".agents" / "skills" / "dcf-model" / "scripts" / "build_dcf_skeleton.py"


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


def _rewrite_formula_cell_as_cached_empty_string(
    xlsx_path: Path, cell_ref: str, sheet_xml: str = "xl/worksheets/sheet1.xml"
) -> None:
    """Rewrite an openpyxl-saved formula cell to Excel/LibreOffice cached empty-string XML (<c t="str"><f>...</f><v></v></c>)."""
    import xml.etree.ElementTree as ET

    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", ns)

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}

    root = ET.fromstring(entries[sheet_xml])
    updated = False
    for cell in root.iter(f"{{{ns}}}c"):
        if cell.attrib.get("r") == cell_ref:
            cell.attrib["t"] = "str"
            v_el = cell.find(f"{{{ns}}}v")
            if v_el is None:
                v_el = ET.SubElement(cell, f"{{{ns}}}v")
            v_el.text = ""
            updated = True
            break
    assert updated, f"Cell {cell_ref} not found in {sheet_xml}"
    entries[sheet_xml] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(xlsx_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)


class TestValidateHarness(unittest.TestCase):
    def test_repo_passes_cleanly(self) -> None:
        errors, warnings, counts = validate_harness_mod.validate_harness(ROOT)
        self.assertEqual(errors, [], f"Unexpected harness errors: {errors}")
        expected_rules = len(list((ROOT / ".agents" / "rules").glob("*.md")))
        expected_skills = len([p for p in (ROOT / ".agents" / "skills").iterdir() if p.is_dir()])
        expected_agents = len(list((ROOT / ".agents" / "agents").glob("*.md")))
        self.assertGreaterEqual(counts["rules"], 1)
        self.assertGreaterEqual(counts["skills"], 1)
        self.assertGreaterEqual(counts["agents"], 1)
        self.assertEqual(counts["rules"], expected_rules)
        self.assertEqual(counts["skills"], expected_skills)
        self.assertEqual(counts["agents"], expected_agents)

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

    def test_validates_backticked_markdown_links_while_ignoring_code_spans(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            # Valid backticked link and links completely inside inline code spans should pass
            (base / "README.md").write_text(
                "Valid: [`docs/guide.md`](docs/guide.md) and [see `guide.md`](docs/guide.md).\n"
                "Ignored code snippets: `[missing](docs/missing.md)` and ``[`missing`](docs/missing.md)``.\n",
                encoding="utf-8",
            )
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertEqual(errors, [], f"Unexpected errors for valid/code-span links: {errors}")

            # Broken link with backticked display text MUST be caught
            (base / "README.md").write_text("Broken: [`missing`](docs/missing.md).\n", encoding="utf-8")
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(
                any("broken relative link 'docs/missing.md'" in e for e in errors),
                f"Expected broken backticked link error, got: {errors}",
            )

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

    def test_validates_markdown_links_in_research_readme(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            research_dir = base / "research"
            research_dir.mkdir(parents=True, exist_ok=True)
            (research_dir / "README.md").write_text(
                "# Research\n\nSee [broken](missing-file.md).\n", encoding="utf-8"
            )
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(
                any("research/README.md" in e and "broken relative link" in e for e in errors),
                f"Expected broken link error in research/README.md, got: {errors}",
            )

    def test_rejects_non_string_trigger_without_type_error(self) -> None:
        bad_triggers = [
            "trigger: [always_on]",
            "trigger:\n  - always_on",
            "trigger: 123",
            "trigger: true",
        ]
        for trigger_yaml in bad_triggers:
            with self.subTest(trigger_yaml=trigger_yaml), tempfile.TemporaryDirectory() as tmp:
                base = Path(tmp)
                create_minimal_valid_harness(base)
                rule_md = base / ".agents" / "rules" / "integrity.md"
                rule_md.write_text(f"---\n{trigger_yaml}\ndescription: Test\n---\n\n# Rule\n", encoding="utf-8")
                errors, _, _ = validate_harness_mod.validate_harness(base)
                self.assertTrue(
                    any("invalid trigger" in e for e in errors),
                    f"Expected invalid trigger error for {trigger_yaml!r}, got: {errors}",
                )

    def test_rejects_unknown_rule_frontmatter_key(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            rule_md = base / ".agents" / "rules" / "integrity.md"
            rule_md.write_text(
                "---\ntrigger: always_on\ndescription: Core\npriority: high\n---\n\n# Integrity\n",
                encoding="utf-8",
            )
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(
                any("unsupported rule frontmatter key(s): priority" in e for e in errors),
                errors,
            )

    def test_rejects_non_bool_agent_flags(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            agent_md = base / ".agents" / "agents" / "sample-agent.md"
            agent_md.write_text(
                "---\nname: sample-agent\ndescription: Sample.\nmainAgent: maybe\n"
                "subagent: 'yes'\nmodel: inherit\ncommandExecutionPolicy: sandbox\n---\n\n# Prompt\n",
                encoding="utf-8",
            )
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(any("'mainAgent' must be a boolean" in e for e in errors), errors)
            self.assertTrue(any("'subagent' must be a boolean" in e for e in errors), errors)

    def test_rejects_non_dict_mcp_root_and_conflicting_server_endpoints(self) -> None:
        for bad_root in ([], "not-an-object", 42):
            with self.subTest(bad_root=bad_root), tempfile.TemporaryDirectory() as tmp:
                base = Path(tmp)
                create_minimal_valid_harness(base)
                (base / "config" / "mcp_config.example.json").write_text(
                    json.dumps(bad_root), encoding="utf-8"
                )
                errors, _, _ = validate_harness_mod.validate_harness(base)
                self.assertTrue(
                    any("must be a JSON object containing a 'mcpServers' object" in e for e in errors),
                    errors,
                )

        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            (base / "config" / "mcp_config.example.json").write_text(
                json.dumps(
                    {
                        "mcpServers": {
                            "dual": {
                                "command": "npx",
                                "serverUrl": "https://mcp.example.com",
                            }
                        }
                    }
                ),
                encoding="utf-8",
            )
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(
                any("must not specify both 'command' and 'serverUrl'/'url'" in e for e in errors),
                errors,
            )

    def test_markdown_link_url_decoding_and_absolute_or_traversal_rejection(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            (base / "docs" / "my guide.md").write_text("# Spaced Guide\n", encoding="utf-8")
            (base / "README.md").write_text("See [spaced](docs/my%20guide.md).\n", encoding="utf-8")
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertEqual(errors, [], f"Expected URL-encoded valid link to pass, got: {errors}")

            (base / "README.md").write_text(
                "Bad abs: [etc](/etc/passwd)\nBad traversal: [escape](../../etc/passwd)\n",
                encoding="utf-8",
            )
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertTrue(any("absolute filesystem path not allowed" in e for e in errors), errors)
            self.assertTrue(any("escapes repository root" in e for e in errors), errors)

    def test_yaml_fallback_strips_inline_comments_consistently(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            rule_md = base / ".agents" / "rules" / "integrity.md"
            rule_md.write_text(
                "---\n"
                "trigger: always_on # load on every session\n"
                'description: "Keep # inside quotes and https://example.com#anchor" # trailing comment\n'
                "---\n\n# Integrity\n",
                encoding="utf-8",
            )
            agent_md = base / ".agents" / "agents" / "sample-agent.md"
            agent_md.write_text(
                "---\n"
                "name: sample-agent # agent id\n"
                "description: See https://example.com/docs#section # inline comment\n"
                "mainAgent: true # primary agent\n"
                "subagent: false # disabled\n"
                "model: inherit # default model\n"
                "commandExecutionPolicy: sandbox # safe\n"
                "---\n\n# Prompt\n",
                encoding="utf-8",
            )
            with mock.patch.object(validate_harness_mod, "yaml", None):
                fm_rule, _, errs_rule = validate_harness_mod.parse_frontmatter(rule_md, root=base)
                fm_agent, _, errs_agent = validate_harness_mod.parse_frontmatter(agent_md, root=base)
                errors, _, _ = validate_harness_mod.validate_harness(base)

            self.assertEqual(errs_rule, [])
            self.assertEqual(fm_rule["trigger"], "always_on")
            self.assertEqual(fm_rule["description"], "Keep # inside quotes and https://example.com#anchor")
            self.assertEqual(errs_agent, [])
            self.assertEqual(fm_agent["name"], "sample-agent")
            self.assertEqual(fm_agent["description"], "See https://example.com/docs#section")
            self.assertIs(fm_agent["mainAgent"], True)
            self.assertIs(fm_agent["subagent"], False)
            self.assertEqual(errors, [], errors)


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

            fake_proc = mock.MagicMock()
            fake_proc.pid = 999999
            fake_proc.communicate.side_effect = subprocess.TimeoutExpired(cmd="soffice", timeout=5)
            fake_proc.wait.return_value = -9

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", return_value=fake_proc),
                mock.patch.object(recalc_xlsx_mod, "_kill_process_group") as mock_killpg,
            ):
                rc, payload = self._run_main([str(path), "5", "--allow-unrecalculated"])

            self.assertEqual(rc, 1)
            self.assertEqual(payload["status"], "error")
            self.assertIn("timed out", payload["message"])
            mock_killpg.assert_called_once_with(fake_proc)
            fake_proc.wait.assert_called_once()

    def test_libreoffice_uses_isolated_profile_and_handles_non_xlsx_conversion(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            src_path = Path(tmp) / "legacy_model.xlsm"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = 100
            wb.save(src_path)

            captured_cmd: list[str] = []
            xcu_contents: list[str] = []

            def fake_popen(cmd, **kwargs):
                captured_cmd.extend(cmd)
                outdir = Path(cmd[cmd.index("--outdir") + 1])
                xcu = outdir / "lo_profile" / "user" / "registrymodifications.xcu"
                if xcu.exists():
                    xcu_contents.append(xcu.read_text(encoding="utf-8"))
                converted = outdir / f"{src_path.stem}.xlsx"
                shutil.copy2(src_path, converted)
                proc = mock.MagicMock()
                proc.returncode = 0
                proc.communicate.return_value = ("", "")
                return proc

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=fake_popen),
            ):
                rc, payload = self._run_main([str(src_path)])

            self.assertEqual(rc, 0)
            self.assertEqual(payload["status"], "success")
            self.assertTrue(payload["recalculated_with_libreoffice"])
            self.assertTrue(
                any(arg.startswith("-env:UserInstallation=file://") and arg.endswith("/lo_profile") for arg in captured_cmd),
                f"Expected isolated LibreOffice profile flag in {captured_cmd}",
            )
            self.assertEqual(len(xcu_contents), 1, "Expected registrymodifications.xcu to be written")
            self.assertIn("OOXMLRecalcMode", xcu_contents[0])
            self.assertIn("ODFRecalcMode", xcu_contents[0])
            converted_target = src_path.with_suffix(".xlsx")
            self.assertEqual(Path(payload["file"]), converted_target.resolve())
            self.assertTrue(converted_target.exists())

    def test_checks_tab_quoted_error_literals_do_not_false_positive(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "checks_tab.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Checks"
            ws["A2"] = 10
            ws["B1"] = '=COUNTIF(A2:A10, "#REF!")'
            ws["B2"] = '=IFERROR(A2/2, "#DIV/0!")'
            ws["B3"] = '=IF(A2="say ""#NAME?""", 0, 1)'
            wb.save(path)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc, payload = self._run_main([str(path), "--allow-unrecalculated"])

            self.assertEqual(rc, 0, payload)
            self.assertEqual(payload["status"], "success")
            self.assertEqual(payload["total_errors"], 0)
            self.assertEqual(payload["error_summary"], {})
            self.assertEqual(payload["total_formulas"], 3)

    def test_array_formula_counted_and_scanned_accurately(self) -> None:
        from openpyxl.worksheet.formula import ArrayFormula

        with tempfile.TemporaryDirectory() as tmp:
            clean_path = Path(tmp) / "array_clean.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Model"
            ws["A1"] = 5
            ws["B1"] = ArrayFormula("B1:B1", "=A1*2")
            ws["B2"] = ArrayFormula("B2:B2", '=COUNTIF(A1:A5, "#REF!")')
            wb.save(clean_path)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_strict, payload_strict = self._run_main([str(clean_path)])

            self.assertEqual(rc_strict, 2, payload_strict)
            self.assertEqual(payload_strict["status"], "unverified")
            self.assertEqual(payload_strict["total_formulas"], 2)
            self.assertEqual(payload_strict["uncached_formulas"], 2)
            self.assertEqual(payload_strict["total_errors"], 0)

            broken_path = Path(tmp) / "array_broken.xlsx"
            wb2 = openpyxl.Workbook()
            ws2 = wb2.active
            ws2.title = "Model"
            ws2["B1"] = ArrayFormula("B1:B1", "=SUM(A1,#REF!)")
            wb2.save(broken_path)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_broken, payload_broken = self._run_main([str(broken_path)])

            self.assertEqual(rc_broken, 1, payload_broken)
            self.assertEqual(payload_broken["status"], "errors_found")
            self.assertEqual(payload_broken["total_formulas"], 1)
            self.assertEqual(payload_broken["total_errors"], 1)
            self.assertEqual(payload_broken["error_summary"]["#REF!"]["locations"], ["Model!B1"])

    def test_corrupt_xlsx_and_directory_return_structured_json_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            corrupt = Path(tmp) / "corrupt.xlsx"
            corrupt.write_bytes(b"definitely not a zip archive")

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_corrupt, payload_corrupt = self._run_main([str(corrupt)])
            self.assertEqual(rc_corrupt, 1)
            self.assertEqual(payload_corrupt["status"], "error")
            self.assertIn("Failed to open workbook", payload_corrupt["message"])

            # Verify CLI execution never dumps a raw Python traceback
            proc = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(tmp)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(proc.returncode, 1)
            self.assertNotIn("Traceback", proc.stdout + proc.stderr)
            dir_payload = json.loads(proc.stdout)
            self.assertEqual(dir_payload["status"], "error")
            self.assertIn("Not a file", dir_payload["message"])

    def test_cached_empty_string_formula_is_not_flagged_uncached(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "cached_empty_string.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Model"
            ws["A1"] = 10
            ws["B1"] = '=IF(A1>0,"","Alert")'
            wb.save(path)
            _rewrite_formula_cell_as_cached_empty_string(path, "B1")

            # Confirm openpyxl genuinely loads cached "" formulas as value=None, data_type='str'
            wb_check = openpyxl.load_workbook(path, data_only=True)
            try:
                cell_v = wb_check["Model"]["B1"]
                self.assertIsNone(cell_v.value)
                self.assertEqual(cell_v.data_type, "str")
            finally:
                wb_check.close()

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc, payload = self._run_main([str(path)])

            self.assertEqual(rc, 0, payload)
            self.assertEqual(payload["status"], "success")
            self.assertEqual(payload["total_formulas"], 1)
            self.assertEqual(payload["uncached_formulas"], 0)
            self.assertEqual(payload["total_errors"], 0)

    def test_atomic_copy_oserror_returns_structured_json_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "readonly_target.xlsx"
            wb = openpyxl.Workbook()
            wb.save(path)

            def fake_popen(cmd, **kwargs):
                outdir = Path(cmd[cmd.index("--outdir") + 1])
                converted = outdir / f"{path.stem}.xlsx"
                shutil.copy2(path, converted)
                proc = mock.MagicMock()
                proc.returncode = 0
                proc.communicate.return_value = ("", "")
                return proc

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=fake_popen),
                mock.patch.object(
                    recalc_xlsx_mod,
                    "atomic_copy",
                    side_effect=PermissionError("Read-only output directory"),
                ),
            ):
                rc, payload = self._run_main([str(path)])

            self.assertEqual(rc, 1, payload)
            self.assertEqual(payload["status"], "error")
            self.assertFalse(payload["recalculated_with_libreoffice"])
            self.assertIn("Failed to write recalculated workbook", payload["message"])

    def test_chartsheet_is_skipped_safely(self) -> None:
        from openpyxl.chart import BarChart, Reference

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "with_chartsheet.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Model"
            ws["A1"] = "Revenue"
            ws["B1"] = 100.0
            cs = wb.create_chartsheet("Valuation Chart")
            chart = BarChart()
            chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=1))
            cs.add_chart(chart)
            wb.save(path)

            # Confirm the saved workbook genuinely contains a Chartsheet in sheetnames
            wb_check = openpyxl.load_workbook(path)
            try:
                self.assertIn("Valuation Chart", wb_check.sheetnames)
                self.assertEqual([w.title for w in wb_check.worksheets], ["Model"])
            finally:
                wb_check.close()

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc, payload = self._run_main([str(path)])

            self.assertEqual(rc, 0, payload)
            self.assertEqual(payload["status"], "success")
            self.assertEqual(payload["total_errors"], 0)
            self.assertEqual(payload["total_formulas"], 0)

    def test_malformed_workbook_attribute_error_returns_structured_json_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # An empty Chartsheet (no drawing rels) triggers openpyxl's AttributeError on load
            path = Path(tmp) / "malformed_chartsheet.xlsx"
            wb = openpyxl.Workbook()
            wb.create_chartsheet("Empty Chart")
            wb.save(path)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc, payload = self._run_main([str(path)])

            self.assertEqual(rc, 1, payload)
            self.assertEqual(payload["status"], "error")
            self.assertIn("Failed to open workbook", payload["message"])

    def test_zero_byte_and_non_zip_rejected_before_libreoffice_without_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            empty_xlsx = Path(tmp) / "empty.xlsx"
            empty_xlsx.write_bytes(b"")
            non_zip_xlsm = Path(tmp) / "fake.xlsm"
            non_zip_xlsm.write_bytes(b"not a zip file")

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
            ):
                rc_empty, payload_empty = self._run_main([str(empty_xlsx)])
                rc_fake, payload_fake = self._run_main([str(non_zip_xlsm)])

            mock_popen.assert_not_called()
            self.assertEqual(rc_empty, 1)
            self.assertEqual(payload_empty["status"], "error")
            self.assertEqual(empty_xlsx.read_bytes(), b"")
            self.assertEqual(rc_fake, 1)
            self.assertEqual(payload_fake["status"], "error")
            self.assertEqual(non_zip_xlsm.read_bytes(), b"not a zip file")

    def test_xlsm_sibling_xlsx_protection_refuses_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            xlsm_path = Path(tmp) / "model.xlsm"
            sibling_xlsx = Path(tmp) / "model.xlsx"
            wb = openpyxl.Workbook()
            wb.save(xlsm_path)
            sibling_xlsx.write_bytes(b"precious sibling content")

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
            ):
                rc, payload = self._run_main([str(xlsm_path)])

            mock_popen.assert_not_called()
            self.assertEqual(rc, 1, payload)
            self.assertEqual(payload["status"], "error")
            self.assertIn("--output", payload["message"])
            self.assertIn("--check-only", payload["message"])
            self.assertEqual(sibling_xlsx.read_bytes(), b"precious sibling content")

    def test_check_only_and_output_flags_preserve_input_and_permissions(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            src_path = Path(tmp) / "sensitive.xlsx"
            out_path = Path(tmp) / "recalced_copy.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = 10
            ws["A2"] = "=A1*2"
            wb.save(src_path)
            if os.name != "nt":
                os.chmod(src_path, 0o640)

            orig_bytes = src_path.read_bytes()
            orig_stat = src_path.stat()

            def fake_popen(cmd, **kwargs):
                outdir = Path(cmd[cmd.index("--outdir") + 1])
                converted = outdir / f"{src_path.stem}.xlsx"
                # Simulate LibreOffice creating a 0o644 recalculated output
                shutil.copyfile(src_path, converted)
                if os.name != "nt":
                    os.chmod(converted, 0o644)
                proc = mock.MagicMock()
                proc.returncode = 0
                proc.communicate.return_value = ("", "")
                return proc

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=fake_popen),
            ):
                # 1. --check-only must not touch src_path or create extra files
                rc_chk, payload_chk = self._run_main([str(src_path), "--check-only"])
                self.assertEqual(rc_chk, 0, payload_chk)
                self.assertTrue(payload_chk["check_only"])
                self.assertEqual(src_path.read_bytes(), orig_bytes)
                self.assertEqual(src_path.stat().st_mtime_ns, orig_stat.st_mtime_ns)
                if os.name != "nt":
                    self.assertEqual(src_path.stat().st_mode & 0o777, 0o640)

                # 2. --output writes to out_path without touching src_path
                rc_out, payload_out = self._run_main([str(src_path), "--output", str(out_path)])
                self.assertEqual(rc_out, 0, payload_out)
                self.assertEqual(Path(payload_out["file"]), out_path.resolve())
                self.assertEqual(Path(payload_out["source_file"]), src_path.resolve())
                self.assertTrue(out_path.exists())
                self.assertEqual(src_path.read_bytes(), orig_bytes)
                if os.name != "nt":
                    self.assertEqual(out_path.stat().st_mode & 0o777, 0o640)

                # 3. In-place update preserves 0o640 permission bits (never loosened to 0o644)
                rc_inplace, payload_inplace = self._run_main([str(src_path)])
                self.assertEqual(rc_inplace, 0, payload_inplace)
                if os.name != "nt":
                    self.assertEqual(src_path.stat().st_mode & 0o777, 0o640)

    def test_output_flag_without_libreoffice_emits_error_and_never_succeeds_silently(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            src_path = Path(tmp) / "input.xlsx"
            out_path = Path(tmp) / "recalced.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = 42
            wb.save(src_path)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                for extra_args in ([], ["--allow-unrecalculated"]):
                    with self.subTest(extra_args=extra_args):
                        rc, payload = self._run_main(
                            [str(src_path), "--output", str(out_path), *extra_args]
                        )
                        self.assertEqual(rc, 1, payload)
                        self.assertEqual(payload["status"], "error")
                        self.assertEqual(Path(payload["file"]), out_path.resolve())
                        self.assertEqual(Path(payload["source_file"]), src_path.resolve())
                        self.assertFalse(payload["recalculated_with_libreoffice"])
                        self.assertEqual(
                            payload["message"],
                            "Cannot write recalculated --output workbook because "
                            "LibreOffice (soffice) is not installed on PATH.",
                        )
                        self.assertFalse(out_path.exists())

    def test_real_process_group_timeout_kills_spawned_child(self) -> None:
        if os.name == "nt":
            self.skipTest("POSIX process-group kill test")
        with tempfile.TemporaryDirectory() as tmp:
            tmp_p = Path(tmp)
            xlsx_path = tmp_p / "hang.xlsx"
            wb = openpyxl.Workbook()
            wb.save(xlsx_path)

            child_pid_file = tmp_p / "child.pid"
            fake_soffice = tmp_p / "fake_soffice.sh"
            fake_soffice.write_text(
                "#!/bin/sh\n"
                f"sleep 60 &\n"
                f"echo $! > '{child_pid_file}'\n"
                "wait\n",
                encoding="utf-8",
            )
            os.chmod(fake_soffice, 0o755)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=str(fake_soffice)):
                rc, payload = self._run_main([str(xlsx_path), "1"])

            self.assertEqual(rc, 1, payload)
            self.assertEqual(payload["status"], "error")
            self.assertIn("timed out", payload["message"])
            if child_pid_file.exists():
                child_pid = int(child_pid_file.read_text(encoding="utf-8").strip())
                # Give kernel a brief moment to reap the SIGKILLed child
                deadline = time.time() + 2.0
                alive = True
                while time.time() < deadline:
                    try:
                        os.kill(child_pid, 0)
                        time.sleep(0.05)
                    except ProcessLookupError:
                        alive = False
                        break
                self.assertFalse(alive, f"Orphaned child PID {child_pid} survived timeout!")

    @unittest.skipIf(
        recalc_xlsx_mod.find_soffice() is None,
        "LibreOffice (soffice) is required for real recalculation integration test",
    )
    def test_real_libreoffice_forces_recalculation_of_stale_cached_values(self) -> None:
        import xml.etree.ElementTree as ET

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "stale_cache.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Stale"
            ws["A1"] = 10
            ws["A2"] = 0
            ws["B1"] = "alpha"
            ws["C1"] = 42
            ws["D1"] = "=10+10"
            ws["D2"] = "=A1/A2"
            ws["D3"] = '=VLOOKUP("missing",B1:C1,2,FALSE)'
            wb.save(path)
            wb.close()

            # Inject intentionally wrong stale cached values (999, 0, 0)
            ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            ET.register_namespace("", ns)
            with zipfile.ZipFile(path, "r") as zin:
                entries = {name: zin.read(name) for name in zin.namelist()}
            root = ET.fromstring(entries["xl/worksheets/sheet1.xml"])
            for cell in root.iter(f"{{{ns}}}c"):
                ref = cell.attrib.get("r")
                if ref in ("D1", "D2", "D3"):
                    cell.attrib.pop("t", None)
                    v_el = cell.find(f"{{{ns}}}v")
                    if v_el is None:
                        v_el = ET.SubElement(cell, f"{{{ns}}}v")
                    v_el.text = "999" if ref == "D1" else "0"
            entries["xl/worksheets/sheet1.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in entries.items():
                    zout.writestr(name, data)

            # Confirm stale cached values exist prior to recalculation
            wb_before = openpyxl.load_workbook(path, data_only=True)
            try:
                self.assertEqual(wb_before["Stale"]["D1"].value, 999)
                self.assertEqual(wb_before["Stale"]["D2"].value, 0)
                self.assertEqual(wb_before["Stale"]["D3"].value, 0)
            finally:
                wb_before.close()

            # First verify --check-only catches the errors WITHOUT modifying path
            orig_bytes = path.read_bytes()
            rc_chk, payload_chk = self._run_main([str(path), "30", "--check-only"])
            self.assertEqual(rc_chk, 1, payload_chk)
            self.assertEqual(payload_chk["status"], "errors_found")
            self.assertTrue(payload_chk["recalculated_with_libreoffice"])
            self.assertEqual(payload_chk["total_errors"], 2)
            self.assertIn("#DIV/0!", payload_chk["error_summary"])
            self.assertIn("#N/A", payload_chk["error_summary"])
            self.assertEqual(path.read_bytes(), orig_bytes)

            # Now run in-place recalculation and verify cached values on disk are updated
            rc, payload = self._run_main([str(path), "30"])
            self.assertEqual(rc, 1, payload)
            self.assertEqual(payload["status"], "errors_found")
            self.assertTrue(payload["recalculated_with_libreoffice"])
            self.assertEqual(payload["total_errors"], 2)
            self.assertEqual(payload["error_summary"]["#DIV/0!"]["locations"], ["Stale!D2"])
            self.assertEqual(payload["error_summary"]["#N/A"]["locations"], ["Stale!D3"])

            wb_after = openpyxl.load_workbook(path, data_only=True)
            try:
                self.assertEqual(wb_after["Stale"]["D1"].value, 20)
                self.assertEqual(wb_after["Stale"]["D2"].value, "#DIV/0!")
                self.assertEqual(wb_after["Stale"]["D3"].value, "#N/A")
            finally:
                wb_after.close()


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
            # Missing required WACC sheet MUST cause FAIL (eliminating vacuous PASS)
            self.assertEqual(results["status"], "FAIL", results)
            self.assertTrue(any("Required sheet missing: WACC" in e for e in results["errors"]), results["errors"])
            self.assertFalse(
                any("AttributeError" in w or "has no attribute" in w for w in results["warnings"]),
                results["warnings"],
            )
            # WACC range check still falls back cleanly to DCF sheet
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

    def test_multi_scenario_invalid_base_case_with_sensitivity_table_and_full_wacc_label(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "multi_scenario_invalid_base.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"

            # Bear Case (valid: g=2% < WACC=11%)
            ws_dcf["A5"] = "BEAR CASE ASSUMPTIONS"
            ws_dcf["A6"] = "WACC"
            ws_dcf["B6"] = 0.11
            ws_dcf["A7"] = "Terminal Growth Rate"
            ws_dcf["B7"] = 0.02

            # Base Case (INVALID: g=10% > WACC=9%)
            ws_dcf["A12"] = "BASE CASE ASSUMPTIONS"
            ws_dcf["A13"] = "WACC"
            ws_dcf["B13"] = 0.09
            ws_dcf["A14"] = "Terminal Growth Rate"
            ws_dcf["B14"] = 0.10

            # Bull Case (valid: g=3% < WACC=8%)
            ws_dcf["A19"] = "BULL CASE ASSUMPTIONS"
            ws_dcf["A20"] = "WACC"
            ws_dcf["B20"] = 0.08
            ws_dcf["A21"] = "Terminal Growth Rate"
            ws_dcf["B21"] = 0.03

            ws_dcf["A50"] = "PV of Terminal Value"
            ws_dcf["B50"] = 600.0
            ws_dcf["A51"] = "Enterprise Value"
            ws_dcf["B51"] = 1000.0

            # Inline sensitivity table at row 87 per dcf-model/SKILL.md spec
            ws_dcf["A86"] = "SENSITIVITY ANALYSIS"
            ws_dcf["A87"] = "WACC vs Terminal Growth"
            ws_dcf["B87"] = 0.02
            ws_dcf["C87"] = 0.025
            ws_dcf["D87"] = 0.03

            # Spec-compliant WACC tab using "WEIGHTED AVERAGE COST OF CAPITAL" output label
            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A1"] = "COST OF EQUITY CALCULATION"
            ws_wacc["A20"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B20"] = 0.09
            wb.save(path)

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "FAIL", results)
            self.assertEqual(results["error_count"], 1, results["errors"])
            self.assertTrue(
                any("Terminal growth (10.00% at DCF!B14) >= WACC (9.00% at DCF!B13)" in e for e in results["errors"]),
                results["errors"],
            )
            # WACC tab label "WEIGHTED AVERAGE COST OF CAPITAL" must be located with zero warnings
            self.assertEqual(results["warning_count"], 0, results["warnings"])
            self.assertTrue(any("WACC (9.00%) in reasonable range" in i for i in results["info"]), results["info"])

    def test_horizontal_scenarios_and_boolean_exclusion_in_validate_dcf(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "horizontal_scenarios.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A4"] = "Scenario"
            ws_dcf["B4"] = "Bear"
            ws_dcf["C4"] = "Base"
            ws_dcf["D4"] = "Bull"
            ws_dcf["A5"] = "WACC"
            ws_dcf["B5"] = 0.11
            ws_dcf["C5"] = 0.09
            ws_dcf["D5"] = 0.08
            ws_dcf["A6"] = "Terminal Growth Rate"
            ws_dcf["B6"] = 0.02
            ws_dcf["C6"] = 0.095  # INVALID Base Case (g > WACC)
            ws_dcf["D6"] = 0.03

            # Put a Python bool next to PV of Terminal Value to ensure bool != 1
            ws_dcf["A10"] = "PV of Terminal Value"
            ws_dcf["B10"] = True
            ws_dcf["C10"] = 600.0
            ws_dcf["A11"] = "Enterprise Value"
            ws_dcf["B11"] = 1000.0
            ws_dcf["A20"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A5"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B5"] = 0.09
            wb.save(path)

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "FAIL", results)
            self.assertTrue(
                any("9.50%" in e and "9.00%" in e for e in results["errors"]),
                results["errors"],
            )
            self.assertTrue(
                any("Terminal value is 60.0% of EV" in i for i in results["info"]),
                results["info"],
            )

    def test_validate_dcf_handles_checks_countif_and_array_formula(self) -> None:
        from openpyxl.worksheet.formula import ArrayFormula

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "dcf_with_checks_and_array.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=False)
            wb = openpyxl.load_workbook(path)
            ws_checks = wb.create_sheet("Checks")
            ws_checks["A1"] = "Ref Errors"
            ws_checks["B1"] = '=COUNTIF(DCF!A1:Z50, "#REF!")'
            ws_dcf = wb["DCF"]
            ws_dcf["C10"] = ArrayFormula("C10:C10", "=B6+B7")
            wb.save(path)

            results = validate_dcf_mod.validate_dcf_model(str(path))
            # Neither Checks!B1 nor DCF!C10 has a formula error; both should be counted as uncached formulas
            self.assertFalse(any("#REF!" in e for e in results["errors"]), results["errors"])
            self.assertTrue(any("Total formulas: 2" in i for i in results["info"]), results["info"])
            self.assertTrue(any("Uncached formulas: 2" in i for i in results["info"]), results["info"])

    def test_cached_empty_string_formula_passes_dcf_validation(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "dcf_cached_empty_str.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=False)
            wb = openpyxl.load_workbook(path)
            wb["DCF"]["B8"] = '=IF(B6>0,"","Flag")'
            wb.save(path)
            _rewrite_formula_cell_as_cached_empty_string(path, "B8")

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "PASS", results)
            self.assertEqual(results["error_count"], 0, results["errors"])
            self.assertTrue(any("Total formulas: 1" in i for i in results["info"]), results["info"])
            self.assertTrue(any("Uncached formulas: 0" in i for i in results["info"]), results["info"])

    def test_wacc_vs_terminal_growth_header_detected_as_inline_sensitivity(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "dcf_wacc_vs_tg_header.xlsx"
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
            # Spec-compliant header without the literal word "sensitivity"
            ws_dcf["A10"] = "WACC vs Terminal Growth"
            ws_dcf["B10"] = 0.02
            ws_dcf["C10"] = 0.025
            ws_dcf["D10"] = 0.03

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A3"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B3"] = 0.09
            wb.save(path)

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "PASS", results)
            self.assertEqual(results["warning_count"], 0, results["warnings"])
            self.assertTrue(
                any("Found inline sensitivity analysis on DCF sheet" in i for i in results["info"]),
                results["info"],
            )

    def test_downstream_ratio_and_multiple_rows_do_not_corrupt_tv_proportion(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "dcf_with_crosschecks.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=False)
            wb = openpyxl.load_workbook(path)
            ws_dcf = wb["DCF"]
            # Downstream cross-check rows placed between Valuation Summary and Sensitivity
            ws_dcf["A8"] = "PV of Terminal Value (% of EV)"
            ws_dcf["B8"] = 0.60
            ws_dcf["A9"] = "Implied Enterprise Value / EBITDA"
            ws_dcf["B9"] = 12.5
            ws_dcf["A10"] = "Enterprise Value / Revenue"
            ws_dcf["B10"] = 3.2
            ws_dcf["A12"] = "Sensitivity Analysis (WACC vs Terminal Growth)"
            ws_dcf["A13"] = "Implied Enterprise Value"
            ws_dcf["B13"] = 50.0
            wb.save(path)

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "PASS", results)
            self.assertEqual(results["error_count"], 0, results["errors"])
            self.assertEqual(results["warning_count"], 0, results["warnings"])
            self.assertTrue(
                any("Terminal value is 60.0% of EV" in i for i in results["info"]),
                results["info"],
            )

            # Also verify ratio rows appearing ABOVE the dollar summary rows are ignored
            path_above = Path(tmp) / "dcf_ratio_above.xlsx"
            wb2 = openpyxl.Workbook()
            ws2 = wb2.active
            ws2.title = "DCF"
            ws2["A2"] = "WACC"
            ws2["B2"] = 0.09
            ws2["A3"] = "Terminal Growth Rate"
            ws2["B3"] = 0.025
            ws2["A5"] = "PV of Terminal Value (% of EV)"
            ws2["B5"] = 0.60
            ws2["A6"] = "Implied Enterprise Value / EBITDA"
            ws2["B6"] = 12.5
            ws2["A7"] = "PV of Terminal Value"
            ws2["B7"] = 600.0
            ws2["A8"] = "Enterprise Value"
            ws2["B8"] = 1000.0
            ws2["A12"] = "Sensitivity Analysis"
            ws_wacc2 = wb2.create_sheet("WACC")
            ws_wacc2["A3"] = "WACC"
            ws_wacc2["B3"] = 0.09
            wb2.save(path_above)

            results_above = validate_dcf_mod.validate_dcf_model(str(path_above))
            self.assertEqual(results_above["status"], "PASS", results_above)
            self.assertEqual(results_above["warning_count"], 0, results_above["warnings"])
            self.assertTrue(
                any("Terminal value is 60.0% of EV" in i for i in results_above["info"]),
                results_above["info"],
            )

    def test_chartsheet_in_dcf_workbook_is_skipped_safely(self) -> None:
        from openpyxl.chart import BarChart, Reference

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "dcf_with_chartsheet.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=False)
            wb = openpyxl.load_workbook(path)
            ws_dcf = wb["DCF"]
            cs = wb.create_chartsheet("Football Field Chart")
            chart = BarChart()
            chart.add_data(Reference(ws_dcf, min_col=2, min_row=6, max_row=7))
            cs.add_chart(chart)
            wb.save(path)

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "PASS", results)
            self.assertEqual(results["error_count"], 0, results["errors"])
            self.assertEqual(results["warning_count"], 0, results["warnings"])

    def test_empty_and_missing_sheet_workbooks_fail_never_vacuous_pass(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # 1. Completely empty default workbook (no DCF or WACC sheets)
            empty_default = Path(tmp) / "empty_default.xlsx"
            wb = openpyxl.Workbook()
            wb.save(empty_default)
            res1 = validate_dcf_mod.validate_dcf_model(str(empty_default))
            self.assertEqual(res1["status"], "FAIL", res1)
            self.assertTrue(any("Required sheet missing: DCF" in e for e in res1["errors"]), res1["errors"])
            self.assertTrue(any("Required sheet missing: WACC" in e for e in res1["errors"]), res1["errors"])

            # 2. Empty DCF + empty WACC sheets (missing critical DCF components)
            empty_tabs = Path(tmp) / "empty_tabs.xlsx"
            wb2 = openpyxl.Workbook()
            wb2.active.title = "DCF"
            wb2.create_sheet("WACC")
            wb2.save(empty_tabs)
            res2 = validate_dcf_mod.validate_dcf_model(str(empty_tabs))
            self.assertEqual(res2["status"], "FAIL", res2)
            self.assertGreaterEqual(res2["error_count"], 3, res2["errors"])

    def test_wacc_weight_labels_excluded_and_out_of_range_dcf_wacc_caught(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "wacc_weights.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=False)
            wb = openpyxl.load_workbook(path)
            ws_wacc = wb["WACC"]
            ws_wacc["A1"] = "Equity weight in WACC"
            ws_wacc["B1"] = 0.80
            ws_wacc["A2"] = "Debt weight in WACC"
            ws_wacc["B2"] = 0.20
            ws_wacc["A4"] = "WACC discount factor"
            ws_wacc["B4"] = 0.65
            wb.save(path)

            res = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res["status"], "PASS", res)
            self.assertEqual(res["warning_count"], 0, f"Weight/factor labels misread as WACC: {res['warnings']}")

            # Now add an out-of-range scenario WACC (25%) on the DCF sheet; union check must catch it
            wb2 = openpyxl.load_workbook(path)
            wb2["DCF"]["A5"] = "Bear WACC"
            wb2["DCF"]["B5"] = 0.25
            wb2.save(path)

            res2 = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertTrue(
                any("25.00%" in w and "outside typical range" in w for w in res2["warnings"]),
                f"Expected out-of-range DCF scenario WACC warning, got: {res2['warnings']}",
            )

    def test_validate_dcf_cli_subprocess_pass_and_fail(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            good_path = Path(tmp) / "good_dcf.xlsx"
            bad_path = Path(tmp) / "empty_dcf.xlsx"
            out_json = Path(tmp) / "report.json"
            self._build_compliant_two_sheet_dcf(good_path, with_formula=False)
            openpyxl.Workbook().save(bad_path)

            proc_ok = subprocess.run(
                [sys.executable, str(VALIDATE_DCF_PATH), str(good_path), str(out_json)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(proc_ok.returncode, 0, proc_ok.stdout + proc_ok.stderr)
            payload_ok = json.loads(proc_ok.stdout)
            self.assertEqual(payload_ok["status"], "PASS")
            self.assertTrue(out_json.exists())
            self.assertEqual(json.loads(out_json.read_text(encoding="utf-8"))["status"], "PASS")

            proc_fail = subprocess.run(
                [sys.executable, str(VALIDATE_DCF_PATH), str(bad_path)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(proc_fail.returncode, 1, proc_fail.stdout + proc_fail.stderr)
            payload_fail = json.loads(proc_fail.stdout)
            self.assertEqual(payload_fail["status"], "FAIL")


@unittest.skipIf(openpyxl is None, "openpyxl is required for DCF skeleton tests")
class TestBuildDcfSkeleton(unittest.TestCase):
    @unittest.skipIf(
        recalc_xlsx_mod.find_soffice() is None,
        "LibreOffice (soffice) is required for end-to-end DCF skeleton verification",
    )
    def test_build_recalc_validate_and_switch_cases_end_to_end(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "institutional_dcf.xlsx"

            # Build default Base Case (Case = 2) via CLI
            build_proc = subprocess.run(
                [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(model_path)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)
            self.assertTrue(model_path.exists())

            prices_by_case: dict[int, float] = {}
            expected_labels = {1: "Bear Case", 2: "Base Case", 3: "Bull Case"}

            for case_num in (2, 1, 3):
                if case_num != 2:
                    # Switch named Case cell DCF!B4 directly and re-save
                    wb_edit = openpyxl.load_workbook(model_path)
                    wb_edit["DCF"]["B4"] = case_num
                    wb_edit.save(model_path)
                    wb_edit.close()

                # Recalculate with real LibreOffice
                recalc_proc = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_proc.returncode, 0, recalc_proc.stdout + recalc_proc.stderr)
                recalc_payload = json.loads(recalc_proc.stdout)
                self.assertEqual(recalc_payload["status"], "success", recalc_payload)
                self.assertTrue(recalc_payload["recalculated_with_libreoffice"])
                self.assertEqual(recalc_payload["total_errors"], 0)
                self.assertEqual(recalc_payload["uncached_formulas"], 0)
                self.assertGreater(recalc_payload["total_formulas"], 100)

                # Validate with hardened validate_dcf.py -> PASS with 0 errors and 0 warnings
                val_proc = subprocess.run(
                    [sys.executable, str(VALIDATE_DCF_PATH), str(model_path)],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(val_proc.returncode, 0, val_proc.stdout + val_proc.stderr)
                val_payload = json.loads(val_proc.stdout)
                self.assertEqual(val_payload["status"], "PASS", val_payload)
                self.assertEqual(val_payload["error_count"], 0, val_payload["errors"])
                self.assertEqual(val_payload["warning_count"], 0, val_payload["warnings"])

                # Inspect calculated values directly
                wb_val = openpyxl.load_workbook(model_path, data_only=True)
                try:
                    ws_dcf = wb_val["DCF"]
                    ws_wacc = wb_val["WACC"]
                    self.assertEqual(ws_dcf["C4"].value, expected_labels[case_num])
                    # Gross-debt capital weights: non-negative and sum to 100%
                    self.assertGreaterEqual(ws_wacc["C23"].value, 0.0)
                    self.assertAlmostEqual(ws_wacc["C24"].value, 1.0, places=6)
                    # Center cell of 5x5 sensitivity matches Active Case Implied Share Price
                    implied_price = float(ws_dcf["C107"].value)
                    sens_center = float(ws_dcf["E126"].value)
                    self.assertAlmostEqual(implied_price, sens_center, places=4)
                    prices_by_case[case_num] = implied_price
                finally:
                    wb_val.close()

            self.assertLess(prices_by_case[1], prices_by_case[2])
            self.assertLess(prices_by_case[2], prices_by_case[3])


if __name__ == "__main__":
    unittest.main()
