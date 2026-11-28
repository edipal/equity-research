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
XLSX_SCAN_PATH = SCRIPTS_DIR / "xlsx_scan.py"
VALIDATE_DCF_PATH = ROOT / ".agents" / "skills" / "dcf-model" / "scripts" / "validate_dcf.py"
BUILD_DCF_SKELETON_PATH = ROOT / ".agents" / "skills" / "dcf-model" / "scripts" / "build_dcf_skeleton.py"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


validate_harness_mod = load_module("validate_harness", VALIDATE_HARNESS_PATH)
xlsx_scan_mod = load_module("xlsx_scan", XLSX_SCAN_PATH)
recalc_xlsx_mod = load_module("recalc_xlsx", RECALC_XLSX_PATH)
validate_dcf_mod = load_module("validate_dcf", VALIDATE_DCF_PATH)

REQUIRE_FULL_TEST_SUITE = os.environ.get("REQUIRE_FULL_TEST_SUITE", "").strip().lower() in (
    "1",
    "true",
    "yes",
)


def _skip_unless_required(condition: bool, reason: str):
    """Skip test when condition is True unless REQUIRE_FULL_TEST_SUITE is enabled."""
    if condition and not REQUIRE_FULL_TEST_SUITE:
        return unittest.skip(reason)
    return lambda obj: obj


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

    def test_parses_crlf_frontmatter_cleanly(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            # Rewrite rule, skill, and agent files with Windows CRLF (\r\n) line endings
            rule_md = base / ".agents" / "rules" / "integrity.md"
            rule_md.write_bytes(
                b"---\r\ntrigger: always_on\r\ndescription: Core rules\r\n---\r\n\r\n# Integrity\r\n"
            )
            skill_md = base / ".agents" / "skills" / "sample-skill" / "SKILL.md"
            skill_md.write_bytes(
                b"---\r\nname: sample-skill\r\ndescription: Valid CRLF skill.\r\n---\r\n\r\n# Skill\r\n"
            )
            agent_md = base / ".agents" / "agents" / "sample-agent.md"
            agent_md.write_bytes(
                b"---\r\nname: sample-agent\r\ndescription: CRLF agent.\r\n"
                b"mainAgent: true\r\nsubagent: true\r\nmodel: inherit\r\n"
                b"commandExecutionPolicy: sandbox\r\n---\r\n\r\n# Prompt\r\n"
            )

            # Must pass both with PyYAML (if installed) and with the fallback parser
            errors, _, _ = validate_harness_mod.validate_harness(base)
            self.assertEqual(errors, [], f"CRLF frontmatter failed: {errors}")

            with mock.patch.object(validate_harness_mod, "yaml", None):
                fm_skill, body, errs = validate_harness_mod.parse_frontmatter(skill_md, root=base)
                self.assertEqual(errs, [])
                self.assertEqual(fm_skill["name"], "sample-skill")
                self.assertEqual(body, "# Skill")
                errors_fb, _, _ = validate_harness_mod.validate_harness(base)
                self.assertEqual(errors_fb, [], f"CRLF fallback failed: {errors_fb}")

    def test_full_test_suite_prerequisites_present_when_required(self) -> None:
        if not REQUIRE_FULL_TEST_SUITE:
            return
        self.assertIsNotNone(openpyxl, "REQUIRE_FULL_TEST_SUITE=1 requires openpyxl")
        self.assertIsNotNone(validate_harness_mod.yaml, "REQUIRE_FULL_TEST_SUITE=1 requires PyYAML")
        self.assertIsNotNone(
            recalc_xlsx_mod.find_soffice(),
            "REQUIRE_FULL_TEST_SUITE=1 requires LibreOffice (soffice) on PATH",
        )

    def test_command_execution_policy_off_passes_under_pyyaml_and_fallback_while_true_fails(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            create_minimal_valid_harness(base)
            agent_md = base / ".agents" / "agents" / "sample-agent.md"
            agent_md.write_text(
                "---\nname: sample-agent\ndescription: Sample agent.\nmainAgent: true\n"
                "subagent: true\nmodel: inherit\ncommandExecutionPolicy: off\n---\n\n# Prompt\n",
                encoding="utf-8",
            )

            # 1. Under active parser (PyYAML when installed, else fallback) -> PASS
            errors_default, _, _ = validate_harness_mod.validate_harness(base)
            self.assertEqual(errors_default, [], f"Unquoted policy 'off' failed: {errors_default}")

            # 2. Explicitly under stdlib fallback -> PASS
            with mock.patch.object(validate_harness_mod, "yaml", None):
                errors_fb, _, _ = validate_harness_mod.validate_harness(base)
            self.assertEqual(errors_fb, [], f"Unquoted policy 'off' failed under fallback: {errors_fb}")

            # 3. Explicitly verify PyYAML parses unquoted `off` as False and validator normalizes it
            if validate_harness_mod.yaml is not None:
                fm_pyyaml, _, fm_errs = validate_harness_mod.parse_frontmatter(agent_md, root=base)
                self.assertEqual(fm_errs, [])
                self.assertIs(fm_pyyaml["commandExecutionPolicy"], False)

            # 4. Invalid boolean/on values (true, on) MUST fail under both PyYAML and fallback
            for bad_val in ("true", "on"):
                with self.subTest(bad_policy=bad_val):
                    agent_md.write_text(
                        "---\nname: sample-agent\ndescription: Sample agent.\nmainAgent: true\n"
                        f"subagent: true\nmodel: inherit\ncommandExecutionPolicy: {bad_val}\n---\n\n# Prompt\n",
                        encoding="utf-8",
                    )
                    errs_active, _, _ = validate_harness_mod.validate_harness(base)
                    self.assertTrue(
                        any("invalid commandExecutionPolicy" in e for e in errs_active),
                        f"Expected rejection for {bad_val!r}, got: {errs_active}",
                    )
                    with mock.patch.object(validate_harness_mod, "yaml", None):
                        errs_fallback, _, _ = validate_harness_mod.validate_harness(base)
                    self.assertTrue(
                        any("invalid commandExecutionPolicy" in e for e in errs_fallback),
                        f"Expected fallback rejection for {bad_val!r}, got: {errs_fallback}",
                    )


@_skip_unless_required(openpyxl is None, "openpyxl is required for Excel script tests")
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
            out_path = Path(tmp) / "converted_model.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = 100
            wb.save(src_path)

            captured_cmd: list[str] = []
            xcu_contents: list[str] = []

            def fake_popen(cmd, **kwargs):
                captured_cmd.clear()
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
                # 1. Default mode is non-destructive: scans in temp dir without creating sibling .xlsx
                rc_def, payload_def = self._run_main([str(src_path)])
                self.assertEqual(rc_def, 0, payload_def)
                self.assertEqual(payload_def["status"], "success")
                self.assertEqual(payload_def["mode"], "check_only")
                self.assertFalse(payload_def["modified_file"])
                self.assertTrue(payload_def["check_only"])
                self.assertTrue(payload_def["recalculated_with_libreoffice"])
                self.assertFalse(src_path.with_suffix(".xlsx").exists())

                # 2. Explicit --output writes the converted .xlsx using isolated profile
                rc_out, payload_out = self._run_main([str(src_path), "--output", str(out_path)])

            self.assertEqual(rc_out, 0, payload_out)
            self.assertEqual(payload_out["status"], "success")
            self.assertEqual(payload_out["mode"], "output")
            self.assertTrue(payload_out["modified_file"])
            self.assertTrue(payload_out["recalculated_with_libreoffice"])
            self.assertTrue(
                any(arg.startswith("-env:UserInstallation=file://") and arg.endswith("/lo_profile") for arg in captured_cmd),
                f"Expected isolated LibreOffice profile flag in {captured_cmd}",
            )
            self.assertGreaterEqual(len(xcu_contents), 1, "Expected registrymodifications.xcu to be written")
            self.assertIn("OOXMLRecalcMode", xcu_contents[0])
            self.assertIn("ODFRecalcMode", xcu_contents[0])
            self.assertEqual(Path(payload_out["file"]), out_path.resolve())
            self.assertEqual(Path(payload_out["source_file"]), src_path.resolve())
            self.assertTrue(out_path.exists())

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
                rc, payload = self._run_main([str(path), "--in-place"])

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

    def test_xlsm_refuses_in_place_and_default_mode_leaves_sibling_untouched(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            xlsm_path = Path(tmp) / "model.xlsm"
            sibling_xlsx = Path(tmp) / "model.xlsx"
            wb = openpyxl.Workbook()
            wb.save(xlsm_path)
            sibling_xlsx.write_bytes(b"precious sibling content")

            # 1. Explicit --in-place on .xlsm MUST refuse before invoking LibreOffice
            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
            ):
                rc, payload = self._run_main([str(xlsm_path), "--in-place"])

            mock_popen.assert_not_called()
            self.assertEqual(rc, 1, payload)
            self.assertEqual(payload["status"], "error")
            self.assertIn("--in-place", payload["message"])
            self.assertIn("--output", payload["message"])
            self.assertIn("--check-only", payload["message"])
            self.assertEqual(sibling_xlsx.read_bytes(), b"precious sibling content")

            # 2. Default mode (no flags) audits in temp dir without touching sibling .xlsx
            def fake_popen(cmd, **kwargs):
                outdir = Path(cmd[cmd.index("--outdir") + 1])
                shutil.copy2(xlsm_path, outdir / f"{xlsm_path.stem}.xlsx")
                proc = mock.MagicMock()
                proc.returncode = 0
                proc.communicate.return_value = ("", "")
                return proc

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=fake_popen),
            ):
                rc_def, payload_def = self._run_main([str(xlsm_path)])

            self.assertEqual(rc_def, 0, payload_def)
            self.assertFalse(payload_def["modified_file"])
            self.assertEqual(sibling_xlsx.read_bytes(), b"precious sibling content")

    def test_default_check_only_output_and_in_place_flags_preserve_input_and_permissions(self) -> None:
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
                # 0. Default mode (no mode flag) is non-destructive and must not touch src_path
                rc_def, payload_def = self._run_main([str(src_path)])
                self.assertEqual(rc_def, 0, payload_def)
                self.assertEqual(payload_def["mode"], "check_only")
                self.assertFalse(payload_def["modified_file"])
                self.assertTrue(payload_def["check_only"])
                self.assertEqual(src_path.read_bytes(), orig_bytes)
                self.assertEqual(src_path.stat().st_mtime_ns, orig_stat.st_mtime_ns)
                if os.name != "nt":
                    self.assertEqual(src_path.stat().st_mode & 0o777, 0o640)

                # 1. Explicit --check-only alias behaves identically
                rc_chk, payload_chk = self._run_main([str(src_path), "--check-only"])
                self.assertEqual(rc_chk, 0, payload_chk)
                self.assertEqual(payload_chk["mode"], "check_only")
                self.assertFalse(payload_chk["modified_file"])
                self.assertTrue(payload_chk["check_only"])
                self.assertEqual(src_path.read_bytes(), orig_bytes)
                self.assertEqual(src_path.stat().st_mtime_ns, orig_stat.st_mtime_ns)
                if os.name != "nt":
                    self.assertEqual(src_path.stat().st_mode & 0o777, 0o640)

                # 2. --output writes to out_path without touching src_path
                rc_out, payload_out = self._run_main([str(src_path), "--output", str(out_path)])
                self.assertEqual(rc_out, 0, payload_out)
                self.assertEqual(payload_out["mode"], "output")
                self.assertTrue(payload_out["modified_file"])
                self.assertEqual(Path(payload_out["file"]), out_path.resolve())
                self.assertEqual(Path(payload_out["source_file"]), src_path.resolve())
                self.assertTrue(out_path.exists())
                self.assertEqual(src_path.read_bytes(), orig_bytes)
                if os.name != "nt":
                    self.assertEqual(out_path.stat().st_mode & 0o777, 0o640)

                # 3. Explicit --in-place updates src_path atomically while preserving 0o640 permission bits
                rc_inplace, payload_inplace = self._run_main([str(src_path), "--in-place"])
                self.assertEqual(rc_inplace, 0, payload_inplace)
                self.assertEqual(payload_inplace["mode"], "in_place")
                self.assertTrue(payload_inplace["modified_file"])
                self.assertNotIn("check_only", payload_inplace)
                if os.name != "nt":
                    self.assertEqual(src_path.stat().st_mode & 0o777, 0o640)

    def test_write_flags_without_libreoffice_emit_error_and_never_succeed_silently(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            src_path = Path(tmp) / "input.xlsx"
            out_path = Path(tmp) / "recalced.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = 42
            wb.save(src_path)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                for extra_args in ([], ["--allow-unrecalculated"]):
                    with self.subTest(mode="output", extra_args=extra_args):
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

                    with self.subTest(mode="in_place", extra_args=extra_args):
                        rc, payload = self._run_main(
                            [str(src_path), "--in-place", *extra_args]
                        )
                        self.assertEqual(rc, 1, payload)
                        self.assertEqual(payload["status"], "error")
                        self.assertEqual(Path(payload["file"]), src_path.resolve())
                        self.assertFalse(payload["recalculated_with_libreoffice"])
                        self.assertEqual(
                            payload["message"],
                            "Cannot update workbook --in-place because "
                            "LibreOffice (soffice) is not installed on PATH.",
                        )

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

    @_skip_unless_required(
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

            # First verify default mode (and explicit --check-only) catch errors WITHOUT modifying path
            orig_bytes = path.read_bytes()
            for flags in ([], ["--check-only"]):
                with self.subTest(flags=flags):
                    rc_chk, payload_chk = self._run_main([str(path), "30", *flags])
                    self.assertEqual(rc_chk, 1, payload_chk)
                    self.assertEqual(payload_chk["status"], "errors_found")
                    self.assertEqual(payload_chk["mode"], "check_only")
                    self.assertFalse(payload_chk["modified_file"])
                    self.assertTrue(payload_chk["recalculated_with_libreoffice"])
                    self.assertEqual(payload_chk["total_errors"], 2)
                    self.assertIn("#DIV/0!", payload_chk["error_summary"])
                    self.assertIn("#N/A", payload_chk["error_summary"])
                    self.assertEqual(path.read_bytes(), orig_bytes)

            # Now run explicit --in-place recalculation and verify cached values on disk are updated
            rc, payload = self._run_main([str(path), "30", "--in-place"])
            self.assertEqual(rc, 1, payload)
            self.assertEqual(payload["status"], "errors_found")
            self.assertEqual(payload["mode"], "in_place")
            self.assertTrue(payload["modified_file"])
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

    def test_what_if_data_table_refuses_inplace_allows_default_and_check_only_and_counts_formulas(self) -> None:
        from openpyxl.worksheet.formula import DataTableFormula

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "what_if_model.xlsx"
            out_path = Path(tmp) / "exported_copy.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sensitivity"
            ws["A1"] = 10
            ws["A2"] = 2
            ws["B2"] = "=A1*A2"
            ws["C2"] = 10
            ws["D2"] = 20
            ws["B3"] = 2
            ws["B4"] = 3
            ws["C3"] = DataTableFormula(ref="C3:D4", dt2D=True, r1="A1", r2="A2")
            wb.save(path)
            wb.close()

            orig_bytes = path.read_bytes()

            # 1. XML detector finds the data table location accurately
            locs = xlsx_scan_mod.inspect_workbook_xml(path)["data_tables"]
            self.assertEqual(locs, ["Sensitivity!C3:D4"])

            # 2. Scan-only (no soffice): counts DataTableFormula in total_formulas and data_table_formulas
            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_scan, payload_scan = self._run_main([str(path)])
            self.assertEqual(rc_scan, 2, payload_scan)
            self.assertEqual(payload_scan["status"], "unverified")
            self.assertEqual(payload_scan["total_formulas"], 2)  # B2 + C3 DataTableFormula
            self.assertEqual(payload_scan["uncached_formulas"], 2)
            self.assertEqual(payload_scan["data_table_formulas"], 1)
            self.assertEqual(path.read_bytes(), orig_bytes)

            # 3. Explicit --in-place with soffice: MUST REFUSE to overwrite and never call Popen
            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
            ):
                rc_inplace, payload_inplace = self._run_main([str(path), "--in-place"])
            mock_popen.assert_not_called()
            self.assertEqual(rc_inplace, 1, payload_inplace)
            self.assertEqual(payload_inplace["status"], "error")
            self.assertIn("What-If data table", payload_inplace["message"])
            self.assertIn("--check-only", payload_inplace["message"])
            self.assertIn("--output", payload_inplace["message"])
            self.assertEqual(path.read_bytes(), orig_bytes)

            # 4. Default mode, --check-only, and --output work safely without modifying original input
            def fake_popen(cmd, **kwargs):
                outdir = Path(cmd[cmd.index("--outdir") + 1])
                converted = outdir / f"{path.stem}.xlsx"
                shutil.copyfile(path, converted)
                proc = mock.MagicMock()
                proc.returncode = 0
                proc.communicate.return_value = ("", "")
                return proc

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=fake_popen),
            ):
                for flags in ([], ["--check-only"]):
                    rc_chk, payload_chk = self._run_main([str(path), *flags])
                    self.assertEqual(rc_chk, 0, payload_chk)
                    self.assertTrue(payload_chk["check_only"])
                    self.assertFalse(payload_chk["modified_file"])
                    self.assertEqual(payload_chk["data_table_formulas"], 1)
                    self.assertIn("What-If data table", payload_chk.get("warning", ""))
                    self.assertEqual(path.read_bytes(), orig_bytes)

                rc_out, payload_out = self._run_main([str(path), "--output", str(out_path)])
                self.assertEqual(rc_out, 0, payload_out)
                self.assertTrue(payload_out["modified_file"])
                self.assertEqual(payload_out["data_table_formulas"], 1)
                self.assertTrue(out_path.exists())
                self.assertIn("TABLE(...)", payload_out.get("warning", ""))
                self.assertEqual(path.read_bytes(), orig_bytes)

    def test_what_if_data_table_refuses_same_path_output_flag(self) -> None:
        from openpyxl.worksheet.formula import DataTableFormula

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "what_if_same_output.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sensitivity"
            ws["A1"] = 10
            ws["A2"] = 2
            ws["B2"] = "=A1*A2"
            ws["C3"] = DataTableFormula(ref="C3:D4", dt2D=True, r1="A1", r2="A2")
            wb.save(path)
            wb.close()

            orig_bytes = path.read_bytes()
            same_via_dot = str(path.parent / "." / path.name)

            for out_arg in (str(path), same_via_dot):
                with self.subTest(out_arg=out_arg):
                    with (
                        mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                        mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
                    ):
                        rc, payload = self._run_main([str(path), "--output", out_arg])

                    mock_popen.assert_not_called()
                    self.assertEqual(rc, 1, payload)
                    self.assertEqual(payload["status"], "error")
                    self.assertFalse(payload["recalculated_with_libreoffice"])
                    self.assertEqual(payload["data_tables"], ["Sensitivity!C3:D4"])
                    self.assertIn("What-If data table", payload["message"])
                    self.assertEqual(path.read_bytes(), orig_bytes)

    def test_what_if_data_table_refuses_hardlink_output_flag(self) -> None:
        from openpyxl.worksheet.formula import DataTableFormula

        with tempfile.TemporaryDirectory() as tmp:
            dt_path = Path(tmp) / "dt.xlsx"
            hardlink_path = Path(tmp) / "dt_hardlink.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sensitivity"
            ws["A1"] = 10
            ws["A2"] = 2
            ws["B2"] = "=A1*A2"
            ws["C3"] = DataTableFormula(ref="C3:D4", dt2D=True, r1="A1", r2="A2")
            wb.save(dt_path)
            wb.close()

            orig_bytes = dt_path.read_bytes()
            os.link(dt_path, hardlink_path)

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
            ):
                rc, payload = self._run_main([str(dt_path), "--output", str(hardlink_path)])

            mock_popen.assert_not_called()
            self.assertEqual(rc, 1, payload)
            self.assertEqual(payload["status"], "error")
            self.assertFalse(payload["recalculated_with_libreoffice"])
            self.assertEqual(payload["data_tables"], ["Sensitivity!C3:D4"])
            self.assertIn("What-If data table", payload["message"])
            self.assertEqual(dt_path.read_bytes(), orig_bytes)

    def test_what_if_data_table_refuses_case_variant_output_flag(self) -> None:
        from openpyxl.worksheet.formula import DataTableFormula

        with tempfile.TemporaryDirectory() as tmp:
            dt_path = Path(tmp) / "dt.xlsx"
            upper_path = Path(tmp) / "DT.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sensitivity"
            ws["A1"] = 10
            ws["A2"] = 2
            ws["B2"] = "=A1*A2"
            ws["C3"] = DataTableFormula(ref="C3:D4", dt2D=True, r1="A1", r2="A2")
            wb.save(dt_path)
            wb.close()

            orig_bytes = dt_path.read_bytes()
            # Ensure upper_path exists on case-sensitive POSIX filesystems to simulate case-insensitive FS lookup
            if not upper_path.exists():
                shutil.copyfile(dt_path, upper_path)

            real_samefile = os.path.samefile
            resolved_dt = dt_path.resolve()
            resolved_upper = upper_path.resolve()

            def fake_samefile(a, b):
                pair = {Path(a).resolve(), Path(b).resolve()}
                if pair == {resolved_dt, resolved_upper}:
                    return True
                return real_samefile(a, b)

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
                mock.patch.object(recalc_xlsx_mod.os.path, "samefile", side_effect=fake_samefile),
            ):
                rc, payload = self._run_main([str(dt_path), "--output", str(upper_path)])

            mock_popen.assert_not_called()
            self.assertEqual(rc, 1, payload)
            self.assertEqual(payload["status"], "error")
            self.assertFalse(payload["recalculated_with_libreoffice"])
            self.assertEqual(payload["data_tables"], ["Sensitivity!C3:D4"])
            self.assertIn("What-If data table", payload["message"])
            self.assertEqual(dt_path.read_bytes(), orig_bytes)

    def test_output_rejects_non_xlsx_extension_before_libreoffice(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            src_path = Path(tmp) / "input.xlsx"
            wb = openpyxl.Workbook()
            wb.save(src_path)
            wb.close()

            for bad_name in ("model.xlsm", "model.xls", "model_no_ext"):
                bad_out = Path(tmp) / bad_name
                with self.subTest(bad_name=bad_name):
                    with (
                        mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                        mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
                    ):
                        rc, payload = self._run_main([str(src_path), "--output", str(bad_out)])
                    mock_popen.assert_not_called()
                    self.assertEqual(rc, 1, payload)
                    self.assertEqual(payload["status"], "error")
                    self.assertIn(".xlsx", payload["message"])
                    self.assertFalse(bad_out.exists())

    def test_xlsm_scan_only_without_soffice_does_not_fail_sibling_check(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            xlsm_path = Path(tmp) / "coverage.xlsm"
            sibling_xlsx = Path(tmp) / "coverage.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = 123.0
            wb.save(xlsm_path)
            wb.close()
            sibling_xlsx.write_bytes(b"existing sibling untouched")

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc, payload = self._run_main([str(xlsm_path)])

            self.assertEqual(rc, 0, payload)
            self.assertEqual(payload["status"], "success")
            self.assertFalse(payload["recalculated_with_libreoffice"])
            self.assertEqual(sibling_xlsx.read_bytes(), b"existing sibling untouched")

    def test_shared_scan_readonly_handles_stale_dimension_and_unified_error_fallback(self) -> None:
        import xml.etree.ElementTree as ET

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "edge_cases.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "DCF"
            ws["A1"] = "WACC"
            ws["B1"] = 0.09
            ws["A2"] = "Terminal Growth Rate"
            ws["B2"] = 0.025
            ws["A3"] = "PV of Terminal Value"
            ws["B3"] = 600.0
            ws["A4"] = "Enterprise Value"
            ws["B4"] = 1000.0
            ws["A5"] = "Sensitivity Analysis"
            # Place an error-typed cell outside a truncated <dimension> range
            ws["E8"] = "=1/0"
            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A1"] = "WACC"
            ws_wacc["B1"] = 0.09
            wb.save(path)
            wb.close()

            # Mutate sheet1.xml: truncate <dimension> to A1:B5 AND make E8 error-typed with empty <v/>
            ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            ET.register_namespace("", ns)
            with zipfile.ZipFile(path, "r") as zin:
                entries = {name: zin.read(name) for name in zin.namelist()}
            root = ET.fromstring(entries["xl/worksheets/sheet1.xml"])
            dim = root.find(f"{{{ns}}}dimension")
            if dim is not None:
                dim.attrib["ref"] = "A1:B5"
            for cell in root.iter(f"{{{ns}}}c"):
                if cell.attrib.get("r") == "E8":
                    cell.attrib["t"] = "e"
                    v_el = cell.find(f"{{{ns}}}v")
                    if v_el is not None:
                        cell.remove(v_el)
            entries["xl/worksheets/sheet1.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in entries.items():
                    zout.writestr(name, data)

            # Both recalc_xlsx.py and validate_dcf.py MUST find DCF!E8 as unified "#ERROR!"
            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc, recalc_payload = self._run_main([str(path)])
            self.assertEqual(rc, 1, recalc_payload)
            self.assertIn("#ERROR!", recalc_payload["error_summary"])
            self.assertEqual(recalc_payload["error_summary"]["#ERROR!"]["locations"], ["DCF!E8"])

            dcf_results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(dcf_results["status"], "FAIL", dcf_results)
            self.assertTrue(
                any("#ERROR! at DCF!E8" in e for e in dcf_results["errors"]),
                f"Expected unified '#ERROR! at DCF!E8', got: {dcf_results['errors']}",
            )

    def test_missing_openpyxl_fails_upfront_before_libreoffice_or_write(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            src_path = Path(tmp) / "model.xlsx"
            out_path = Path(tmp) / "out.xlsx"
            wb = openpyxl.Workbook()
            wb.save(src_path)
            wb.close()
            orig_bytes = src_path.read_bytes()

            with (
                mock.patch.object(
                    recalc_xlsx_mod,
                    "check_openpyxl_available",
                    return_value="openpyxl is required: No module named 'openpyxl'.",
                ),
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
            ):
                for argv in (
                    [str(src_path)],
                    [str(src_path), "--in-place"],
                    [str(src_path), "--output", str(out_path)],
                ):
                    with self.subTest(argv=argv):
                        rc, payload = self._run_main(argv)
                        self.assertEqual(rc, 1, payload)
                        self.assertEqual(payload["status"], "error")
                        self.assertFalse(payload["recalculated_with_libreoffice"])
                        self.assertIn("openpyxl", payload["message"])

            mock_popen.assert_not_called()
            self.assertEqual(src_path.read_bytes(), orig_bytes)
            self.assertFalse(out_path.exists())

    def test_corrupt_worksheet_xml_in_xlsx_fails_closed_before_libreoffice(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "corrupt_xml.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = 123
            wb.save(path)
            wb.close()

            # Replace xl/worksheets/sheet1.xml with malformed XML while keeping valid zip container
            with zipfile.ZipFile(path, "r") as zin:
                entries = {name: zin.read(name) for name in zin.namelist()}
            entries["xl/worksheets/sheet1.xml"] = b"<worksheet><sheetData><broken"
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in entries.items():
                    zout.writestr(name, data)

            corrupt_bytes = path.read_bytes()

            # 1. inspect_workbook_xml MUST fail closed by raising WorkbookLoadError
            with self.assertRaises(xlsx_scan_mod.WorkbookLoadError):
                xlsx_scan_mod.inspect_workbook_xml(path)

            # 2. recalc_xlsx.py --in-place MUST exit 1 without invoking LibreOffice or modifying the file
            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
            ):
                rc, payload = self._run_main([str(path), "--in-place"])

            mock_popen.assert_not_called()
            self.assertEqual(rc, 1, payload)
            self.assertEqual(payload["status"], "error")
            self.assertFalse(payload["recalculated_with_libreoffice"])
            self.assertIn("What-If data tables", payload["message"])
            self.assertEqual(path.read_bytes(), corrupt_bytes)

            # 3. Default read-only mode writes nothing: data-table pre-check error should warn, not block audit
            valid_path = Path(tmp) / "valid_readonly.xlsx"
            wb_ok = openpyxl.Workbook()
            ws_ok = wb_ok.active
            ws_ok["A1"] = 42
            wb_ok.save(valid_path)
            wb_ok.close()

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None),
                mock.patch.object(
                    recalc_xlsx_mod,
                    "inspect_workbook_xml",
                    side_effect=recalc_xlsx_mod.WorkbookLoadError("unparseable custom sheet XML"),
                ),
            ):
                rc_ro, payload_ro = self._run_main([str(valid_path)])

            self.assertEqual(rc_ro, 0, payload_ro)
            self.assertEqual(payload_ro["status"], "success")
            self.assertEqual(payload_ro["mode"], "check_only")
            self.assertFalse(payload_ro["modified_file"])
            self.assertIn("What-If data tables", payload_ro.get("warning", ""))

    def test_literal_na_in_formula_valid_cached_passes_while_cached_na_and_static_ref_fail(self) -> None:
        import xml.etree.ElementTree as ET

        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ET.register_namespace("", ns)

        def _set_cached_cell(xlsx_path: Path, cell_ref: str, value: str, cell_type: str | None = None) -> None:
            with zipfile.ZipFile(xlsx_path, "r") as zin:
                entries = {name: zin.read(name) for name in zin.namelist()}
            root = ET.fromstring(entries["xl/worksheets/sheet1.xml"])
            for cell in root.iter(f"{{{ns}}}c"):
                if cell.attrib.get("r") == cell_ref:
                    if cell_type is None:
                        cell.attrib.pop("t", None)
                    else:
                        cell.attrib["t"] = cell_type
                    v_el = cell.find(f"{{{ns}}}v")
                    if v_el is None:
                        v_el = ET.SubElement(cell, f"{{{ns}}}v")
                    v_el.text = value
                    break
            entries["xl/worksheets/sheet1.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            with zipfile.ZipFile(xlsx_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in entries.items():
                    zout.writestr(name, data)

        with tempfile.TemporaryDirectory() as tmp:
            # 1. =IF(ISBLANK(B5),#N/A,B5) cached as numeric 42.0 -> PASS, 0 errors
            valid_na_path = Path(tmp) / "valid_na_branch.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "ChartData"
            ws["B5"] = 42.0
            ws["C5"] = "=IF(ISBLANK(B5),#N/A,B5)"
            wb.save(valid_na_path)
            wb.close()
            _set_cached_cell(valid_na_path, "C5", "42.0", cell_type=None)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_ok, payload_ok = self._run_main([str(valid_na_path)])
            self.assertEqual(rc_ok, 0, payload_ok)
            self.assertEqual(payload_ok["status"], "success")
            self.assertEqual(payload_ok["total_errors"], 0)
            self.assertEqual(payload_ok["uncached_formulas"], 0)

            # 2. Same formula cached as #N/A -> FAIL, #N/A detected
            cached_na_path = Path(tmp) / "cached_na_branch.xlsx"
            shutil.copyfile(valid_na_path, cached_na_path)
            _set_cached_cell(cached_na_path, "C5", "#N/A", cell_type="e")

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_na, payload_na = self._run_main([str(cached_na_path)])
            self.assertEqual(rc_na, 1, payload_na)
            self.assertEqual(payload_na["status"], "errors_found")
            self.assertEqual(payload_na["total_errors"], 1)
            self.assertEqual(payload_na["error_summary"]["#N/A"]["locations"], ["ChartData!C5"])

            # 3. =IF(TRUE,1,#REF!) cached as 1 -> FAIL, #REF! always flagged
            static_ref_path = Path(tmp) / "static_ref_branch.xlsx"
            wb3 = openpyxl.Workbook()
            ws3 = wb3.active
            ws3.title = "BrokenRef"
            ws3["A1"] = "=IF(TRUE,1,#REF!)"
            wb3.save(static_ref_path)
            wb3.close()
            _set_cached_cell(static_ref_path, "A1", "1", cell_type=None)

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_ref, payload_ref = self._run_main([str(static_ref_path)])
            self.assertEqual(rc_ref, 1, payload_ref)
            self.assertEqual(payload_ref["status"], "errors_found")
            self.assertEqual(payload_ref["total_errors"], 1)
            self.assertEqual(payload_ref["error_summary"]["#REF!"]["locations"], ["BrokenRef!A1"])

    @_skip_unless_required(
        recalc_xlsx_mod.find_soffice() is None,
        "LibreOffice (soffice) is required for on_disk_uncached_formulas integration test",
    )
    def test_check_only_reports_on_disk_uncached_formulas_and_clears_after_recalc(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "uncached_check_only.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Model"
            ws["A1"] = 25
            ws["A2"] = "=A1*4"
            wb.save(path)
            wb.close()

            orig_bytes = path.read_bytes()

            # 1. Default / --check-only on uncached workbook: temp copy recalculated (uncached_formulas == 0),
            #    but on_disk_uncached_formulas == 1 and warning advises persisting with --in-place or --output.
            for flags in ([], ["--check-only"]):
                with self.subTest(flags=flags):
                    rc, payload = self._run_main([str(path), "30", *flags])
                    self.assertEqual(rc, 0, payload)
                    self.assertEqual(payload["status"], "success")
                    self.assertTrue(payload["recalculated_with_libreoffice"])
                    self.assertFalse(payload["modified_file"])
                    self.assertEqual(payload["uncached_formulas"], 0)
                    self.assertEqual(payload["on_disk_uncached_formulas"], 1)
                    self.assertIn("Input file on disk still has 1 uncached formula(s)", payload.get("warning", ""))
                    self.assertIn("--in-place", payload.get("warning", ""))
                    self.assertIn("--output", payload.get("warning", ""))
                    self.assertEqual(path.read_bytes(), orig_bytes)

            # 2. Persist cached values with --in-place, then re-run --check-only -> on_disk_uncached_formulas == 0, no warning
            rc_ip, payload_ip = self._run_main([str(path), "30", "--in-place"])
            self.assertEqual(rc_ip, 0, payload_ip)
            self.assertEqual(payload_ip["uncached_formulas"], 0)
            self.assertEqual(payload_ip["on_disk_uncached_formulas"], 0)

            rc_after, payload_after = self._run_main([str(path), "30", "--check-only"])
            self.assertEqual(rc_after, 0, payload_after)
            self.assertEqual(payload_after["uncached_formulas"], 0)
            self.assertEqual(payload_after["on_disk_uncached_formulas"], 0)
            self.assertNotIn("Input file on disk still has", payload_after.get("warning", ""))

    @_skip_unless_required(
        recalc_xlsx_mod.find_soffice() is None,
        "LibreOffice (soffice) is required for real What-If data table pre-scan test",
    )
    def test_real_libreoffice_preserves_prescan_data_table_count_and_locations(self) -> None:
        from openpyxl.worksheet.formula import DataTableFormula

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "real_lo_datatable.xlsx"
            out_path = Path(tmp) / "real_lo_datatable_out.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sensitivity"
            ws["A1"] = 10
            ws["A2"] = 2
            ws["B2"] = "=A1*A2"
            ws["C2"] = 10
            ws["D2"] = 20
            ws["B3"] = 2
            ws["B4"] = 3
            ws["C3"] = DataTableFormula(ref="C3:D4", dt2D=True, r1="A1", r2="A2")
            wb.save(path)
            wb.close()

            orig_bytes = path.read_bytes()

            # Both --check-only and --output MUST report data_table_formulas == 1 and data_tables
            # even though LibreOffice converts <f t="dataTable"> into Calc TABLE(...) formulas.
            rc_chk, payload_chk = self._run_main([str(path), "30", "--check-only"])
            self.assertEqual(rc_chk, 0, payload_chk)
            self.assertTrue(payload_chk["recalculated_with_libreoffice"])
            self.assertEqual(payload_chk["data_table_formulas"], 1)
            self.assertEqual(payload_chk["data_tables"], ["Sensitivity!C3:D4"])
            self.assertEqual(path.read_bytes(), orig_bytes)

            rc_out, payload_out = self._run_main([str(path), "30", "--output", str(out_path)])
            self.assertEqual(rc_out, 0, payload_out)
            self.assertTrue(payload_out["recalculated_with_libreoffice"])
            self.assertEqual(payload_out["data_table_formulas"], 1)
            self.assertEqual(payload_out["data_tables"], ["Sensitivity!C3:D4"])
            self.assertTrue(out_path.exists())
            self.assertEqual(path.read_bytes(), orig_bytes)

    def test_plain_text_na_cell_not_flagged_while_error_cell_and_formula_na_fail(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # 1. Plain non-formula text cells ("#N/A", "#DIV/0!") with data_type="s" MUST pass with 0 errors
            text_path = Path(tmp) / "plain_text_na.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "DCF"
            ws["A1"] = "WACC"
            ws["B1"] = 0.09
            ws["A2"] = "Terminal Growth Rate"
            ws["B2"] = 0.025
            ws["A3"] = "PV of Terminal Value"
            ws["B3"] = 600.0
            ws["A4"] = "Enterprise Value"
            ws["B4"] = 1000.0
            ws["A5"] = "Sensitivity Analysis"
            c_na = ws["C1"]
            c_na.value = "#N/A"
            c_na.data_type = "s"
            c_div = ws["C2"]
            c_div.value = "#DIV/0!"
            c_div.data_type = "s"
            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A1"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B1"] = 0.09
            wb.save(text_path)
            wb.close()

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_txt, payload_txt = self._run_main([str(text_path)])
            self.assertEqual(rc_txt, 0, payload_txt)
            self.assertEqual(payload_txt["status"], "success")
            self.assertEqual(payload_txt["total_errors"], 0)

            dcf_txt = validate_dcf_mod.validate_dcf_model(str(text_path))
            self.assertEqual(dcf_txt["status"], "PASS", dcf_txt)
            self.assertEqual(dcf_txt["error_count"], 0, dcf_txt["errors"])

            if recalc_xlsx_mod.find_soffice() is not None:
                rc_lo, payload_lo = self._run_main([str(text_path), "30", "--check-only"])
                self.assertEqual(rc_lo, 0, payload_lo)
                self.assertEqual(payload_lo["total_errors"], 0)

            # 2. Real Excel error cell (data_type == "e") MUST fail
            err_cell_path = Path(tmp) / "real_error_cell.xlsx"
            wb2 = openpyxl.Workbook()
            ws2 = wb2.active
            ws2.title = "Model"
            c_err = ws2["A1"]
            c_err.value = "#N/A"
            c_err.data_type = "e"
            wb2.save(err_cell_path)
            wb2.close()

            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_err, payload_err = self._run_main([str(err_cell_path)])
            self.assertEqual(rc_err, 1, payload_err)
            self.assertEqual(payload_err["status"], "errors_found")
            self.assertEqual(payload_err["error_summary"]["#N/A"]["locations"], ["Model!A1"])

            # 3. Formula evaluating to #N/A via real LibreOffice MUST fail
            if recalc_xlsx_mod.find_soffice() is not None:
                formula_na_path = Path(tmp) / "formula_na.xlsx"
                wb3 = openpyxl.Workbook()
                ws3 = wb3.active
                ws3.title = "Model"
                ws3["A1"] = "=NA()"
                wb3.save(formula_na_path)
                wb3.close()

                rc_fna, payload_fna = self._run_main([str(formula_na_path), "30", "--check-only"])
                self.assertEqual(rc_fna, 1, payload_fna)
                self.assertEqual(payload_fna["status"], "errors_found")
                self.assertEqual(payload_fna["error_summary"]["#N/A"]["locations"], ["Model!A1"])

    def test_scan_workbook_missing_openpyxl_raises_workbook_load_error(self) -> None:
        with mock.patch.dict(sys.modules, {"openpyxl": None}):
            with self.assertRaises(recalc_xlsx_mod.WorkbookLoadError) as ctx:
                recalc_xlsx_mod.scan_workbook_formulas(Path("dummy.xlsx"))
        self.assertIn("openpyxl is required", str(ctx.exception))

    def test_inspect_workbook_xml_matches_scan_workbook_formulas_across_formula_types(self) -> None:
        import xml.etree.ElementTree as ET
        from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ET.register_namespace("", ns)

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "mixed_formulas.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Model"
            ws["A1"] = 10
            ws["A2"] = 20
            ws["A3"] = 30
            # Normal uncached formula
            ws["B1"] = "=A1*2"
            # Cached numeric formula (injected below)
            ws["B2"] = "=A2*2"
            # Cached empty-string formula (<c t="str"><f>...</f><v></v></c>)
            ws["B3"] = '=IF(A1>0,"","Alert")'
            # Cached error formula (<c t="e"><f>1/0</f><v>#DIV/0!</v></c>)
            ws["B4"] = "=1/0"
            # Shared formulas: C1 master (uncached), C2 follower (uncached), C3 follower (cached)
            ws["C1"] = "=A1+1"
            ws["C2"] = "=A2+1"
            ws["C3"] = "=A3+1"
            # ArrayFormula (uncached)
            ws["D1"] = ArrayFormula("D1:D1", "=A1*3")
            # DataTableFormula (uncached)
            ws["E1"] = DataTableFormula(ref="E1:F2", dt2D=True, r1="A1", r2="A2")
            wb.save(path)
            wb.close()

            with zipfile.ZipFile(path, "r") as zin:
                entries = {name: zin.read(name) for name in zin.namelist()}
            root = ET.fromstring(entries["xl/worksheets/sheet1.xml"])
            for cell in root.iter(f"{{{ns}}}c"):
                ref = cell.attrib.get("r")
                f_el = cell.find(f"{{{ns}}}f")
                v_el = cell.find(f"{{{ns}}}v")
                if ref == "B2":
                    cell.attrib.pop("t", None)
                    if v_el is None:
                        v_el = ET.SubElement(cell, f"{{{ns}}}v")
                    v_el.text = "40"
                elif ref == "B3":
                    cell.attrib["t"] = "str"
                    if v_el is None:
                        v_el = ET.SubElement(cell, f"{{{ns}}}v")
                    v_el.text = ""
                elif ref == "B4":
                    cell.attrib["t"] = "e"
                    if v_el is None:
                        v_el = ET.SubElement(cell, f"{{{ns}}}v")
                    v_el.text = "#DIV/0!"
                elif ref == "C1" and f_el is not None:
                    f_el.attrib["t"] = "shared"
                    f_el.attrib["ref"] = "C1:C3"
                    f_el.attrib["si"] = "0"
                elif ref == "C2" and f_el is not None:
                    f_el.attrib["t"] = "shared"
                    f_el.attrib["si"] = "0"
                    f_el.text = None
                elif ref == "C3" and f_el is not None:
                    f_el.attrib["t"] = "shared"
                    f_el.attrib["si"] = "0"
                    f_el.text = None
                    cell.attrib.pop("t", None)
                    if v_el is None:
                        v_el = ET.SubElement(cell, f"{{{ns}}}v")
                    v_el.text = "31"

            entries["xl/worksheets/sheet1.xml"] = ET.tostring(
                root, encoding="utf-8", xml_declaration=True
            )
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in entries.items():
                    zout.writestr(name, data)

            xml_info = xlsx_scan_mod.inspect_workbook_xml(path)
            scan_info = xlsx_scan_mod.scan_workbook_formulas(path)

            # Uncached formulas: B1 (normal), C1 (shared master), C2 (shared follower), D1 (array), E1 (dataTable) = 5
            # Cached formulas: B2 (numeric 40), B3 (cached ""), B4 (cached #DIV/0!), C3 (shared follower 31) = 4
            self.assertEqual(xml_info["uncached_formulas"], 5, xml_info)
            self.assertEqual(xml_info["uncached_formulas"], scan_info["uncached_formulas"])
            self.assertEqual(xml_info["data_tables"], ["Model!E1:F2"])
            self.assertTrue(xml_info["lossy_features_inspected"])
            self.assertEqual(xml_info["lossy_features"], [])
            self.assertEqual(xml_info["lossy_feature_details"], {})

            # Non-OOXML legacy .xls returns empty data_tables, lossy_features, lossy_feature_details,
            # lossy_features_inspected=False, and uncached_formulas=None
            legacy_xls = Path(tmp) / "legacy.xls"
            legacy_xls.write_bytes(b"not-a-zip-legacy-xls")
            self.assertEqual(
                xlsx_scan_mod.inspect_workbook_xml(legacy_xls),
                {
                    "data_tables": [],
                    "uncached_formulas": None,
                    "lossy_features": [],
                    "lossy_feature_details": {},
                    "lossy_features_inspected": False,
                },
            )

            # Corrupt .xlsx (non-zip bytes) MUST fail closed with WorkbookLoadError
            corrupt_xlsx = Path(tmp) / "corrupt.xlsx"
            corrupt_xlsx.write_bytes(b"not-a-zip-archive")
            with self.assertRaises(xlsx_scan_mod.WorkbookLoadError):
                xlsx_scan_mod.inspect_workbook_xml(corrupt_xlsx)

    @_skip_unless_required(
        recalc_xlsx_mod.find_soffice() is None,
        "LibreOffice (soffice) is required for single-scan check-only regression test",
    )
    def test_check_only_with_libreoffice_calls_scan_workbook_once_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "single_scan_check.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Model"
            ws["A1"] = 10
            ws["A2"] = "=A1*5"
            ws["A3"] = "=A2+7"
            wb.save(path)
            wb.close()

            orig_bytes = path.read_bytes()
            real_scan = recalc_xlsx_mod.scan_workbook_formulas

            for flags in ([], ["--check-only"]):
                with self.subTest(flags=flags):
                    with mock.patch.object(
                        recalc_xlsx_mod, "scan_workbook_formulas", wraps=real_scan
                    ) as spy_scan:
                        rc, payload = self._run_main([str(path), "30", *flags])

                    self.assertEqual(rc, 0, payload)
                    self.assertEqual(payload["status"], "success")
                    self.assertTrue(payload["recalculated_with_libreoffice"])
                    self.assertFalse(payload["modified_file"])
                    self.assertEqual(payload["uncached_formulas"], 0)
                    self.assertEqual(payload["on_disk_uncached_formulas"], 2)
                    self.assertIn(
                        "Input file on disk still has 2 uncached formula(s)",
                        payload.get("warning", ""),
                    )
                    # Must call scan_workbook_formulas() EXACTLY ONCE (on LibreOffice temp copy, NEVER on input path)
                    self.assertEqual(spy_scan.call_count, 1)
                    scanned_target = spy_scan.call_args[0][0]
                    self.assertNotEqual(Path(scanned_target).resolve(), path.resolve())
                    self.assertEqual(path.read_bytes(), orig_bytes)

    def test_text_cells_starting_with_equals_not_treated_as_formulas_or_errors(self) -> None:
        import xml.etree.ElementTree as ET

        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        ET.register_namespace("", ns)

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "equals_text_cells.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A1"] = "WACC"
            ws_dcf["B1"] = 0.09
            ws_dcf["A2"] = "Terminal Growth Rate"
            ws_dcf["B2"] = 0.025
            ws_dcf["A3"] = "PV of Terminal Value"
            ws_dcf["B3"] = 600.0
            ws_dcf["A4"] = "Enterprise Value"
            ws_dcf["B4"] = 1000.0
            ws_dcf["A6"] = "Sensitivity Analysis"
            # Placeholders rewritten below as inlineStr and shared string (t="s")
            ws_dcf["D1"] = "inline_placeholder"
            ws_dcf["D2"] = "shared_placeholder"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A1"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B1"] = 0.09
            wb.save(path)
            wb.close()

            with zipfile.ZipFile(path, "r") as zin:
                entries = {name: zin.read(name) for name in zin.namelist()}

            root = ET.fromstring(entries["xl/worksheets/sheet1.xml"])
            for cell in root.iter(f"{{{ns}}}c"):
                ref = cell.attrib.get("r")
                if ref == "D1":
                    cell.attrib["t"] = "inlineStr"
                    for child in list(cell):
                        cell.remove(child)
                    is_el = ET.SubElement(cell, f"{{{ns}}}is")
                    t_el = ET.SubElement(is_el, f"{{{ns}}}t")
                    t_el.text = "== Notes: old link was #REF! =="
                elif ref == "D2":
                    cell.attrib["t"] = "s"
                    for child in list(cell):
                        cell.remove(child)
                    v_el = ET.SubElement(cell, f"{{{ns}}}v")
                    v_el.text = "0"
            entries["xl/worksheets/sheet1.xml"] = ET.tostring(
                root, encoding="utf-8", xml_declaration=True
            )
            entries["xl/sharedStrings.xml"] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">'
                "<si><t>== Notes: old link was #REF! ==</t></si></sst>"
            ).encode("utf-8")

            ct_root = ET.fromstring(entries["[Content_Types].xml"])
            ET.SubElement(
                ct_root,
                f"{{{ct_ns}}}Override",
                {
                    "PartName": "/xl/sharedStrings.xml",
                    "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                },
            )
            entries["[Content_Types].xml"] = ET.tostring(
                ct_root, encoding="utf-8", xml_declaration=True
            )

            rels_root = ET.fromstring(entries["xl/_rels/workbook.xml.rels"])
            ET.SubElement(
                rels_root,
                f"{{{rel_ns}}}Relationship",
                {
                    "Id": "rId99",
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                    "Target": "sharedStrings.xml",
                },
            )
            entries["xl/_rels/workbook.xml.rels"] = ET.tostring(
                rels_root, encoding="utf-8", xml_declaration=True
            )

            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in entries.items():
                    zout.writestr(name, data)

            # 1. inspect_workbook_xml and scan_workbook_formulas both count 0 formulas and 0 errors
            xml_info = xlsx_scan_mod.inspect_workbook_xml(path)
            self.assertEqual(xml_info["uncached_formulas"], 0, xml_info)

            scan_info = xlsx_scan_mod.scan_workbook_formulas(path)
            self.assertEqual(scan_info["total_formulas"], 0, scan_info)
            self.assertEqual(scan_info["uncached_formulas"], 0, scan_info)
            self.assertEqual(scan_info["total_errors"], 0, scan_info)

            # 2. recalc_xlsx.py (scan-only and real LibreOffice --check-only) reports 0 formulas, 0 errors
            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_no_lo, payload_no_lo = self._run_main([str(path)])
            self.assertEqual(rc_no_lo, 0, payload_no_lo)
            self.assertEqual(payload_no_lo["status"], "success")
            self.assertEqual(payload_no_lo["total_formulas"], 0)
            self.assertEqual(payload_no_lo["total_errors"], 0)

            if recalc_xlsx_mod.find_soffice() is not None:
                rc_lo, payload_lo = self._run_main([str(path), "30", "--check-only"])
                self.assertEqual(rc_lo, 0, payload_lo)
                self.assertEqual(payload_lo["status"], "success")
                self.assertEqual(payload_lo["total_formulas"], 0)
                self.assertEqual(payload_lo["total_errors"], 0)

            # 3. validate_dcf.py reports PASS, 0 errors, Total formulas: 0
            dcf_res = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(dcf_res["status"], "PASS", dcf_res)
            self.assertEqual(dcf_res["error_count"], 0, dcf_res["errors"])
            self.assertIn("Total formulas: 0", dcf_res["info"])

    def test_lossy_excel_features_refuse_in_place_and_warn_on_check_only_and_output(self) -> None:
        import xml.etree.ElementTree as ET
        from openpyxl.chart import BarChart, Reference

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Real openpyxl chart on a worksheet -> detects charts + drawings_or_shapes
            chart_path = Path(tmp) / "with_chart.xlsx"
            out_chart_path = Path(tmp) / "with_chart_out.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Model"
            ws["A1"] = 10
            ws["A2"] = "=A1*2"
            chart = BarChart()
            chart.add_data(Reference(ws, min_col=1, min_row=1, max_row=2))
            ws.add_chart(chart, "C1")
            wb.save(chart_path)
            wb.close()

            orig_chart_bytes = chart_path.read_bytes()
            xml_chart = xlsx_scan_mod.inspect_workbook_xml(chart_path)
            self.assertEqual(xml_chart["lossy_features"], ["charts", "drawings_or_shapes"])

            # Refuses --in-place AND --output pointing to the same file without invoking LibreOffice
            for argv in ([str(chart_path), "--in-place"], [str(chart_path), "--output", str(chart_path)]):
                with self.subTest(argv=argv):
                    with (
                        mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                        mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
                    ):
                        rc, payload = self._run_main(argv)
                    mock_popen.assert_not_called()
                    self.assertEqual(rc, 1, payload)
                    self.assertEqual(payload["status"], "error")
                    self.assertEqual(payload["lossy_features"], ["charts", "drawings_or_shapes"])
                    self.assertIn("--check-only", payload["message"])
                    self.assertIn("--output", payload["message"])
                    self.assertEqual(chart_path.read_bytes(), orig_chart_bytes)

            def fake_popen(cmd, **kwargs):
                outdir = Path(cmd[cmd.index("--outdir") + 1])
                src_arg = Path(cmd[-1])
                shutil.copyfile(src_arg, outdir / f"{src_arg.stem}.xlsx")
                proc = mock.MagicMock()
                proc.returncode = 0
                proc.communicate.return_value = ("", "")
                return proc

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=fake_popen),
            ):
                # Default check-only succeeds with warning and leaves input untouched
                rc_chk, payload_chk = self._run_main([str(chart_path)])
                self.assertEqual(rc_chk, 0, payload_chk)
                self.assertEqual(payload_chk["status"], "success")
                self.assertEqual(payload_chk["lossy_features"], ["charts", "drawings_or_shapes"])
                self.assertIn("charts, drawings_or_shapes", payload_chk.get("warning", ""))
                self.assertEqual(chart_path.read_bytes(), orig_chart_bytes)

                # Separate --output succeeds with warning and leaves input untouched
                rc_out, payload_out = self._run_main([str(chart_path), "--output", str(out_chart_path)])
                self.assertEqual(rc_out, 0, payload_out)
                self.assertEqual(payload_out["status"], "success")
                self.assertEqual(payload_out["lossy_features"], ["charts", "drawings_or_shapes"])
                self.assertIn("charts, drawings_or_shapes", payload_out.get("warning", ""))
                self.assertTrue(out_chart_path.exists())
                self.assertEqual(chart_path.read_bytes(), orig_chart_bytes)

            # 2. Pivot table, slicer/timeline (including caches), drawing XML, VBA macro, form control, and sparklines
            feature_cases = [
                ("pivot_tables", {"xl/pivotTables/pivotTable1.xml": b"<pivotTableDefinition/>"}, None),
                ("slicers_or_timelines", {"xl/slicers/slicer1.xml": b"<slicers/>"}, None),
                ("slicers_or_timelines", {"xl/slicerCaches/slicerCache1.xml": b"<slicerCacheDefinition/>"}, None),
                ("slicers_or_timelines", {"xl/timelineCaches/timelineCache1.xml": b"<timelineCacheDefinition/>"}, None),
                (
                    "drawings_or_shapes",
                    {
                        "xl/drawings/drawing1.xml": (
                            b'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
                            b"<xdr:twoCellAnchor/></xdr:wsDr>"
                        )
                    },
                    None,
                ),
                ("vba_macros", {"xl/vbaProject.bin": b"VBA"}, None),
                ("form_or_activex_controls", {"xl/ctrlProps/ctrlProp1.xml": b"<formControlPr/>"}, None),
                ("sparklines", {}, "sparklines"),
            ]
            for idx, (expected_feature, extra_entries, xml_mode) in enumerate(feature_cases):
                with self.subTest(feature=expected_feature, extra_entries=list(extra_entries)):
                    fpath = Path(tmp) / f"feat_{idx}_{expected_feature}.xlsx"
                    wb_f = openpyxl.Workbook()
                    ws_f = wb_f.active
                    ws_f["A1"] = 1
                    wb_f.save(fpath)
                    wb_f.close()

                    with zipfile.ZipFile(fpath, "r") as zin:
                        entries = {name: zin.read(name) for name in zin.namelist()}
                    for k, v in extra_entries.items():
                        entries[k] = v
                    if xml_mode == "sparklines":
                        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                        x14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                        ET.register_namespace("", ns)
                        root = ET.fromstring(entries["xl/worksheets/sheet1.xml"])
                        ext_lst = ET.SubElement(root, f"{{{ns}}}extLst")
                        ext = ET.SubElement(ext_lst, f"{{{ns}}}ext", {"uri": "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"})
                        ET.SubElement(ext, f"{{{x14}}}sparklineGroups")
                        entries["xl/worksheets/sheet1.xml"] = ET.tostring(
                            root, encoding="utf-8", xml_declaration=True
                        )
                    with zipfile.ZipFile(fpath, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                        for name, data in entries.items():
                            zout.writestr(name, data)

                    self.assertEqual(
                        xlsx_scan_mod.inspect_workbook_xml(fpath)["lossy_features"],
                        [expected_feature],
                    )
                    with (
                        mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                        mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
                    ):
                        rc_feat, payload_feat = self._run_main([str(fpath), "--in-place"])
                    mock_popen.assert_not_called()
                    self.assertEqual(rc_feat, 1, payload_feat)
                    self.assertEqual(payload_feat["lossy_features"], [expected_feature])

            # 3. Legacy comment VML (commentsDrawing*.vml / vmlDrawing*.vml) MUST NOT trigger lossy_features
            comment_path = Path(tmp) / "comment_vml_only.xlsx"
            wb_c = openpyxl.Workbook()
            ws_c = wb_c.active
            ws_c["A1"] = 42
            ws_c["A1"].comment = openpyxl.comments.Comment("Source: 10-K", "Analyst")
            wb_c.save(comment_path)
            wb_c.close()
            self.assertEqual(xlsx_scan_mod.inspect_workbook_xml(comment_path)["lossy_features"], [])

    def test_output_error_paths_include_uniform_source_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            src_path = Path(tmp) / "input.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws["A1"] = 10
            wb.save(src_path)
            wb.close()

            expected_src = str(src_path.resolve())

            # 1. Missing --output parent directory
            missing_dir_out = Path(tmp) / "does_not_exist" / "out.xlsx"
            rc_dir, payload_dir = self._run_main([str(src_path), "--output", str(missing_dir_out)])
            self.assertEqual(rc_dir, 1, payload_dir)
            self.assertEqual(payload_dir["status"], "error")
            self.assertEqual(payload_dir.get("source_file"), expected_src)
            self.assertEqual(Path(payload_dir["file"]), missing_dir_out.resolve())

            # 2. LibreOffice conversion failure under --output
            out_path = Path(tmp) / "recalced.xlsx"
            fail_proc = mock.MagicMock()
            fail_proc.returncode = 1
            fail_proc.communicate.return_value = ("", "soffice crashed")
            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", return_value=fail_proc),
            ):
                rc_lo, payload_lo = self._run_main([str(src_path), "--output", str(out_path)])
            self.assertEqual(rc_lo, 1, payload_lo)
            self.assertEqual(payload_lo["status"], "error")
            self.assertEqual(payload_lo.get("source_file"), expected_src)
            self.assertEqual(Path(payload_lo["file"]), out_path.resolve())

            # 3. atomic_copy OSError under --output
            def ok_popen(cmd, **kwargs):
                outdir = Path(cmd[cmd.index("--outdir") + 1])
                shutil.copyfile(src_path, outdir / f"{src_path.stem}.xlsx")
                proc = mock.MagicMock()
                proc.returncode = 0
                proc.communicate.return_value = ("", "")
                return proc

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=ok_popen),
                mock.patch.object(
                    recalc_xlsx_mod,
                    "atomic_copy",
                    side_effect=PermissionError("Read-only target"),
                ),
            ):
                rc_copy, payload_copy = self._run_main([str(src_path), "--output", str(out_path)])
            self.assertEqual(rc_copy, 1, payload_copy)
            self.assertEqual(payload_copy["status"], "error")
            self.assertEqual(payload_copy.get("source_file"), expected_src)
            self.assertEqual(Path(payload_copy["file"]), out_path.resolve())

            # 4. Post-recalc scan_workbook_formulas WorkbookLoadError under --output
            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=ok_popen),
                mock.patch.object(
                    recalc_xlsx_mod,
                    "scan_workbook_formulas",
                    side_effect=recalc_xlsx_mod.WorkbookLoadError("Corrupt output"),
                ),
            ):
                rc_scan, payload_scan = self._run_main([str(src_path), "--output", str(out_path)])
            self.assertEqual(rc_scan, 1, payload_scan)
            self.assertEqual(payload_scan["status"], "error")
            self.assertEqual(payload_scan.get("source_file"), expected_src)
            self.assertEqual(Path(payload_scan["file"]), out_path.resolve())
            self.assertTrue(payload_scan["recalculated_with_libreoffice"])

    def test_defined_name_point_captures_non_target_sheet_without_loading_full_grid(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "large_data_point_capture.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A1"] = "WACC"
            ws_dcf["B1"] = 0.09
            ws_dcf["A2"] = "Terminal Growth Rate"
            ws_dcf["B2"] = 0.025

            # Large non-target sheet (2,000 rows x 10 cols = 20,000 cells) with one valued defined name
            # (Data!$A$1), one blank defined name (Data!$B$1), and a distant error cell at row 1500
            ws_data = wb.create_sheet("Data")
            ws_data.append([123.45, None, "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10"])
            for r in range(2, 2001):
                if r == 1500:
                    ws_data.append(["#DIV/0!", r, r + 1, r + 2, r + 3, r + 4, r + 5, r + 6, r + 7, r + 8])
                else:
                    ws_data.append([r, r + 1, r + 2, r + 3, r + 4, r + 5, r + 6, r + 7, r + 8, r + 9])

            # Non-target sheet whose ONLY targeted defined name points to a blank cell
            ws_blank = wb.create_sheet("BlankTarget")
            ws_blank["A1"] = "Unrelated label"

            # Completely unreferenced non-target sheet
            ws_unused = wb.create_sheet("Unused")
            ws_unused["A1"] = 999

            wb.defined_names["LastPrice"] = DefinedName("LastPrice", attr_text="Data!$A$1")
            wb.defined_names["EmptyRef"] = DefinedName("EmptyRef", attr_text="Data!$B$1")
            wb.defined_names["BlankOnly"] = DefinedName("BlankOnly", attr_text="BlankTarget!$C$4")
            wb.save(path)
            wb.close()

            scan = xlsx_scan_mod.scan_workbook_formulas(
                path,
                capture_sheet_grids=True,
                grid_sheet_names=["DCF"],
                capture_defined_names=True,
            )

            # 1. Target sheet ("DCF") captures its full grid
            self.assertIn("DCF", scan["sheet_grids"])
            dcf_cells = sum(len(row) for row in scan["sheet_grids"]["DCF"].values())
            self.assertEqual(dcf_cells, 4)

            # 2. Point-captured "Data" sheet stores ONLY the 1 non-empty named cell (A1), NOT 19,999+ cells
            self.assertIn("Data", scan["sheet_grids"])
            self.assertEqual(scan["sheet_grids"]["Data"], {1: {1: (123.45, "A1")}})

            # 3. Point-captured sheet with only a blank named cell stores {} so empty-cell detection works
            self.assertIn("BlankTarget", scan["sheet_grids"])
            self.assertEqual(scan["sheet_grids"]["BlankTarget"], {})
            defined_lower = {k.lower(): v for k, v in scan["defined_names"].items()}
            self.assertEqual(
                validate_dcf_mod._find_invalid_named_metric(
                    defined_lower, ("emptyref",), scan["sheet_grids"]
                ),
                (None, "Data!B1"),
            )
            self.assertEqual(
                validate_dcf_mod._find_invalid_named_metric(
                    defined_lower, ("blankonly",), scan["sheet_grids"]
                ),
                (None, "BlankTarget!C4"),
            )

            # 4. Unreferenced non-target sheet is not captured in sheet_grids at all
            self.assertNotIn("Unused", scan["sheet_grids"])

            # 5. Formula/error scanning STILL scans every cell of non-target sheets (row 1500 error caught)
            self.assertEqual(scan["total_errors"], 1)
            self.assertEqual(scan["error_summary"]["#DIV/0!"]["locations"], ["Data!A1500"])

    def test_allowlist_ooxml_parts_and_worksheet_extlst_detection_and_clean_double_inplace(self) -> None:
        import xml.etree.ElementTree as ET
        from openpyxl.comments import Comment
        from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
        from openpyxl.styles import PatternFill
        from openpyxl.worksheet.datavalidation import DataValidation

        def _fake_popen(cmd, **kwargs):
            outdir = Path(cmd[cmd.index("--outdir") + 1])
            src_arg = Path(cmd[-1])
            shutil.copyfile(src_arg, outdir / f"{src_arg.stem}.xlsx")
            proc = mock.MagicMock()
            proc.returncode = 0
            proc.communicate.return_value = ("", "")
            return proc

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Verify allowlist + worksheet <extLst> URI detection catches all lossy categories
            lossy_cases = [
                ("threaded_comments", {"xl/threadedComments/threadedComment1.xml": b"<ThreadedComments/>"}, None),
                ("threaded_comments", {"xl/persons/person.xml": b"<personList/>"}, None),
                ("connections_or_power_query", {"xl/connections.xml": b"<connections/>"}, None),
                ("connections_or_power_query", {"xl/queryTables/queryTable1.xml": b"<queryTable/>"}, None),
                ("connections_or_power_query", {"xl/model/item.data": b"MODEL"}, None),
                ("connections_or_power_query", {"customXml/item1.xml": b"<root/>"}, None),
                ("embedded_media_or_objects", {"xl/media/image1.png": b"\x89PNG"}, None),
                ("embedded_media_or_objects", {"xl/embeddings/oleObject1.bin": b"OLE"}, None),
                ("excel_metadata_or_rich_data", {"xl/metadata.xml": b"<metadata/>"}, None),
                ("excel_metadata_or_rich_data", {"xl/richData/rdrichvalue.xml": b"<rvData/>"}, None),
                ("external_links", {"xl/externalLinks/externalLink1.xml": b"<externalLink/>"}, None),
                ("x14_conditional_formatting", {}, "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"),
                ("x14_data_validation", {}, "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"),
                ("slicers_or_timelines", {}, "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}"),
                ("worksheet_extensions", {}, "{DEADBEEF-0000-0000-0000-000000000000}"),
                ("unsupported_ooxml_parts", {"xl/webExtensions/webExtension1.xml": b"<webextension/>"}, None),
            ]
            for idx, (expected_feature, extra_entries, ext_uri) in enumerate(lossy_cases):
                with self.subTest(feature=expected_feature, extra=list(extra_entries), ext_uri=ext_uri):
                    fpath = Path(tmp) / f"allowlist_{idx}_{expected_feature}.xlsx"
                    out_path = Path(tmp) / f"allowlist_{idx}_out.xlsx"
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws["A1"] = 10
                    wb.save(fpath)
                    wb.close()

                    with zipfile.ZipFile(fpath, "r") as zin:
                        entries = {name: zin.read(name) for name in zin.namelist()}
                    for k, v in extra_entries.items():
                        entries[k] = v
                    if ext_uri is not None:
                        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                        ET.register_namespace("", ns)
                        root = ET.fromstring(entries["xl/worksheets/sheet1.xml"])
                        ext_lst = ET.SubElement(root, f"{{{ns}}}extLst")
                        ET.SubElement(ext_lst, f"{{{ns}}}ext", {"uri": ext_uri})
                        entries["xl/worksheets/sheet1.xml"] = ET.tostring(
                            root, encoding="utf-8", xml_declaration=True
                        )
                    with zipfile.ZipFile(fpath, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                        for name, data in entries.items():
                            zout.writestr(name, data)

                    self.assertEqual(
                        xlsx_scan_mod.inspect_workbook_xml(fpath)["lossy_features"],
                        [expected_feature],
                    )
                    # Refuses --in-place without spawning LibreOffice
                    with (
                        mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                        mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
                    ):
                        rc_ip, payload_ip = self._run_main([str(fpath), "--in-place"])
                    mock_popen.assert_not_called()
                    self.assertEqual(rc_ip, 1, payload_ip)
                    self.assertEqual(payload_ip["status"], "error")
                    self.assertEqual(payload_ip["lossy_features"], [expected_feature])

                    # Warns cleanly on --check-only and --output
                    with (
                        mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                        mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=_fake_popen),
                    ):
                        rc_chk, payload_chk = self._run_main([str(fpath), "--check-only"])
                        self.assertEqual(rc_chk, 0, payload_chk)
                        self.assertEqual(payload_chk["lossy_features"], [expected_feature])
                        self.assertIn(expected_feature, payload_chk.get("warning", ""))

                        rc_out, payload_out = self._run_main([str(fpath), "--output", str(out_path)])
                        self.assertEqual(rc_out, 0, payload_out)
                        self.assertEqual(payload_out["lossy_features"], [expected_feature])
                        self.assertIn(expected_feature, payload_out.get("warning", ""))

            # 2. Safe core OOXML parts (including directory entries, calcChain, printerSettings, custom.xml)
            #    MUST report lossy_features == []
            safe_path = Path(tmp) / "safe_parts.xlsx"
            wb_s = openpyxl.Workbook()
            ws_s = wb_s.active
            ws_s["A1"] = 1
            wb_s.save(safe_path)
            wb_s.close()
            with zipfile.ZipFile(safe_path, "r") as zin:
                safe_entries = {name: zin.read(name) for name in zin.namelist()}
            safe_entries["xl/calcChain.xml"] = b"<calcChain/>"
            safe_entries["xl/printerSettings/printerSettings1.bin"] = b"PRINTER"
            safe_entries["docProps/custom.xml"] = b"<Properties/>"
            with zipfile.ZipFile(safe_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                zout.writestr("xl/worksheets/", b"")
                for name, data in safe_entries.items():
                    zout.writestr(name, data)
            self.assertEqual(xlsx_scan_mod.inspect_workbook_xml(safe_path)["lossy_features"], [])

            # 3. Clean build_dcf_skeleton.py with cell comments, data validation, and classic conditional
            #    formatting passes recalc_xlsx.py --in-place TWICE in a row with real LibreOffice and
            #    reports lossy_features == [] before, after #1, and after #2.
            if recalc_xlsx_mod.find_soffice() is not None:
                skel_path = Path(tmp) / "clean_double_inplace.xlsx"
                build_proc = subprocess.run(
                    [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(skel_path), "--case", "2"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

                wb_rich = openpyxl.load_workbook(skel_path)
                ws_dcf = wb_rich["DCF"]
                ws_dcf["H4"] = "Audit Note"
                ws_dcf["H4"].comment = Comment("Verified clean comment", "Analyst")
                dv = DataValidation(type="list", formula1='"Bear,Base,Bull"', allow_blank=True)
                ws_dcf.add_data_validation(dv)
                dv.add("H5")
                fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                ws_dcf.conditional_formatting.add(
                    "C74:G74", CellIsRule(operator="greaterThan", formula=["0"], fill=fill)
                )
                ws_dcf.conditional_formatting.add(
                    "C86:G86",
                    ColorScaleRule(start_type="min", start_color="F8696B", end_type="max", end_color="63BE7B"),
                )
                ws_dcf.conditional_formatting.add(
                    "C89:G89", FormulaRule(formula=["C89>0"], fill=fill)
                )
                wb_rich.save(skel_path)
                wb_rich.close()

                self.assertEqual(xlsx_scan_mod.inspect_workbook_xml(skel_path)["lossy_features"], [])

                for run_idx in (1, 2):
                    with self.subTest(inplace_run=run_idx):
                        rc_r, payload_r = self._run_main([str(skel_path), "30", "--in-place"])
                        self.assertEqual(rc_r, 0, payload_r)
                        self.assertEqual(payload_r["status"], "success", payload_r)
                        self.assertNotIn("lossy_features", payload_r)
                        self.assertEqual(
                            xlsx_scan_mod.inspect_workbook_xml(skel_path)["lossy_features"],
                            [],
                        )

    def test_autofilter_empty_wsdr_stub_allows_double_inplace_while_nonempty_and_malformed_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # 1. Real LibreOffice 24.2 test: clean workbook with AutoFilter passes --in-place TWICE
            #    even though LibreOffice writes an empty <xdr:wsDr/> stub at xl/drawings/drawing1.xml.
            if recalc_xlsx_mod.find_soffice() is not None:
                af_path = Path(tmp) / "autofilter_clean.xlsx"
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Data"
                ws["A1"] = "Item"
                ws["B1"] = "Value"
                ws["A2"] = "Alpha"
                ws["B2"] = "=20+5"
                ws.auto_filter.ref = "A1:B2"
                wb.save(af_path)
                wb.close()

                for pass_num in (1, 2):
                    with self.subTest(autofilter_inplace_pass=pass_num):
                        rc, payload = self._run_main([str(af_path), "30", "--in-place"])
                        self.assertEqual(rc, 0, payload)
                        self.assertEqual(payload["status"], "success", payload)
                        self.assertEqual(payload["mode"], "in_place")
                        self.assertTrue(payload["recalculated_with_libreoffice"])
                        self.assertEqual(payload["total_errors"], 0)
                        self.assertEqual(payload["uncached_formulas"], 0)
                        self.assertNotIn("lossy_features", payload)
                        self.assertEqual(
                            xlsx_scan_mod.inspect_workbook_xml(af_path)["lossy_features"],
                            [],
                        )

                with zipfile.ZipFile(af_path, "r") as zf:
                    self.assertIn("xl/drawings/drawing1.xml", zf.namelist())

                wb_check = openpyxl.load_workbook(af_path, data_only=True)
                try:
                    ws_check = wb_check["Data"]
                    self.assertEqual(ws_check.auto_filter.ref, "A1:B2")
                    self.assertEqual(ws_check["B2"].value, 25)
                finally:
                    wb_check.close()

            def _make_with_extra(name: str, extra: dict[str, bytes]) -> Path:
                p = Path(tmp) / name
                wb_s = openpyxl.Workbook()
                ws_s = wb_s.active
                ws_s["A1"] = 10
                wb_s.save(p)
                wb_s.close()
                with zipfile.ZipFile(p, "r") as zin:
                    entries = {n: zin.read(n) for n in zin.namelist()}
                entries.update(extra)
                with zipfile.ZipFile(p, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                    for n, data in entries.items():
                        zout.writestr(n, data)
                return p

            # 2. Empty <xdr:wsDr/> stub + empty drawing .rels -> lossy_features == []
            empty_stub_path = _make_with_extra(
                "empty_wsdr_and_rels.xlsx",
                {
                    "xl/drawings/drawing1.xml": (
                        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        b'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
                        b"</xdr:wsDr>"
                    ),
                    "xl/drawings/_rels/drawing1.xml.rels": (
                        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
                    ),
                },
            )
            self.assertEqual(xlsx_scan_mod.inspect_workbook_xml(empty_stub_path)["lossy_features"], [])

            # 3. Non-empty drawing anchor STILL refuses --in-place with drawings_or_shapes
            nonempty_path = _make_with_extra(
                "nonempty_wsdr.xlsx",
                {
                    "xl/drawings/drawing1.xml": (
                        b'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
                        b"<xdr:twoCellAnchor/></xdr:wsDr>"
                    )
                },
            )
            self.assertEqual(
                xlsx_scan_mod.inspect_workbook_xml(nonempty_path)["lossy_features"],
                ["drawings_or_shapes"],
            )
            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
            ):
                rc_ne, payload_ne = self._run_main([str(nonempty_path), "--in-place"])
            mock_popen.assert_not_called()
            self.assertEqual(rc_ne, 1, payload_ne)
            self.assertEqual(payload_ne["status"], "error")
            self.assertEqual(payload_ne["lossy_features"], ["drawings_or_shapes"])

            # 4. Non-wsDr root and non-empty drawing .rels STILL flag drawings_or_shapes
            non_wsdr_path = _make_with_extra(
                "non_wsdr_root.xlsx",
                {"xl/drawings/drawing1.xml": b"<customDrawingRoot/>"},
            )
            self.assertEqual(
                xlsx_scan_mod.inspect_workbook_xml(non_wsdr_path)["lossy_features"],
                ["drawings_or_shapes"],
            )
            nonempty_rels_path = _make_with_extra(
                "nonempty_drawing_rels.xlsx",
                {
                    "xl/drawings/drawing1.xml": (
                        b'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>'
                    ),
                    "xl/drawings/_rels/drawing1.xml.rels": (
                        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                        b'Target="../media/image1.png"/></Relationships>'
                    ),
                },
            )
            self.assertEqual(
                xlsx_scan_mod.inspect_workbook_xml(nonempty_rels_path)["lossy_features"],
                ["drawings_or_shapes"],
            )

            # 5. Malformed drawing XML and malformed drawing .rels STILL fail closed
            malformed_drawing_path = _make_with_extra(
                "malformed_drawing.xlsx",
                {"xl/drawings/drawing1.xml": b"<xdr:wsDr><unclosed"},
            )
            with self.assertRaises(xlsx_scan_mod.WorkbookLoadError):
                xlsx_scan_mod.inspect_workbook_xml(malformed_drawing_path)
            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen2,
            ):
                rc_mal, payload_mal = self._run_main([str(malformed_drawing_path), "--in-place"])
            mock_popen2.assert_not_called()
            self.assertEqual(rc_mal, 1, payload_mal)
            self.assertEqual(payload_mal["status"], "error")

            malformed_rels_path = _make_with_extra(
                "malformed_drawing_rels.xlsx",
                {"xl/drawings/_rels/drawing1.xml.rels": b"<Relationships><broken"},
            )
            with self.assertRaises(xlsx_scan_mod.WorkbookLoadError):
                xlsx_scan_mod.inspect_workbook_xml(malformed_rels_path)

    def test_excel_table_allows_double_inplace_while_query_backed_table_remains_flagged(self) -> None:
        from openpyxl.worksheet.table import Table, TableStyleInfo

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Standard openpyxl Excel Table (xl/tables/table1.xml) with formulas passes
            #    recalc_xlsx.py --in-place TWICE with real LibreOffice, preserving table intact.
            table_path = Path(tmp) / "comps_table.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Comps"
            ws.append(["Ticker", "Rev", "DoubleRev"])
            ws.append(["AAA", 100, "=B2*2"])
            ws.append(["BBB", 200, "=B3*2"])
            ws.append(["CCC", 300, "=B4*2"])
            tbl = Table(displayName="Comps", ref="A1:C4")
            tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
            ws.add_table(tbl)
            wb.save(table_path)
            wb.close()

            xml_before = xlsx_scan_mod.inspect_workbook_xml(table_path)
            self.assertEqual(xml_before["lossy_features"], [])
            self.assertEqual(xml_before["lossy_feature_details"], {})

            if recalc_xlsx_mod.find_soffice() is not None:
                for pass_num in (1, 2):
                    with self.subTest(table_inplace_pass=pass_num):
                        rc, payload = self._run_main([str(table_path), "30", "--in-place"])
                        self.assertEqual(rc, 0, payload)
                        self.assertEqual(payload["status"], "success", payload)
                        self.assertEqual(payload["mode"], "in_place")
                        self.assertTrue(payload["recalculated_with_libreoffice"])
                        self.assertEqual(payload["total_errors"], 0)
                        self.assertEqual(payload["uncached_formulas"], 0)
                        self.assertNotIn("lossy_features", payload)
                        self.assertNotIn("lossy_feature_details", payload)
                        xml_after = xlsx_scan_mod.inspect_workbook_xml(table_path)
                        self.assertEqual(xml_after["lossy_features"], [])
                        self.assertEqual(xml_after["lossy_feature_details"], {})

                with zipfile.ZipFile(table_path, "r") as zf:
                    self.assertIn("xl/tables/table1.xml", zf.namelist())

                wb_v = openpyxl.load_workbook(table_path, data_only=True)
                try:
                    ws_v = wb_v["Comps"]
                    self.assertIn("Comps", ws_v.tables)
                    self.assertEqual(ws_v.tables["Comps"].displayName, "Comps")
                    self.assertEqual(ws_v.tables["Comps"].ref, "A1:C4")
                    self.assertEqual([ws_v[f"C{r}"].value for r in (2, 3, 4)], [200, 400, 600])
                finally:
                    wb_v.close()

            # 2. Query-backed table (xl/tables/table1.xml + xl/tables/_rels/table1.xml.rels +
            #    xl/queryTables/queryTable1.xml + xl/connections.xml) STILL flags connections_or_power_query
            #    without flagging xl/tables/ parts.
            query_table_path = Path(tmp) / "query_backed_table.xlsx"
            shutil.copyfile(table_path, query_table_path)
            with zipfile.ZipFile(query_table_path, "r") as zin:
                entries = {name: zin.read(name) for name in zin.namelist()}
            entries["xl/tables/_rels/table1.xml.rels"] = (
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
            )
            entries["xl/queryTables/queryTable1.xml"] = b"<queryTable/>"
            entries["xl/connections.xml"] = b"<connections/>"
            with zipfile.ZipFile(query_table_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in entries.items():
                    zout.writestr(name, data)

            xml_qt = xlsx_scan_mod.inspect_workbook_xml(query_table_path)
            self.assertEqual(xml_qt["lossy_features"], ["connections_or_power_query"])
            self.assertEqual(
                xml_qt["lossy_feature_details"],
                {
                    "connections_or_power_query": [
                        "xl/connections.xml",
                        "xl/queryTables/queryTable1.xml",
                    ]
                },
            )

    def test_lossy_feature_details_exposed_in_refusal_and_warning_json_and_messages(self) -> None:
        import xml.etree.ElementTree as ET

        def _fake_popen(cmd, **kwargs):
            outdir = Path(cmd[cmd.index("--outdir") + 1])
            src_arg = Path(cmd[-1])
            shutil.copyfile(src_arg, outdir / f"{src_arg.stem}.xlsx")
            proc = mock.MagicMock()
            proc.returncode = 0
            proc.communicate.return_value = ("", "")
            return proc

        cases = [
            (
                "unsupported_part",
                "unsupported_ooxml_parts",
                {"xl/webExtensions/webExtension1.xml": b"<webextension/>"},
                None,
                {"unsupported_ooxml_parts": ["xl/webExtensions/webExtension1.xml"]},
                "xl/webExtensions/webExtension1.xml",
            ),
            (
                "x14_cf_ext",
                "x14_conditional_formatting",
                {},
                "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",
                {
                    "x14_conditional_formatting": [
                        "xl/worksheets/sheet1.xml:{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                    ]
                },
                "xl/worksheets/sheet1.xml:{78C0D931-6437-407D-A8EE-F0AAD7539E65}",
            ),
            (
                "unknown_ext_uri",
                "worksheet_extensions",
                {},
                "{DEADBEEF-0000-0000-0000-000000000000}",
                {
                    "worksheet_extensions": [
                        "xl/worksheets/sheet1.xml:{DEADBEEF-0000-0000-0000-000000000000}"
                    ]
                },
                "xl/worksheets/sheet1.xml:{DEADBEEF-0000-0000-0000-000000000000}",
            ),
            (
                "blank_ext_uri",
                "worksheet_extensions",
                {},
                "   ",
                {"worksheet_extensions": ["xl/worksheets/sheet1.xml:<missing-uri>"]},
                "xl/worksheets/sheet1.xml:<missing-uri>",
            ),
        ]

        with tempfile.TemporaryDirectory() as tmp:
            for label, expected_cat, extra_parts, ext_uri, expected_details, expected_token in cases:
                with self.subTest(case=label):
                    fpath = Path(tmp) / f"{label}.xlsx"
                    out_path = Path(tmp) / f"{label}_out.xlsx"
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws["A1"] = 42
                    wb.save(fpath)
                    wb.close()

                    with zipfile.ZipFile(fpath, "r") as zin:
                        entries = {name: zin.read(name) for name in zin.namelist()}
                    entries.update(extra_parts)
                    if ext_uri is not None:
                        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                        ET.register_namespace("", ns)
                        root = ET.fromstring(entries["xl/worksheets/sheet1.xml"])
                        ext_lst = ET.SubElement(root, f"{{{ns}}}extLst")
                        ET.SubElement(ext_lst, f"{{{ns}}}ext", {"uri": ext_uri})
                        entries["xl/worksheets/sheet1.xml"] = ET.tostring(
                            root, encoding="utf-8", xml_declaration=True
                        )
                    with zipfile.ZipFile(fpath, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                        for name, data in entries.items():
                            zout.writestr(name, data)

                    xml_info = xlsx_scan_mod.inspect_workbook_xml(fpath)
                    self.assertEqual(xml_info["lossy_features"], [expected_cat])
                    self.assertEqual(xml_info["lossy_feature_details"], expected_details)

                    # 1. --in-place refusal JSON & message include exact offending part/URI
                    with (
                        mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                        mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
                    ):
                        rc_ip, payload_ip = self._run_main([str(fpath), "--in-place"])
                    mock_popen.assert_not_called()
                    self.assertEqual(rc_ip, 1, payload_ip)
                    self.assertEqual(payload_ip["status"], "error")
                    self.assertEqual(payload_ip["lossy_features"], [expected_cat])
                    self.assertEqual(payload_ip["lossy_feature_details"], expected_details)
                    self.assertIn(expected_token, payload_ip["message"])

                    # 2. --check-only and --output warning JSON & warning text include exact offending part/URI
                    with (
                        mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                        mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=_fake_popen),
                    ):
                        rc_chk, payload_chk = self._run_main([str(fpath), "--check-only"])
                        self.assertEqual(rc_chk, 0, payload_chk)
                        self.assertEqual(payload_chk["lossy_feature_details"], expected_details)
                        self.assertIn(expected_token, payload_chk.get("warning", ""))

                        rc_out, payload_out = self._run_main([str(fpath), "--output", str(out_path)])
                        self.assertEqual(rc_out, 0, payload_out)
                        self.assertEqual(payload_out["lossy_feature_details"], expected_details)
                        self.assertIn(expected_token, payload_out.get("warning", ""))

            # 2b. <ext/> with completely omitted uri attribute is also reported as <missing-uri> on --in-place refusal
            omitted_uri_path = Path(tmp) / "omitted_ext_uri.xlsx"
            wb_o = openpyxl.Workbook()
            wb_o.active["A1"] = 1
            wb_o.save(omitted_uri_path)
            wb_o.close()
            with zipfile.ZipFile(omitted_uri_path, "r") as zin:
                entries = {name: zin.read(name) for name in zin.namelist()}
            ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            ET.register_namespace("", ns)
            root = ET.fromstring(entries["xl/worksheets/sheet1.xml"])
            ext_lst = ET.SubElement(root, f"{{{ns}}}extLst")
            ET.SubElement(ext_lst, f"{{{ns}}}ext")
            entries["xl/worksheets/sheet1.xml"] = ET.tostring(
                root, encoding="utf-8", xml_declaration=True
            )
            with zipfile.ZipFile(omitted_uri_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in entries.items():
                    zout.writestr(name, data)
            self.assertEqual(
                xlsx_scan_mod.inspect_workbook_xml(omitted_uri_path)["lossy_feature_details"],
                {"worksheet_extensions": ["xl/worksheets/sheet1.xml:<missing-uri>"]},
            )
            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
            ):
                rc_o, payload_o = self._run_main([str(omitted_uri_path), "--in-place"])
            mock_popen.assert_not_called()
            self.assertEqual(rc_o, 1, payload_o)
            self.assertEqual(
                payload_o["lossy_feature_details"],
                {"worksheet_extensions": ["xl/worksheets/sheet1.xml:<missing-uri>"]},
            )
            self.assertIn("xl/worksheets/sheet1.xml:<missing-uri>", payload_o["message"])

            # 3. Preview truncation caps human-readable text at 5 items while preserving full sorted details in JSON
            many_path = Path(tmp) / "many_unsupported.xlsx"
            wb_m = openpyxl.Workbook()
            wb_m.active["A1"] = 1
            wb_m.save(many_path)
            wb_m.close()
            with zipfile.ZipFile(many_path, "r") as zin:
                entries = {name: zin.read(name) for name in zin.namelist()}
            expected_parts = [f"xl/webExtensions/webExtension{i}.xml" for i in range(1, 8)]
            for part in expected_parts:
                entries[part] = b"<webextension/>"
            with zipfile.ZipFile(many_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in entries.items():
                    zout.writestr(name, data)

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen") as mock_popen,
            ):
                rc_m, payload_m = self._run_main([str(many_path), "--in-place"])
            mock_popen.assert_not_called()
            self.assertEqual(rc_m, 1, payload_m)
            self.assertEqual(
                payload_m["lossy_feature_details"],
                {"unsupported_ooxml_parts": expected_parts},
            )
            self.assertIn("(+2 more)", payload_m["message"])
            self.assertIn("xl/webExtensions/webExtension1.xml", payload_m["message"])
            self.assertNotIn("xl/webExtensions/webExtension7.xml", payload_m["message"])
            self.assertTrue(payload_m["lossy_features_inspected"])

    def test_non_ooxml_xls_and_ods_report_lossy_features_inspected_false_with_warning(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_p = Path(tmp)
            # Clean .xlsx reference workbook to convert/copy in tests
            clean_xlsx = tmp_p / "clean.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Model"
            ws["A1"] = 10
            ws["A2"] = "=A1*3"
            wb.save(clean_xlsx)
            wb.close()

            # 1. Mocked .xls input -> check-only and --output report lossy_features_inspected=False + warning;
            #    --in-place refuses cleanly with lossy_features_inspected=False.
            legacy_xls = tmp_p / "legacy.xls"
            legacy_xls.write_bytes(b"dummy-binary-xls-bytes")
            out_xls_xlsx = tmp_p / "converted_from_xls.xlsx"

            def fake_popen(cmd, **kwargs):
                outdir = Path(cmd[cmd.index("--outdir") + 1])
                src_arg = Path(cmd[-1])
                shutil.copyfile(clean_xlsx, outdir / f"{src_arg.stem}.xlsx")
                proc = mock.MagicMock()
                proc.returncode = 0
                proc.communicate.return_value = ("", "")
                return proc

            with (
                mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value="/usr/bin/soffice"),
                mock.patch.object(recalc_xlsx_mod.subprocess, "Popen", side_effect=fake_popen),
            ):
                rc_chk, payload_chk = self._run_main([str(legacy_xls)])
                self.assertEqual(rc_chk, 0, payload_chk)
                self.assertIs(payload_chk["lossy_features_inspected"], False)
                self.assertIn("cannot be pre-inspected for lossy Excel features", payload_chk.get("warning", ""))
                self.assertIn(".xls", payload_chk.get("warning", ""))

                rc_out, payload_out = self._run_main([str(legacy_xls), "--output", str(out_xls_xlsx)])
                self.assertEqual(rc_out, 0, payload_out)
                self.assertIs(payload_out["lossy_features_inspected"], False)
                self.assertIn("cannot be pre-inspected for lossy Excel features", payload_out.get("warning", ""))

                rc_ip, payload_ip = self._run_main([str(legacy_xls), "--in-place"])
                self.assertEqual(rc_ip, 1, payload_ip)
                self.assertEqual(payload_ip["status"], "error")
                self.assertIs(payload_ip["lossy_features_inspected"], False)

            # 2. OOXML .xlsx reports lossy_features_inspected=True
            with mock.patch.object(recalc_xlsx_mod, "find_soffice", return_value=None):
                rc_xlsx, payload_xlsx = self._run_main([str(clean_xlsx), "--allow-unrecalculated"])
            self.assertEqual(rc_xlsx, 0, payload_xlsx)
            self.assertIs(payload_xlsx["lossy_features_inspected"], True)

            # 3. Real LibreOffice .ods integration test when soffice is available
            soffice = recalc_xlsx_mod.find_soffice()
            if soffice is not None:
                ods_dir = tmp_p / "ods_src"
                ods_dir.mkdir()
                lo_profile = tmp_p / "lo_ods_profile"
                recalc_xlsx_mod.write_libreoffice_recalc_profile(lo_profile)
                conv = subprocess.run(
                    [
                        soffice,
                        f"-env:UserInstallation={lo_profile.as_uri()}",
                        "--headless",
                        "--norestore",
                        "--nolockcheck",
                        "--convert-to",
                        "ods",
                        "--outdir",
                        str(ods_dir),
                        str(clean_xlsx),
                    ],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(conv.returncode, 0, conv.stdout + conv.stderr)
                ods_path = ods_dir / "clean.ods"
                self.assertTrue(ods_path.exists())

                rc_ods, payload_ods = self._run_main([str(ods_path), "30", "--check-only"])
                self.assertEqual(rc_ods, 0, payload_ods)
                self.assertEqual(payload_ods["status"], "success")
                self.assertIs(payload_ods["lossy_features_inspected"], False)
                self.assertIn(".ods", payload_ods.get("warning", ""))
                self.assertIn("cannot be pre-inspected for lossy Excel features", payload_ods.get("warning", ""))


@_skip_unless_required(openpyxl is None, "openpyxl is required for DCF validation tests")
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
            # Single-tab DCF with valid inline WACC passes without failing on missing separate WACC tab
            self.assertEqual(results["status"], "PASS", results)
            self.assertEqual(results["error_count"], 0, results["errors"])
            self.assertEqual(results["warning_count"], 0, results["warnings"])
            self.assertTrue(
                any("No separate WACC sheet found; using WACC from DCF sheet / named range" in i for i in results["info"]),
                results["info"],
            )
            self.assertFalse(
                any("AttributeError" in w or "has no attribute" in w for w in results["warnings"]),
                results["warnings"],
            )
            # WACC range check still evaluates cleanly from the DCF sheet
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
            self.assertTrue(any("No dedicated WACC sheet found" in w for w in res1["warnings"]), res1["warnings"])

            # 2. Empty DCF + empty WACC sheets (missing critical DCF components)
            empty_tabs = Path(tmp) / "empty_tabs.xlsx"
            wb2 = openpyxl.Workbook()
            wb2.active.title = "DCF"
            wb2.create_sheet("WACC")
            wb2.save(empty_tabs)
            res2 = validate_dcf_mod.validate_dcf_model(str(empty_tabs))
            self.assertEqual(res2["status"], "FAIL", res2)
            self.assertGreaterEqual(res2["error_count"], 3, res2["errors"])

    def test_unrecognized_tv_ev_labels_warn_instead_of_hard_fail(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "custom_tv_ev_labels.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A2"] = "WACC"
            ws_dcf["B2"] = 0.095
            ws_dcf["A3"] = "Terminal Growth Rate"
            ws_dcf["B3"] = 0.025
            # Custom non-standard valuation summary labels without named ranges
            ws_dcf["A6"] = "Custom Perpetuity Tail"
            ws_dcf["B6"] = 600.0
            ws_dcf["A7"] = "Custom Total Firm Value"
            ws_dcf["B7"] = 1000.0
            ws_dcf["A12"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A2"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B2"] = 0.095
            wb.save(path)

            res = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res["status"], "PASS", res)
            self.assertEqual(res["error_count"], 0, res["errors"])
            self.assertTrue(
                any("Could not locate PV of Terminal Value and Enterprise Value" in w for w in res["warnings"]),
                res["warnings"],
            )

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

    def test_tv_proportion_warns_outside_50_to_70_percent_target(self) -> None:
        for tv_val, expect_warning in ((740.0, True), (450.0, True), (650.0, False)):
            with self.subTest(tv_val=tv_val), tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / f"tv_{int(tv_val)}.xlsx"
                self._build_compliant_two_sheet_dcf(path, with_formula=False)
                wb = openpyxl.load_workbook(path)
                wb["DCF"]["B6"] = tv_val  # EV in B7 is 1000.0
                wb.save(path)
                wb.close()

                results = validate_dcf_mod.validate_dcf_model(str(path))
                self.assertEqual(results["status"], "PASS", results)
                if expect_warning:
                    self.assertGreater(results["warning_count"], 0, results)
                    self.assertTrue(
                        any("typically should be 50-70%" in w for w in results["warnings"]),
                        results["warnings"],
                    )
                else:
                    self.assertEqual(results["warning_count"], 0, results["warnings"])

    def test_validate_dcf_counts_data_table_formula_as_uncached(self) -> None:
        from openpyxl.worksheet.formula import DataTableFormula

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "dcf_datatable.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=False)
            wb = openpyxl.load_workbook(path)
            wb["DCF"]["C12"] = DataTableFormula(ref="C12:D13", dt2D=True, r1="B3", r2="B4")
            wb.save(path)
            wb.close()

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "FAIL", results)
            self.assertTrue(any("Total formulas: 1" in i for i in results["info"]), results["info"])
            self.assertTrue(any("Uncached formulas: 1" in i for i in results["info"]), results["info"])

    def test_current_and_market_enterprise_value_rows_do_not_mask_dcf_ev(self) -> None:
        for ev_label in ("Enterprise Value", "Implied Enterprise Value"):
            with self.subTest(ev_label=ev_label), tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / "dcf_with_current_market_ev.xlsx"
                wb = openpyxl.Workbook()
                ws_dcf = wb.active
                ws_dcf.title = "DCF"
                # Section 1 market reference rows placed ABOVE Section 5 DCF summary
                ws_dcf["A2"] = "Current Enterprise Value"
                ws_dcf["B2"] = 2500.0
                ws_dcf["A3"] = "Market Enterprise Value"
                ws_dcf["B3"] = 2600.0
                ws_dcf["A5"] = "WACC"
                ws_dcf["B5"] = 0.09
                ws_dcf["A6"] = "Terminal Growth Rate"
                ws_dcf["B6"] = 0.025
                # Section 5 DCF valuation summary
                ws_dcf["A10"] = "PV of Terminal Value"
                ws_dcf["B10"] = 600.0
                ws_dcf["A11"] = ev_label
                ws_dcf["B11"] = 1000.0
                ws_dcf["A15"] = "Sensitivity Analysis (WACC vs Terminal Growth)"

                ws_wacc = wb.create_sheet("WACC")
                ws_wacc["A3"] = "WEIGHTED AVERAGE COST OF CAPITAL"
                ws_wacc["B3"] = 0.09
                wb.save(path)
                wb.close()

                results = validate_dcf_mod.validate_dcf_model(str(path))
                self.assertEqual(results["status"], "PASS", results)
                self.assertEqual(results["error_count"], 0, results["errors"])
                self.assertEqual(results["warning_count"], 0, results["warnings"])
                self.assertTrue(
                    any("Terminal value is 60.0% of EV" in i for i in results["info"]),
                    results["info"],
                )

    def test_standalone_validate_dcf_outside_repo_emits_graceful_json_error(self) -> None:
        with tempfile.TemporaryDirectory(prefix="standalone_dcf_") as tmp:
            isolated_dir = Path(tmp)
            isolated_script = isolated_dir / "validate_dcf.py"
            shutil.copyfile(VALIDATE_DCF_PATH, isolated_script)

            sample_xlsx = isolated_dir / "sample.xlsx"
            self._build_compliant_two_sheet_dcf(sample_xlsx, with_formula=False)

            clean_env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
            proc = subprocess.run(
                [sys.executable, str(isolated_script), str(sample_xlsx)],
                capture_output=True,
                text=True,
                check=False,
                env=clean_env,
                cwd=tmp,
            )
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertEqual(proc.stderr, "", f"Unexpected stderr output: {proc.stderr}")
            self.assertNotIn("Traceback", proc.stdout + proc.stderr)

            payload = json.loads(proc.stdout)
            self.assertEqual(payload["status"], "ERROR", payload)
            self.assertIn("validation_date", payload)
            self.assertIn("xlsx_scan", payload.get("error", ""))
            self.assertGreaterEqual(payload["error_count"], 1, payload)
            self.assertEqual(payload["warning_count"], 0, payload)
            self.assertEqual(payload["warnings"], [], payload)
            self.assertEqual(payload["info"], [], payload)
            self.assertTrue(
                any("xlsx_scan" in err for err in payload["errors"]),
                f"Expected helpful error mentioning xlsx_scan, got: {payload['errors']}",
            )

            # Programmatic guard when scan_workbook_formulas is unavailable
            with mock.patch.object(validate_dcf_mod, "scan_workbook_formulas", None):
                prog_res = validate_dcf_mod.validate_dcf_model(str(sample_xlsx))
            self.assertEqual(prog_res["status"], "ERROR", prog_res)
            self.assertIn("xlsx_scan", prog_res.get("error", ""))
            self.assertEqual(prog_res["error_count"], 1, prog_res)
            self.assertTrue(any("xlsx_scan" in err for err in prog_res["errors"]), prog_res["errors"])

            # Verify documented PYTHONPATH workaround succeeds from the isolated location
            env_with_scripts = dict(clean_env)
            env_with_scripts["PYTHONPATH"] = str(SCRIPTS_DIR)
            proc_ok = subprocess.run(
                [sys.executable, str(isolated_script), str(sample_xlsx)],
                capture_output=True,
                text=True,
                check=False,
                env=env_with_scripts,
                cwd=tmp,
            )
            self.assertEqual(proc_ok.returncode, 0, proc_ok.stdout + proc_ok.stderr)
            self.assertEqual(json.loads(proc_ok.stdout)["status"], "PASS")

    def test_negative_terminal_growth_rate_supported_and_validated(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "negative_tg_dcf.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=False)
            wb = openpyxl.load_workbook(path)
            wb["DCF"]["B4"] = -0.01  # -1.0% declining terminal growth rate
            wb.save(path)
            wb.close()

            results = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(results["status"], "PASS", results)
            self.assertEqual(results["error_count"], 0, results["errors"])
            self.assertEqual(results["warning_count"], 0, results["warnings"])
            self.assertTrue(
                any("-1.00%" in i and "9.00%" in i for i in results["info"]),
                f"Expected negative terminal growth validation info, got: {results['info']}",
            )

            # Negative WACC MUST still be rejected (WACC must remain strictly 0 < w < 1)
            neg_wacc_path = Path(tmp) / "negative_wacc_dcf.xlsx"
            self._build_compliant_two_sheet_dcf(neg_wacc_path, with_formula=False)
            wb2 = openpyxl.load_workbook(neg_wacc_path)
            wb2["DCF"]["B3"] = -0.05
            wb2["WACC"]["B3"] = -0.05
            wb2.save(neg_wacc_path)
            wb2.close()

            res_neg_wacc = validate_dcf_mod.validate_dcf_model(str(neg_wacc_path))
            self.assertEqual(res_neg_wacc["status"], "FAIL", res_neg_wacc)
            self.assertTrue(
                any("Could not locate WACC" in e for e in res_neg_wacc["errors"]),
                res_neg_wacc["errors"],
            )

    def test_broadened_tv_and_ev_labels_including_exit_multiple_pv(self) -> None:
        label_pairs = (
            ("PV of TV", "Enterprise Value"),
            ("Discounted Terminal Value", "Implied Enterprise Value"),
            ("Present Value of Terminal Value", "Total Enterprise Value"),
            ("PV Terminal Value", "Implied EV"),
            ("PV of Terminal Value (Exit Multiple)", "Implied Enterprise Value (Exit Multiple Method)"),
            ("PV of Terminal Value (Perpetuity)", "Enterprise Value"),
        )
        for tv_label, ev_label in label_pairs:
            with self.subTest(tv_label=tv_label, ev_label=ev_label), tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / "broad_labels.xlsx"
                wb = openpyxl.Workbook()
                ws_dcf = wb.active
                ws_dcf.title = "DCF"
                ws_dcf["A2"] = "WACC"
                ws_dcf["B2"] = 0.09
                ws_dcf["A3"] = "Terminal Growth Rate"
                ws_dcf["B3"] = 0.025
                # Undiscounted TV and ratio/multiple distractors placed ABOVE the real PV/EV rows
                ws_dcf["A4"] = "Undiscounted Terminal Value"
                ws_dcf["B4"] = 900.0
                ws_dcf["A5"] = "PV of Terminal Value (% of EV)"
                ws_dcf["B5"] = 0.60
                ws_dcf["A6"] = "Exit EV/EBITDA Multiple"
                ws_dcf["B6"] = 8.5
                ws_dcf["A7"] = tv_label
                ws_dcf["B7"] = 600.0
                ws_dcf["A8"] = ev_label
                ws_dcf["B8"] = 1000.0
                ws_dcf["A12"] = "Sensitivity Analysis"

                ws_wacc = wb.create_sheet("WACC")
                ws_wacc["A3"] = "WEIGHTED AVERAGE COST OF CAPITAL"
                ws_wacc["B3"] = 0.09
                wb.save(path)
                wb.close()

                results = validate_dcf_mod.validate_dcf_model(str(path))
                self.assertEqual(results["status"], "PASS", results)
                self.assertEqual(results["error_count"], 0, results["errors"])
                self.assertEqual(results["warning_count"], 0, results["warnings"])
                self.assertTrue(
                    any("Terminal value is 60.0% of EV" in i for i in results["info"]),
                    results["info"],
                )

    def test_named_ranges_resolve_even_with_non_standard_row_labels(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "named_ranges_custom_labels.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A1"] = "Row 1"
            ws_dcf["B1"] = 0.025
            ws_dcf["A2"] = "Row 2"
            ws_dcf["B2"] = 600.0
            ws_dcf["A3"] = "Row 3"
            ws_dcf["B3"] = 1000.0
            ws_dcf["A10"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A1"] = "Metric Alpha"
            ws_wacc["B1"] = 0.095

            wb.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="WACC!$B$1")
            wb.defined_names["TerminalGrowth"] = DefinedName("TerminalGrowth", attr_text="DCF!$B$1")
            wb.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="DCF!$B$2")
            wb.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$B$3")
            wb.save(path)
            wb.close()

            res_ok = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res_ok["status"], "PASS", res_ok)
            self.assertEqual(res_ok["error_count"], 0, res_ok["errors"])
            self.assertEqual(res_ok["warning_count"], 0, res_ok["warnings"])
            self.assertTrue(any("9.50%" in i for i in res_ok["info"]), res_ok["info"])
            self.assertTrue(any("Terminal value is 60.0% of EV" in i for i in res_ok["info"]), res_ok["info"])

            # Invalid TerminalGrowth >= DCF_WACC via named ranges must still FAIL
            wb_bad = openpyxl.load_workbook(path)
            wb_bad["DCF"]["B1"] = 0.10
            wb_bad.save(path)
            wb_bad.close()

            res_bad = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res_bad["status"], "FAIL", res_bad)
            self.assertTrue(
                any("Terminal growth (10.00% at DCF!B1) >= WACC (9.50% at WACC!B1)" in e for e in res_bad["errors"]),
                res_bad["errors"],
            )

    def test_scenario_scoped_pairing_never_crosses_bull_tg_with_bear_wacc(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # 1. Global WACC (9%) + Bear (WACC 11%, g 2%) + Bull (WACC 7%, g 7.5% INVALID)
            #    Old zip/nearest logic could cross Bull g with Bear WACC (11%) or global WACC (9%).
            path = Path(tmp) / "scenario_scoped.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A2"] = "Discount Rate (WACC)"
            ws_dcf["B2"] = 0.09

            ws_dcf["A5"] = "BEAR CASE ASSUMPTIONS"
            ws_dcf["A6"] = "WACC"
            ws_dcf["B6"] = 0.11
            ws_dcf["A7"] = "Terminal Growth Rate"
            ws_dcf["B7"] = 0.02

            ws_dcf["A12"] = "BULL CASE ASSUMPTIONS"
            ws_dcf["A13"] = "WACC"
            ws_dcf["B13"] = 0.07
            ws_dcf["A14"] = "Terminal Growth Rate"
            ws_dcf["B14"] = 0.075  # INVALID: 7.5% >= Bull WACC 7.0%

            ws_dcf["A20"] = "PV of Terminal Value"
            ws_dcf["B20"] = 600.0
            ws_dcf["A21"] = "Enterprise Value"
            ws_dcf["B21"] = 1000.0
            ws_dcf["A25"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A3"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B3"] = 0.09
            wb.save(path)
            wb.close()

            res = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res["status"], "FAIL", res)
            self.assertTrue(
                any("Terminal growth (7.50% at DCF!B14) >= WACC (7.00% at DCF!B13)" in e for e in res["errors"]),
                res["errors"],
            )

            # 2. Un-scoped Terminal Growth row when multiple scenario WACCs exist uses conservative min WACC (7.0%)
            path_unscoped = Path(tmp) / "scenario_unscoped_tg.xlsx"
            wb2 = openpyxl.load_workbook(path)
            wb2["DCF"]["B14"] = 0.03  # Fix Bull TG so Bull passes
            wb2["DCF"]["A17"] = "Active Summary"
            wb2["DCF"]["A18"] = "Terminal Growth Rate"
            wb2["DCF"]["B18"] = 0.072  # 7.2% >= min WACC (7.0%)
            wb2.save(path_unscoped)
            wb2.close()

            res_unscoped = validate_dcf_mod.validate_dcf_model(str(path_unscoped))
            self.assertEqual(res_unscoped["status"], "FAIL", res_unscoped)
            self.assertTrue(
                any("Terminal growth (7.20% at DCF!B18) >= WACC (7.00% at DCF!B13)" in e for e in res_unscoped["errors"]),
                res_unscoped["errors"],
            )

    def test_full_used_range_scans_beyond_row_260_and_column_aa(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "large_grid_dcf.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["AA265"] = "WACC"
            ws_dcf["AB265"] = 0.09
            ws_dcf["AA266"] = "Terminal Growth Rate"
            ws_dcf["AB266"] = 0.025
            ws_dcf["AA270"] = "PV of Terminal Value"
            ws_dcf["AB270"] = 600.0
            ws_dcf["AA271"] = "Enterprise Value"
            ws_dcf["AB271"] = 1000.0
            ws_dcf["AA280"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["AA265"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["AB265"] = 0.09
            wb.save(path)
            wb.close()

            res_ok = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res_ok["status"], "PASS", res_ok)
            self.assertEqual(res_ok["error_count"], 0, res_ok["errors"])
            self.assertEqual(res_ok["warning_count"], 0, res_ok["warnings"])
            self.assertTrue(any("Terminal value is 60.0% of EV" in i for i in res_ok["info"]), res_ok["info"])

            # Invalid g >= WACC far beyond row 250 and column Z MUST still be caught
            wb_bad = openpyxl.load_workbook(path)
            wb_bad["DCF"]["AB266"] = 0.11
            wb_bad.save(path)
            wb_bad.close()

            res_bad = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res_bad["status"], "FAIL", res_bad)
            self.assertTrue(
                any("Terminal growth (11.00% at DCF!AB266) >= WACC (9.00% at DCF!AB265)" in e for e in res_bad["errors"]),
                res_bad["errors"],
            )

    def test_detect_scenario_block_header_accepts_banners_and_rejects_ordinary_phrases(self) -> None:
        detect = validate_dcf_mod._detect_scenario_block_header

        positive_cases = [
            ("BEAR CASE", "bear"),
            ("BASE CASE", "base"),
            ("BULL CASE", "bull"),
            ("2A. BEAR CASE ASSUMPTIONS (Case = 1)", "bear"),
            ("2B. BASE CASE ASSUMPTIONS (Case = 2)", "base"),
            ("2C. BULL CASE ASSUMPTIONS (Case = 3)", "bull"),
            ("3. ACTIVE CASE ASSUMPTIONS (DRIVEN BY INDEX / CASE SELECTOR)", "active"),
            ("Active Case", "active"),
            ("Bear Scenario", "bear"),
            ("Upside Case", "upside"),
            ("Downside Case", "downside"),
            ("Worst Case", "worst"),
            ("Best Case", "best"),
            ("Scenario 1: Bear", "bear"),
            ("SECTION 2: BEAR CASE", "bear"),
            ("A) Bear Case", "bear"),
            ("B) Base Case", "base"),
            ("C) Bull Case", "bull"),
            ("A. Bear Case", "bear"),
            ("A: Bear Case", "bear"),
            ("Scenario 1", "scenario_1"),
            ("Scenario A", "scenario_a"),
            ("Case 2", "scenario_2"),
            ("Case B", "scenario_b"),
            ("Scenario II", "scenario_ii"),
            ("Bear", "bear"),
            ("Base", "base"),
            ("Bull", "bull"),
            ("Upside", "upside"),
            ("Downside", "downside"),
        ]
        for label, expected in positive_cases:
            with self.subTest(positive=label):
                self.assertEqual(detect(label, row_has_numeric=False), expected)

        negative_cases = [
            "LTM / Base Year Revenue (FY2025A)",
            "Revenue CAGR vs. base year",
            "Growth vs base",
            "Delta from base",
            "% of base",
            "Base Period Cash Flow",
            "Base Rate",
            "Base Revenue",
            "Base Margin",
            "Base Value",
            "Base Price",
            "Use Case",
            "Business Case",
            "Best Practice",
            "Case Selector",
            "Active Scenario Selector (1=Bear, 2=Base, 3=Bull)",
            "Case Notes",
            "Scenario Description",
            "Case Summary",
            "Case Study",
            "Scenario Inputs",
        ]
        for label in negative_cases:
            with self.subTest(negative=label):
                self.assertIsNone(detect(label, row_has_numeric=False))
                self.assertIsNone(detect(label, row_has_numeric=True))

        # Any row carrying unrelated numeric data MUST NOT act as a scenario block banner
        self.assertIsNone(detect("BASE CASE", row_has_numeric=True))
        self.assertIsNone(detect("Base", row_has_numeric=True))

    def test_base_year_label_does_not_split_bull_scenario_block(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # 1. Section-header per-scenario model: intervening "Revenue CAGR vs. base year"
            #    inside Bull Case MUST NOT start a 'base' block or hide Bull g (6.5%) >= WACC (6.0%).
            path = Path(tmp) / "base_year_intervening.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"

            ws_dcf["A2"] = "BASE CASE"
            ws_dcf["A3"] = "WACC"
            ws_dcf["B3"] = 0.09
            ws_dcf["A4"] = "Terminal Growth Rate"
            ws_dcf["B4"] = 0.025

            ws_dcf["A6"] = "BULL CASE"
            ws_dcf["A7"] = "WACC"
            ws_dcf["B7"] = 0.06
            ws_dcf["A8"] = "Revenue CAGR vs. base year"
            ws_dcf["B8"] = 0.12
            ws_dcf["A9"] = "Terminal Growth Rate"
            ws_dcf["B9"] = 0.065  # INVALID: 6.50% >= Bull WACC 6.00%

            ws_dcf["A12"] = "PV of Terminal Value"
            ws_dcf["B12"] = 600.0
            ws_dcf["A13"] = "Enterprise Value"
            ws_dcf["B13"] = 1000.0
            ws_dcf["A16"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A3"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B3"] = 0.09
            wb.save(path)
            wb.close()

            res = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res["status"], "FAIL", res)
            self.assertTrue(
                any(
                    "Terminal growth (6.50% at DCF!B9) >= WACC (6.00% at DCF!B7)" in e
                    for e in res["errors"]
                ),
                res["errors"],
            )

            # 2. Explicit per-row labels ("Bull WACC" -> "Revenue CAGR vs. base year" -> "Terminal Growth Rate")
            #    MUST also keep the Bull block intact and catch 6.50% >= 6.00%.
            path_inline = Path(tmp) / "base_year_inline_bull.xlsx"
            wb2 = openpyxl.Workbook()
            ws_dcf2 = wb2.active
            ws_dcf2.title = "DCF"
            ws_dcf2["A2"] = "Base WACC"
            ws_dcf2["B2"] = 0.09
            ws_dcf2["A3"] = "Base Terminal Growth"
            ws_dcf2["B3"] = 0.025
            ws_dcf2["A4"] = "Bull WACC"
            ws_dcf2["B4"] = 0.06
            ws_dcf2["A5"] = "Revenue CAGR vs. base year"
            ws_dcf2["B5"] = 0.12
            ws_dcf2["A6"] = "Terminal Growth Rate"
            ws_dcf2["B6"] = 0.065  # INVALID vs Bull WACC 6.0%
            ws_dcf2["A10"] = "PV of Terminal Value"
            ws_dcf2["B10"] = 600.0
            ws_dcf2["A11"] = "Enterprise Value"
            ws_dcf2["B11"] = 1000.0
            ws_dcf2["A14"] = "Sensitivity Analysis"

            ws_wacc2 = wb2.create_sheet("WACC")
            ws_wacc2["A3"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc2["B3"] = 0.09
            wb2.save(path_inline)
            wb2.close()

            res2 = validate_dcf_mod.validate_dcf_model(str(path_inline))
            self.assertEqual(res2["status"], "FAIL", res2)
            self.assertTrue(
                any(
                    "Terminal growth (6.50% at DCF!B6) >= WACC (6.00% at DCF!B4)" in e
                    for e in res2["errors"]
                ),
                res2["errors"],
            )

    @_skip_unless_required(
        recalc_xlsx_mod.find_soffice() is None,
        "LibreOffice (soffice) is required to evaluate build_dcf_skeleton.py",
    )
    def test_dcf_skeleton_base_year_revenue_does_not_tag_global_wacc_as_base(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "skeleton_tag_check.xlsx"
            build_proc = subprocess.run(
                [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(model_path)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

            recalc_proc = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30", "--in-place"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(recalc_proc.returncode, 0, recalc_proc.stdout + recalc_proc.stderr)

            validator = validate_dcf_mod.DCFModelValidator(str(model_path))
            assert validator._dcf_grid is not None
            tg_rows, wacc_rows = validator._extract_rate_rows("DCF", validator._dcf_grid)

            # Row 21 ("Discount Rate (WACC)") sits immediately after Row 20 ("LTM / Base Year Revenue (FY2025A)")
            # and MUST remain global (scenario_tag=None), NOT tagged as 'base'.
            row21_waccs = [entry for entry in wacc_rows if entry[0] == 21]
            self.assertEqual(len(row21_waccs), 1, f"Expected DCF!A21 WACC row, got: {wacc_rows}")
            self.assertIsNone(
                row21_waccs[0][2],
                f"Global WACC row DCF!A21 was wrongly tagged: {row21_waccs[0]}",
            )

            # Verify Bear/Base/Bull/Active scenario tags on Terminal Growth rows remain accurate
            tg_tags_by_row = {row_idx: tag for row_idx, _, tag, _ in tg_rows}
            self.assertEqual(tg_tags_by_row.get(33), "bear")
            self.assertEqual(tg_tags_by_row.get(45), "base")
            self.assertEqual(tg_tags_by_row.get(57), "bull")
            self.assertEqual(tg_tags_by_row.get(69), "active")

    def test_letter_prefixed_scenario_banners_isolate_bull_invalid_tg_from_bear_wacc(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # 1. Invalid B) Bull Case (g=7.5% >= WACC=7.0%) next to A) Bear Case (WACC=11.0%)
            #    MUST be caught rather than pairing Bull g against Bear WACC.
            path = Path(tmp) / "letter_prefixed_invalid_bull.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"

            ws_dcf["A2"] = "A) Bear Case"
            ws_dcf["A3"] = "WACC"
            ws_dcf["B3"] = 0.11
            ws_dcf["A4"] = "Terminal Growth Rate"
            ws_dcf["B4"] = 0.02

            ws_dcf["A6"] = "B) Bull Case"
            ws_dcf["A7"] = "WACC"
            ws_dcf["B7"] = 0.07
            ws_dcf["A8"] = "Terminal Growth Rate"
            ws_dcf["B8"] = 0.075  # INVALID: 7.50% >= Bull WACC 7.00%

            ws_dcf["A12"] = "PV of Terminal Value"
            ws_dcf["B12"] = 600.0
            ws_dcf["A13"] = "Enterprise Value"
            ws_dcf["B13"] = 1000.0
            ws_dcf["A16"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A3"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B3"] = 0.09
            wb.save(path)
            wb.close()

            res = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res["status"], "FAIL", res)
            self.assertEqual(res["error_count"], 1, res["errors"])
            self.assertTrue(
                any(
                    "Terminal growth (7.50% at DCF!B8) >= WACC (7.00% at DCF!B7)" in e
                    for e in res["errors"]
                ),
                res["errors"],
            )

            # 2. Valid A) Bear Case / B) Base Case / C) Bull Case workbook passes cleanly
            valid_path = Path(tmp) / "letter_prefixed_valid_three_cases.xlsx"
            wb2 = openpyxl.Workbook()
            ws2 = wb2.active
            ws2.title = "DCF"

            ws2["A2"] = "A) Bear Case"
            ws2["A3"] = "WACC"
            ws2["B3"] = 0.11
            ws2["A4"] = "Terminal Growth Rate"
            ws2["B4"] = 0.02

            ws2["A6"] = "B) Base Case"
            ws2["A7"] = "WACC"
            ws2["B7"] = 0.09
            ws2["A8"] = "Terminal Growth Rate"
            ws2["B8"] = 0.025

            ws2["A10"] = "C) Bull Case"
            ws2["A11"] = "WACC"
            ws2["B11"] = 0.07
            ws2["A12"] = "Terminal Growth Rate"
            ws2["B12"] = 0.03

            ws2["A15"] = "PV of Terminal Value"
            ws2["B15"] = 600.0
            ws2["A16"] = "Enterprise Value"
            ws2["B16"] = 1000.0
            ws2["A19"] = "Sensitivity Analysis"

            ws_wacc2 = wb2.create_sheet("WACC")
            ws_wacc2["A3"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc2["B3"] = 0.09
            wb2.save(valid_path)
            wb2.close()

            res_ok = validate_dcf_mod.validate_dcf_model(str(valid_path))
            self.assertEqual(res_ok["status"], "PASS", res_ok)
            self.assertEqual(res_ok["error_count"], 0, res_ok["errors"])
            self.assertEqual(res_ok["warning_count"], 0, res_ok["warnings"])

    def test_case_notes_and_scenario_prose_headers_do_not_split_bear_scenario_block(self) -> None:
        prose_headers = (
            "Case Notes",
            "Scenario Description",
            "Case Summary",
            "Case Study",
            "Scenario Inputs",
        )
        for prose_header in prose_headers:
            with self.subTest(prose_header=prose_header), tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / "prose_header_inside_bear.xlsx"
                wb = openpyxl.Workbook()
                ws_dcf = wb.active
                ws_dcf.title = "DCF"

                # BEAR CASE (valid: g=7.5% < Bear WACC=11.0%) with intervening prose subheader.
                # If the prose header wrongly starts a new generic scenario block, Bear g=7.5%
                # would fall back to Bull WACC=7.0% and falsely FAIL!
                ws_dcf["A2"] = "BEAR CASE"
                ws_dcf["A3"] = "WACC"
                ws_dcf["B3"] = 0.11
                ws_dcf["A4"] = prose_header
                ws_dcf["A5"] = "Terminal Growth Rate"
                ws_dcf["B5"] = 0.075

                # BULL CASE (valid: g=3.0% < Bull WACC=7.0%)
                ws_dcf["A7"] = "BULL CASE"
                ws_dcf["A8"] = "WACC"
                ws_dcf["B8"] = 0.07
                ws_dcf["A9"] = "Terminal Growth Rate"
                ws_dcf["B9"] = 0.03

                ws_dcf["A12"] = "PV of Terminal Value"
                ws_dcf["B12"] = 600.0
                ws_dcf["A13"] = "Enterprise Value"
                ws_dcf["B13"] = 1000.0
                ws_dcf["A16"] = "Sensitivity Analysis"

                ws_wacc = wb.create_sheet("WACC")
                ws_wacc["A3"] = "WEIGHTED AVERAGE COST OF CAPITAL"
                ws_wacc["B3"] = 0.09
                wb.save(path)
                wb.close()

                res = validate_dcf_mod.validate_dcf_model(str(path))
                self.assertEqual(res["status"], "PASS", res)
                self.assertEqual(res["error_count"], 0, res["errors"])
                self.assertEqual(res["warning_count"], 0, res["warnings"])
                self.assertTrue(
                    any(
                        "Terminal growth (7.50% at DCF!B5) < WACC (11.00% at DCF!B3)" in i
                        for i in res["info"]
                    ),
                    res["info"],
                )
                self.assertTrue(
                    any(
                        "Terminal growth (3.00% at DCF!B9) < WACC (7.00% at DCF!B8)" in i
                        for i in res["info"]
                    ),
                    res["info"],
                )

    def test_sensitivity_header_strict_and_note_column_does_not_skip_bull_tg(self) -> None:
        is_sens = validate_dcf_mod._is_sensitivity_header
        for positive in (
            "Sensitivity Analysis",
            "7. SENSITIVITY ANALYSIS: WACC vs TERMINAL GROWTH (IMPLIED SHARE PRICE $)",
            "WACC vs Terminal Growth",
            "WACC vs. Exit Multiple",
            "Discount Rate vs. Terminal Growth",
            "WACC \\ Terminal g",
            "WACC x Terminal Growth",
            "Exit EV/EBITDA Multiple vs. WACC",
        ):
            with self.subTest(positive=positive):
                self.assertTrue(is_sens(positive))

        for negative in (
            "Margin vs. peer median",
            "Revenue vs. consensus",
            "Growth vs base",
            "WACC vs. peer average",
            "Terminal growth vs. GDP",
            "Gross margin vs. historical",
        ):
            with self.subTest(negative=negative):
                self.assertFalse(is_sens(negative))

        if recalc_xlsx_mod.find_soffice() is None:
            return

        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "repro_margin_note_bull_tg.xlsx"
            build_proc = subprocess.run(
                [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(model_path)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

            # Exact repro: Bull TG = 12% (> WACC ~11.04%) AND "Margin vs. peer median" in DCF!H39
            wb = openpyxl.load_workbook(model_path)
            wb["DCF"]["C57"] = 0.12
            wb["DCF"]["H39"] = "Margin vs. peer median"
            wb.save(model_path)
            wb.close()

            recalc_proc = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30", "--in-place"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(recalc_proc.returncode, 0, recalc_proc.stdout + recalc_proc.stderr)

            # 1. With named ranges intact -> FAIL catching 12.00% >= WACC
            res_named = validate_dcf_mod.validate_dcf_model(str(model_path))
            self.assertEqual(res_named["status"], "FAIL", res_named)
            self.assertTrue(
                any("12.00% at DCF!C57" in e and ">= WACC" in e for e in res_named["errors"]),
                res_named["errors"],
            )

            # 2. Strip ALL defined names to exercise the label-scan fallback path directly -> STILL FAIL!
            no_names_path = Path(tmp) / "repro_margin_note_no_names.xlsx"
            wb_strip = openpyxl.load_workbook(model_path)
            for dn_key in list(wb_strip.defined_names.keys()):
                del wb_strip.defined_names[dn_key]
            # Replace INT(Case) with INT($B$4) so stripping Case doesn't cause #NAME?
            ws_s = wb_strip["DCF"]
            ws_s["C4"] = '=CHOOSE(INT($B$4),"Bear Case","Base Case","Bull Case")'
            for r in range(62, 71):
                for c in range(3, 8):
                    v = ws_s.cell(row=r, column=c).value
                    if isinstance(v, str) and "INT(Case)" in v:
                        ws_s.cell(row=r, column=c).value = v.replace("INT(Case)", "INT($B$4)")
            wb_strip.save(no_names_path)
            wb_strip.close()

            recalc_strip = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(no_names_path), "30", "--in-place"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(recalc_strip.returncode, 0, recalc_strip.stdout + recalc_strip.stderr)

            res_fallback = validate_dcf_mod.validate_dcf_model(str(no_names_path))
            self.assertEqual(res_fallback["status"], "FAIL", res_fallback)
            self.assertTrue(
                any("12.00% at DCF!C57" in e and ">= WACC" in e for e in res_fallback["errors"]),
                res_fallback["errors"],
            )

    def test_tv_proportion_pairs_ev_in_valuation_block_ignoring_top_market_ev(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            # 1. Unlabeled Market Data "Enterprise Value = 26000" at Row 8 above Valuation Summary at Rows 90-91
            path = Path(tmp) / "market_ev_above_summary.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A8"] = "Enterprise Value"
            ws_dcf["B8"] = 26000.0
            ws_dcf["A20"] = "WACC"
            ws_dcf["B20"] = 0.09
            ws_dcf["A21"] = "Terminal Growth Rate"
            ws_dcf["B21"] = 0.025
            ws_dcf["A90"] = "PV of Terminal Value"
            ws_dcf["B90"] = 600.0
            ws_dcf["A91"] = "Enterprise Value"
            ws_dcf["B91"] = 1000.0
            ws_dcf["A100"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A5"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B5"] = 0.09
            wb.save(path)
            wb.close()

            res = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res["status"], "PASS", res)
            self.assertEqual(res["error_count"], 0, res["errors"])
            self.assertEqual(res["warning_count"], 0, res["warnings"])
            self.assertTrue(
                any("Terminal value is 60.0% of EV" in i for i in res["info"]),
                res["info"],
            )

            # 2. EV placed immediately preceding PV of TV (Row 89 before Row 90, no EV after Row 90)
            path_prec = Path(tmp) / "ev_immediately_above_pv_tv.xlsx"
            wb2 = openpyxl.Workbook()
            ws2 = wb2.active
            ws2.title = "DCF"
            ws2["A8"] = "Enterprise Value"
            ws2["B8"] = 26000.0
            ws2["A20"] = "WACC"
            ws2["B20"] = 0.09
            ws2["A21"] = "Terminal Growth Rate"
            ws2["B21"] = 0.025
            ws2["A89"] = "Enterprise Value"
            ws2["B89"] = 1000.0
            ws2["A90"] = "PV of Terminal Value"
            ws2["B90"] = 600.0
            ws2["A100"] = "Sensitivity Analysis"

            ws_wacc2 = wb2.create_sheet("WACC")
            ws_wacc2["A5"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc2["B5"] = 0.09
            wb2.save(path_prec)
            wb2.close()

            res_prec = validate_dcf_mod.validate_dcf_model(str(path_prec))
            self.assertEqual(res_prec["status"], "PASS", res_prec)
            self.assertEqual(res_prec["warning_count"], 0, res_prec["warnings"])
            self.assertTrue(
                any("Terminal value is 60.0% of EV" in i for i in res_prec["info"]),
                res_prec["info"],
            )

    def test_per_scenario_named_ranges_validate_all_scenarios_without_regex_fallback(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "per_scenario_named.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            # Intentionally opaque labels so only named ranges can resolve these cells
            ws_dcf["A1"] = "Item 1"
            ws_dcf["B1"] = 0.02   # Bear TG
            ws_dcf["A2"] = "Item 2"
            ws_dcf["B2"] = 0.025  # Base TG
            ws_dcf["A3"] = "Item 3"
            ws_dcf["B3"] = 0.03   # Bull TG
            ws_dcf["A4"] = "Item 4"
            ws_dcf["B4"] = 600.0  # PV_TerminalValue
            ws_dcf["A5"] = "Item 5"
            ws_dcf["B5"] = 1000.0  # EnterpriseValue
            ws_dcf["A10"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A1"] = "Rate 1"
            ws_wacc["B1"] = 0.11  # Bear WACC
            ws_wacc["A2"] = "Rate 2"
            ws_wacc["B2"] = 0.09  # Base WACC
            ws_wacc["A3"] = "Rate 3"
            ws_wacc["B3"] = 0.07  # Bull WACC

            wb.defined_names["TerminalGrowth_Bear"] = DefinedName("TerminalGrowth_Bear", attr_text="DCF!$B$1")
            wb.defined_names["TerminalGrowth_Base"] = DefinedName("TerminalGrowth_Base", attr_text="DCF!$B$2")
            wb.defined_names["TerminalGrowth_Bull"] = DefinedName("TerminalGrowth_Bull", attr_text="DCF!$B$3")
            wb.defined_names["WACC_Bear"] = DefinedName("WACC_Bear", attr_text="WACC!$B$1")
            wb.defined_names["WACC_Base"] = DefinedName("WACC_Base", attr_text="WACC!$B$2")
            wb.defined_names["WACC_Bull"] = DefinedName("WACC_Bull", attr_text="WACC!$B$3")
            wb.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="DCF!$B$4")
            wb.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$B$5")
            wb.save(path)
            wb.close()

            real_extract = validate_dcf_mod.DCFModelValidator._extract_rate_rows
            with mock.patch.object(
                validate_dcf_mod.DCFModelValidator,
                "_extract_rate_rows",
                side_effect=lambda self_obj, title, grid: real_extract(self_obj, title, grid),
                autospec=True,
            ) as spy_extract:
                res_ok = validate_dcf_mod.validate_dcf_model(str(path))

            self.assertEqual(res_ok["status"], "PASS", res_ok)
            self.assertEqual(res_ok["error_count"], 0, res_ok["errors"])
            self.assertEqual(res_ok["warning_count"], 0, res_ok["warnings"])
            self.assertEqual(
                spy_extract.call_count,
                2,
                "Expected _extract_rate_rows once per sheet (DCF + WACC) so unnamed rate cells are always merged",
            )
            self.assertTrue(
                any("Validated scenario Terminal Growth and WACC via workbook named ranges" in i for i in res_ok["info"]),
                res_ok["info"],
            )
            self.assertFalse(
                any("Using label-scan fallback" in i for i in res_ok["info"]),
                res_ok["info"],
            )

            # Make Bull TG (7.5%) >= Bull WACC (7.0%) and Bear WACC out-of-range (25%)
            wb_bad = openpyxl.load_workbook(path)
            wb_bad["DCF"]["B3"] = 0.075
            wb_bad["WACC"]["B1"] = 0.25
            wb_bad.save(path)
            wb_bad.close()

            res_bad = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res_bad["status"], "FAIL", res_bad)
            self.assertTrue(
                any("Terminal growth (7.50% at DCF!B3) >= WACC (7.00% at WACC!B3)" in e for e in res_bad["errors"]),
                res_bad["errors"],
            )
            self.assertTrue(
                any("WACC (25.00% at WACC!B1) is outside typical range" in w for w in res_bad["warnings"]),
                res_bad["warnings"],
            )

    def test_validate_dcf_single_pass_load_and_single_rate_row_extraction(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "single_pass_heuristic.xlsx"
            self._build_compliant_two_sheet_dcf(path, with_formula=False)

            real_scan = validate_dcf_mod.scan_workbook_formulas
            real_load = openpyxl.load_workbook
            real_extract = validate_dcf_mod.DCFModelValidator._extract_rate_rows

            with (
                mock.patch.object(validate_dcf_mod, "scan_workbook_formulas", wraps=real_scan) as spy_scan,
                mock.patch.object(openpyxl, "load_workbook", wraps=real_load) as spy_load,
                mock.patch.object(
                    validate_dcf_mod.DCFModelValidator,
                    "_extract_rate_rows",
                    side_effect=lambda self_obj, title, grid: real_extract(self_obj, title, grid),
                    autospec=True,
                ) as spy_extract,
            ):
                res = validate_dcf_mod.validate_dcf_model(str(path))

            self.assertEqual(res["status"], "PASS", res)
            self.assertEqual(spy_scan.call_count, 1, "Expected exactly 1 scan_workbook_formulas call")
            # Single lockstep pass loads wb_formulas + wb_values once each (= 2 opens total, never 3)
            self.assertEqual(spy_load.call_count, 2, "Expected exactly 2 openpyxl.load_workbook calls (never 3)")
            # Label-scan fallback extracts DCF once and WACC once (= 2 total, never 4)
            self.assertEqual(spy_extract.call_count, 2, "Expected _extract_rate_rows once per sheet (DCF + WACC)")

    def test_named_out_of_bounds_tg_and_negative_pv_tv_caught_without_exit_multiple_mis_pairing(self) -> None:
        if recalc_xlsx_mod.find_soffice() is None:
            if REQUIRE_FULL_TEST_SUITE:
                self.fail("REQUIRE_FULL_TEST_SUITE=1 requires LibreOffice (soffice)")
            self.skipTest("LibreOffice (soffice) is required for end-to-end DCF skeleton named-range repro")

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Exact repro under Case = 2 (Base active) with TerminalGrowth_Bull = 2.5 (250%) and 1.0 (100%)
            for bad_tg, expected_pct in ((2.5, "250.00%"), (1.0, "100.00%")):
                with self.subTest(case=2, bad_tg=bad_tg):
                    model_path = Path(tmp) / f"case2_tg_{str(bad_tg).replace('.', '_')}.xlsx"
                    build_proc = subprocess.run(
                        [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(model_path), "--case", "2"],
                        capture_output=True,
                        text=True,
                        check=False,
                    )
                    self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

                    wb = openpyxl.load_workbook(model_path)
                    wb["DCF"]["C57"] = bad_tg
                    wb.save(model_path)
                    wb.close()

                    recalc_proc = subprocess.run(
                        [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30", "--in-place"],
                        capture_output=True,
                        text=True,
                        check=False,
                    )
                    self.assertEqual(recalc_proc.returncode, 0, recalc_proc.stdout + recalc_proc.stderr)

                    res = validate_dcf_mod.validate_dcf_model(str(model_path))
                    self.assertEqual(res["status"], "FAIL", res)
                    self.assertEqual(res["warning_count"], 0, res["warnings"])
                    self.assertTrue(
                        any(f"Terminal growth ({expected_pct} at DCF!C57) >= WACC" in e for e in res["errors"]),
                        res["errors"],
                    )

            # 2. Exact repro under Case = 3 (Bull active) with TerminalGrowth_Bull = 2.5:
            #    MUST FAIL with g >= WACC error AND MUST NOT emit bogus Exit Multiple TV % of EV warning
            model_c3 = Path(tmp) / "case3_tg_2_5.xlsx"
            build_c3 = subprocess.run(
                [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(model_c3), "--case", "3"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(build_c3.returncode, 0, build_c3.stdout + build_c3.stderr)

            wb3 = openpyxl.load_workbook(model_c3)
            wb3["DCF"]["C57"] = 2.5
            wb3.save(model_c3)
            wb3.close()

            recalc_c3 = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(model_c3), "30", "--in-place"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(recalc_c3.returncode, 0, recalc_c3.stdout + recalc_c3.stderr)

            res_c3 = validate_dcf_mod.validate_dcf_model(str(model_c3))
            self.assertEqual(res_c3["status"], "FAIL", res_c3)
            self.assertTrue(
                any("Terminal growth (250.00% at DCF!C57) >= WACC" in e for e in res_c3["errors"]),
                res_c3["errors"],
            )
            self.assertFalse(
                any("typically should be 50-70%" in w for w in res_c3["warnings"]),
                f"Bogus Exit Multiple TV % of EV warning emitted: {res_c3['warnings']}",
            )
            self.assertTrue(
                any("is non-positive" in w for w in res_c3["warnings"]),
                f"Expected explicit non-positive TV/EV warning, got: {res_c3['warnings']}",
            )

    def test_named_invalid_wacc_negative_tv_ev_and_non_numeric_named_cells(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        def _make_named_wb(path: Path, **overrides) -> None:
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A1"] = "Bear TG"
            ws_dcf["B1"] = overrides.get("tg_bear", 0.02)
            ws_dcf["A2"] = "Base TG"
            ws_dcf["B2"] = overrides.get("tg_base", 0.025)
            ws_dcf["A3"] = "Bull TG"
            ws_dcf["B3"] = overrides.get("tg_bull", 0.03)
            ws_dcf["A4"] = "PV of Terminal Value"
            ws_dcf["B4"] = overrides.get("pv_tv", 600.0)
            ws_dcf["A5"] = "Enterprise Value"
            ws_dcf["B5"] = overrides.get("ev", 1000.0)
            # Downstream Exit Multiple TV row that must NEVER be paired with named Perpetuity EV
            ws_dcf["A8"] = "PV of Exit Multiple TV"
            ws_dcf["B8"] = 2500.0
            ws_dcf["A12"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A1"] = "Bear WACC"
            ws_wacc["B1"] = overrides.get("wacc_bear", 0.11)
            ws_wacc["A2"] = "Base WACC"
            ws_wacc["B2"] = overrides.get("wacc_base", 0.09)
            ws_wacc["A3"] = "Bull WACC"
            ws_wacc["B3"] = overrides.get("wacc_bull", 0.08)

            wb.defined_names["TerminalGrowth_Bear"] = DefinedName("TerminalGrowth_Bear", attr_text="DCF!$B$1")
            wb.defined_names["TerminalGrowth_Base"] = DefinedName("TerminalGrowth_Base", attr_text="DCF!$B$2")
            wb.defined_names["TerminalGrowth_Bull"] = DefinedName("TerminalGrowth_Bull", attr_text="DCF!$B$3")
            wb.defined_names["WACC_Bear"] = DefinedName("WACC_Bear", attr_text="WACC!$B$1")
            wb.defined_names["WACC_Base"] = DefinedName("WACC_Base", attr_text="WACC!$B$2")
            wb.defined_names["WACC_Bull"] = DefinedName("WACC_Bull", attr_text="WACC!$B$3")
            if overrides.get("define_pv_tv", True):
                wb.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="DCF!$B$4")
            if overrides.get("define_ev", True):
                wb.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$B$5")
            wb.save(path)
            wb.close()

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Negative WACC_Bull (-2.00%) -> FAIL with non-positive WACC error
            p_neg_wacc = Path(tmp) / "neg_wacc.xlsx"
            _make_named_wb(p_neg_wacc, wacc_bull=-0.02, tg_bull=-0.05)
            res_neg = validate_dcf_mod.validate_dcf_model(str(p_neg_wacc))
            self.assertEqual(res_neg["status"], "FAIL", res_neg)
            self.assertTrue(
                any("Non-positive WACC (-2.00% at WACC!B3)" in e for e in res_neg["errors"]),
                res_neg["errors"],
            )
            self.assertFalse(
                any("-5.00%" in i and "-2.00%" in i for i in res_neg["info"]),
                f"Should not emit pass checkmark against negative WACC: {res_neg['info']}",
            )

            # 2. Over-100% WACC_Bull (1.05 = 105%) -> FAIL with >= 100% decimal-entry error
            p_high_wacc = Path(tmp) / "high_wacc.xlsx"
            _make_named_wb(p_high_wacc, wacc_bull=1.05)
            res_high = validate_dcf_mod.validate_dcf_model(str(p_high_wacc))
            self.assertEqual(res_high["status"], "FAIL", res_high)
            self.assertTrue(
                any("WACC (105.00% at WACC!B3) is >= 100%" in e for e in res_high["errors"]),
                res_high["errors"],
            )
            self.assertFalse(
                any("105.00%" in i for i in res_high["info"]),
                f"Should not emit pass checkmark against >= 100% WACC: {res_high['info']}",
            )

            # 3. Extreme negative TerminalGrowth_Bull (-2.5 = -250%) -> FAIL
            p_neg_tg = Path(tmp) / "neg_tg.xlsx"
            _make_named_wb(p_neg_tg, tg_bull=-2.5)
            res_neg_tg = validate_dcf_mod.validate_dcf_model(str(p_neg_tg))
            self.assertEqual(res_neg_tg["status"], "FAIL", res_neg_tg)
            self.assertTrue(
                any("Terminal growth (-250.00% at DCF!B3) is <= -100%" in e for e in res_neg_tg["errors"]),
                res_neg_tg["errors"],
            )

            # 4. Boolean or text string in named rate cells -> FAIL explicitly (never silently dropped or coerced)
            for bad_val in (True, "2.5%"):
                with self.subTest(bad_tg_val=bad_val):
                    p_bad_tg = Path(tmp) / f"bad_tg_{type(bad_val).__name__}.xlsx"
                    _make_named_wb(p_bad_tg, tg_bull=bad_val)
                    res_btg = validate_dcf_mod.validate_dcf_model(str(p_bad_tg))
                    self.assertEqual(res_btg["status"], "FAIL", res_btg)
                    self.assertTrue(
                        any("non-numeric value" in e and "DCF!B3" in e for e in res_btg["errors"]),
                        res_btg["errors"],
                    )

                with self.subTest(bad_wacc_val=bad_val):
                    p_bad_w = Path(tmp) / f"bad_wacc_{type(bad_val).__name__}.xlsx"
                    _make_named_wb(p_bad_w, wacc_bull=bad_val)
                    res_bw = validate_dcf_mod.validate_dcf_model(str(p_bad_w))
                    self.assertEqual(res_bw["status"], "FAIL", res_bw)
                    self.assertTrue(
                        any("non-numeric value" in e and "WACC!B3" in e for e in res_bw["errors"]),
                        res_bw["errors"],
                    )
                    self.assertFalse(
                        any("DCF!B3" in i for i in res_bw["info"]),
                        f"Should not cross-pair Bull TG against another scenario's WACC: {res_bw['info']}",
                    )

            # 5. Non-positive or non-numeric named PV_TerminalValue / EnterpriseValue (including partial naming)
            #    must warn cleanly and NEVER fall back to pair Exit Multiple TV (2500) with EV (1000) -> 250%
            for label, kwargs in (
                ("neg_pv_tv_only_named", {"pv_tv": -500.0, "define_ev": False}),
                ("neg_ev_both_named", {"pv_tv": -500.0, "ev": -200.0}),
                ("zero_ev_only_named", {"ev": 0.0, "define_pv_tv": False}),
                ("nm_pv_tv", {"pv_tv": "NM"}),
            ):
                with self.subTest(case=label):
                    p_tv = Path(tmp) / f"tv_{label}.xlsx"
                    _make_named_wb(p_tv, **kwargs)
                    res_tv = validate_dcf_mod.validate_dcf_model(str(p_tv))
                    self.assertFalse(
                        any("typically should be 50-70%" in w for w in res_tv["warnings"]),
                        f"Mis-paired Exit Multiple TV on {label}: {res_tv['warnings']}",
                    )
                    self.assertTrue(
                        any(("is non-positive" in w or "non-numeric value" in w) for w in res_tv["warnings"]),
                        res_tv["warnings"],
                    )

    def test_defined_name_case_insensitive_sheet_title_resolves_in_scan_and_validator(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "mixed_case_defined_names.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A1"] = "Row 1"
            ws_dcf["B2"] = 0.025
            ws_dcf["B3"] = 600.0
            ws_dcf["B4"] = 1000.0
            ws_dcf["A10"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["D10"] = 0.095

            ws_extra = wb.create_sheet("Assumptions")
            ws_extra["C5"] = 0.03

            # Deliberately use lowercase / mixed-case sheet titles in defined-name formulas
            wb.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="wacc!$D$10")
            wb.defined_names["TerminalGrowth_Bear"] = DefinedName("TerminalGrowth_Bear", attr_text="dcf!$B$2")
            wb.defined_names["TerminalGrowth_Base"] = DefinedName("TerminalGrowth_Base", attr_text="Dcf!$B$2")
            wb.defined_names["TerminalGrowth_Bull"] = DefinedName("TerminalGrowth_Bull", attr_text="assumptions!$C$5")
            wb.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="dcf!$B$3")
            wb.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$B$4")
            wb.save(path)
            wb.close()

            # 1. scan_workbook_formulas canonicalizes sheet titles in defined_names to worksheet casing
            scan = xlsx_scan_mod.scan_workbook_formulas(
                path, capture_sheet_grids=True, capture_defined_names=True
            )
            self.assertEqual(scan["defined_names"]["DCF_WACC"], ("WACC", 10, 4))
            self.assertEqual(scan["defined_names"]["TerminalGrowth_Bear"], ("DCF", 2, 2))
            self.assertEqual(scan["defined_names"]["TerminalGrowth_Bull"], ("Assumptions", 5, 3))
            self.assertIn("Assumptions", scan["sheet_grids"])

            # 2. DCFModelValidator resolves every mixed-case defined name and passes cleanly
            validator = validate_dcf_mod.DCFModelValidator(str(path))
            self.assertEqual(validator._named_metrics["wacc"][0], 0.095)
            self.assertEqual(validator._named_metrics["wacc"][1], "WACC!D10")
            self.assertEqual(validator._named_scenario_tg["bull"][0], 0.03)
            self.assertEqual(validator._named_scenario_tg["bull"][1], "Assumptions!C5")

            res = validator.validate_all()
            self.assertEqual(res["status"], "PASS", res)
            self.assertEqual(res["error_count"], 0, res["errors"])
            self.assertEqual(res["warning_count"], 0, res["warnings"])

            # 3. Direct defense-in-depth check on _resolve_named_metric with un-canonicalized sheet title
            direct = validate_dcf_mod._resolve_named_metric(
                {"dcf_wacc": ("wacc", 10, 4)},
                validate_dcf_mod.NAMED_WACC_ALIASES,
                {"WACC": {10: {4: (0.095, "D10")}}},
            )
            self.assertEqual(direct, (0.095, "WACC!D10", "WACC", 10, 4))

    def test_defined_name_escaped_apostrophe_sheet_title_resolves_and_catches_invalid_rate(self) -> None:
        from openpyxl.utils import quote_sheetname
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "apostrophe_sheet_defined_names.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A1"] = "Row 1"
            ws_dcf["B2"] = 0.02
            ws_dcf["B3"] = 0.025
            ws_dcf["B4"] = 600.0
            ws_dcf["B5"] = 1000.0
            ws_dcf["A10"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["D10"] = 0.095

            ws_apostrophe = wb.create_sheet("Company's Model")
            ws_apostrophe["C7"] = 0.03

            # Verify openpyxl's own escaping rule produces doubled apostrophes
            self.assertEqual(quote_sheetname("Company's Model"), "'Company''s Model'")

            # Reference "Company's Model" using OOXML doubled-apostrophe escaping AND lowercase
            wb.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="WACC!$D$10")
            wb.defined_names["TerminalGrowth_Bear"] = DefinedName("TerminalGrowth_Bear", attr_text="DCF!$B$2")
            wb.defined_names["TerminalGrowth_Base"] = DefinedName("TerminalGrowth_Base", attr_text="DCF!$B$3")
            wb.defined_names["TerminalGrowth_Bull"] = DefinedName(
                "TerminalGrowth_Bull", attr_text="'company''s model'!$C$7"
            )
            wb.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="DCF!$B$4")
            wb.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$B$5")
            wb.save(path)
            wb.close()

            # 1. scan_workbook_formulas unescapes '' -> ' and captures "Company's Model" grid
            scan = xlsx_scan_mod.scan_workbook_formulas(
                path, capture_sheet_grids=True, capture_defined_names=True
            )
            self.assertEqual(scan["defined_names"]["TerminalGrowth_Bull"], ("Company's Model", 7, 3))
            self.assertIn("Company's Model", scan["sheet_grids"])
            self.assertEqual(scan["sheet_grids"]["Company's Model"][7][3], (0.03, "C7"))

            # 2. Valid rate on "Company's Model" resolves and passes cleanly
            res_ok = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res_ok["status"], "PASS", res_ok)
            self.assertEqual(res_ok["error_count"], 0, res_ok["errors"])
            self.assertEqual(res_ok["warning_count"], 0, res_ok["warnings"])

            # 3. Invalid TerminalGrowth_Bull = 0.99 (99% >= WACC 9.5%) on "Company's Model" MUST FAIL
            wb_bad = openpyxl.load_workbook(path)
            wb_bad["Company's Model"]["C7"] = 0.99
            wb_bad.save(path)
            wb_bad.close()

            res_bad = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res_bad["status"], "FAIL", res_bad)
            self.assertTrue(
                any(
                    "Terminal growth (99.00% at Company's Model!C7) >= WACC (9.50% at WACC!D10)" in e
                    for e in res_bad["errors"]
                ),
                res_bad["errors"],
            )

            # 4. Defense-in-depth check on _lookup_sheet_grid / _resolve_named_metric / _find_invalid_named_metric
            direct_val = validate_dcf_mod._resolve_named_metric(
                {"terminalgrowth_bull": ("company''s model", 7, 3)},
                validate_dcf_mod.NAMED_SCENARIO_TG_ALIASES["bull"],
                {"Company's Model": {7: {3: (0.03, "C7")}}},
            )
            self.assertEqual(direct_val, (0.03, "Company's Model!C7", "Company's Model", 7, 3))

            direct_empty = validate_dcf_mod._find_invalid_named_metric(
                {"terminalgrowth_bull": ("company''s model", 7, 3)},
                validate_dcf_mod.NAMED_SCENARIO_TG_ALIASES["bull"],
                {"Company's Model": {}},
            )
            self.assertEqual(direct_empty, (None, "Company's Model!C7"))

    def test_cleared_empty_named_cells_caught_without_mis_pairing_or_uncached_noise(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Cleared TerminalGrowth_Bull (None) on recalculated skeleton (Case = 2) -> FAIL with explicit empty error
            if recalc_xlsx_mod.find_soffice() is not None:
                skel_path = Path(tmp) / "skeleton_cleared_bull_tg.xlsx"
                build_proc = subprocess.run(
                    [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(skel_path), "--case", "2"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

                # First check unrecalculated skeleton (uncached_formulas > 0): must FAIL on uncached formulas
                # WITHOUT emitting spurious "is empty" errors for formula-driven named cells
                res_uncached = validate_dcf_mod.validate_dcf_model(str(skel_path))
                self.assertEqual(res_uncached["status"], "FAIL", res_uncached)
                self.assertTrue(
                    any("Uncached formula values detected" in e for e in res_uncached["errors"]),
                    res_uncached["errors"],
                )
                self.assertFalse(
                    any("is empty" in e for e in res_uncached["errors"]),
                    f"Spurious empty-named-cell error on uncached workbook: {res_uncached['errors']}",
                )

                # Clear DCF!C57 (TerminalGrowth_Bull) and recalculate with LibreOffice
                wb_clr = openpyxl.load_workbook(skel_path)
                wb_clr["DCF"]["C57"] = None
                wb_clr.save(skel_path)
                wb_clr.close()

                recalc_proc = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(skel_path), "30", "--in-place"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_proc.returncode, 0, recalc_proc.stdout + recalc_proc.stderr)

                res_clr = validate_dcf_mod.validate_dcf_model(str(skel_path))
                self.assertEqual(res_clr["status"], "FAIL", res_clr)
                self.assertTrue(
                    any(
                        "Named terminal growth (Bull at DCF!C57) is empty" in e
                        for e in res_clr["errors"]
                    ),
                    res_clr["errors"],
                )

            # 2. Cleared PV_TerminalValue (None) alongside valid EnterpriseValue (1000) and Exit Multiple TV (2500)
            #    must warn cleanly and NEVER mis-pair Exit Multiple TV (250.0% of EV)
            p_empty_tv = Path(tmp) / "empty_pv_tv.xlsx"
            wb_tv = openpyxl.Workbook()
            ws_dcf = wb_tv.active
            ws_dcf.title = "DCF"
            ws_dcf["A1"] = "Bear TG"
            ws_dcf["B1"] = 0.02
            ws_dcf["A2"] = "Base TG"
            ws_dcf["B2"] = 0.025
            ws_dcf["A3"] = "Bull TG"
            ws_dcf["B3"] = 0.03
            ws_dcf["A4"] = "PV of Terminal Value"
            ws_dcf["B4"] = None
            ws_dcf["A5"] = "Enterprise Value"
            ws_dcf["B5"] = 1000.0
            ws_dcf["A8"] = "PV of Exit Multiple TV"
            ws_dcf["B8"] = 2500.0
            ws_dcf["A12"] = "Sensitivity Analysis"

            ws_wacc = wb_tv.create_sheet("WACC")
            ws_wacc["A1"] = "Bear WACC"
            ws_wacc["B1"] = 0.11
            ws_wacc["A2"] = "Base WACC"
            ws_wacc["B2"] = 0.09
            ws_wacc["A3"] = "Bull WACC"
            ws_wacc["B3"] = 0.08

            wb_tv.defined_names["TerminalGrowth_Bear"] = DefinedName("TerminalGrowth_Bear", attr_text="DCF!$B$1")
            wb_tv.defined_names["TerminalGrowth_Base"] = DefinedName("TerminalGrowth_Base", attr_text="DCF!$B$2")
            wb_tv.defined_names["TerminalGrowth_Bull"] = DefinedName("TerminalGrowth_Bull", attr_text="DCF!$B$3")
            wb_tv.defined_names["WACC_Bear"] = DefinedName("WACC_Bear", attr_text="WACC!$B$1")
            wb_tv.defined_names["WACC_Base"] = DefinedName("WACC_Base", attr_text="WACC!$B$2")
            wb_tv.defined_names["WACC_Bull"] = DefinedName("WACC_Bull", attr_text="WACC!$B$3")
            wb_tv.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="DCF!$B$4")
            wb_tv.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$B$5")
            wb_tv.save(p_empty_tv)
            wb_tv.close()

            res_tv = validate_dcf_mod.validate_dcf_model(str(p_empty_tv))
            self.assertEqual(res_tv["status"], "PASS", res_tv)
            self.assertTrue(
                any(
                    "Named PV_TerminalValue at DCF!B4 is empty or uncached" in w
                    for w in res_tv["warnings"]
                ),
                res_tv["warnings"],
            )
            self.assertFalse(
                any("typically should be 50-70%" in w or "250.0%" in w for w in res_tv["warnings"]),
                f"Bogus Exit Multiple TV mis-pairing warning emitted: {res_tv['warnings']}",
            )

            # 3. Cleared WACC_Bull (None) -> FAIL with explicit empty WACC error and no cross-scenario pairing
            p_empty_wacc = Path(tmp) / "empty_wacc_bull.xlsx"
            shutil.copyfile(p_empty_tv, p_empty_wacc)
            wb_ew = openpyxl.load_workbook(p_empty_wacc)
            wb_ew["DCF"]["B4"] = 600.0
            wb_ew["WACC"]["B3"] = None
            wb_ew.save(p_empty_wacc)
            wb_ew.close()

            res_ew = validate_dcf_mod.validate_dcf_model(str(p_empty_wacc))
            self.assertEqual(res_ew["status"], "FAIL", res_ew)
            self.assertTrue(
                any("Named WACC (Bull at WACC!B3) is empty" in e for e in res_ew["errors"]),
                res_ew["errors"],
            )
            self.assertFalse(
                any("DCF!B3" in i for i in res_ew["info"]),
                f"Should not cross-pair Bull TG against another scenario's WACC: {res_ew['info']}",
            )

            # 4. Named cell targeting nonexistent sheet (grid is None) -> no crash; formula error -> no duplicate
            p_missing_and_err = Path(tmp) / "missing_sheet_and_err.xlsx"
            shutil.copyfile(p_empty_tv, p_missing_and_err)
            wb_me = openpyxl.load_workbook(p_missing_and_err)
            wb_me["DCF"]["B4"] = 600.0
            wb_me["DCF"]["B3"] = "#DIV/0!"
            wb_me.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="'Deleted Sheet'!$A$1")
            wb_me.save(p_missing_and_err)
            wb_me.close()

            res_me = validate_dcf_mod.validate_dcf_model(str(p_missing_and_err))
            self.assertEqual(res_me["status"], "FAIL", res_me)
            self.assertTrue(any("#DIV/0! at DCF!B3" in e for e in res_me["errors"]), res_me["errors"])
            self.assertFalse(
                any("DCF!B3" in e and ("non-numeric" in e or "is empty" in e) for e in res_me["errors"]),
                f"Duplicate non-numeric/empty error emitted for formula error cell: {res_me['errors']}",
            )

    def test_plain_text_error_token_in_named_cells_caught_while_real_formula_errors_deduplicated(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            has_soffice = recalc_xlsx_mod.find_soffice() is not None

            # 1. Plain-text "'#N/A" (data_type == "s") in TerminalGrowth_Bull (DCF!C57) on a
            #    build_dcf_skeleton.py --case 2 workbook recalculated via real LibreOffice must FAIL
            #    with a non-numeric TerminalGrowth error (not silently PASS).
            if has_soffice:
                skel_path = Path(tmp) / "skeleton_plain_text_na_bull_tg.xlsx"
                build_proc = subprocess.run(
                    [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(skel_path), "--case", "2"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

                wb_na = openpyxl.load_workbook(skel_path)
                cell_tg_bull = wb_na["DCF"]["C57"]
                cell_tg_bull.value = "#N/A"
                cell_tg_bull.data_type = "s"
                wb_na.save(skel_path)
                wb_na.close()

                recalc_na = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(skel_path), "30", "--in-place"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_na.returncode, 0, recalc_na.stdout + recalc_na.stderr)
                recalc_na_payload = json.loads(recalc_na.stdout)
                self.assertEqual(recalc_na_payload["total_errors"], 0, recalc_na_payload)

                res_tg_na = validate_dcf_mod.validate_dcf_model(str(skel_path))
                self.assertEqual(res_tg_na["status"], "FAIL", res_tg_na)
                self.assertTrue(
                    any(
                        "Named terminal growth (Bull at DCF!C57) has non-numeric value ('#N/A')" in e
                        for e in res_tg_na["errors"]
                    ),
                    res_tg_na["errors"],
                )

                # 2. Real Excel formula error (=1/0 -> #DIV/0!) in TerminalGrowth_Bull (DCF!C57)
                #    must still emit only the formula error (deduplicated via _formula_error_coords).
                wb_div0 = openpyxl.load_workbook(skel_path)
                wb_div0["DCF"]["C57"] = "=1/0"
                wb_div0.save(skel_path)
                wb_div0.close()

                recalc_div0 = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(skel_path), "30", "--in-place"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_div0.returncode, 1, recalc_div0.stdout + recalc_div0.stderr)

                res_tg_div0 = validate_dcf_mod.validate_dcf_model(str(skel_path))
                self.assertEqual(res_tg_div0["status"], "FAIL", res_tg_div0)
                self.assertTrue(
                    any("#DIV/0! at DCF!C57" in e for e in res_tg_div0["errors"]),
                    res_tg_div0["errors"],
                )
                self.assertFalse(
                    any("DCF!C57" in e and ("non-numeric" in e or "is empty" in e) for e in res_tg_div0["errors"]),
                    f"Duplicate non-numeric/empty error emitted for formula error cell: {res_tg_div0['errors']}",
                )

            # Helper for named-range WACC / PV_TerminalValue tests
            def _make_named_rates_wb(dest: Path) -> None:
                wb = openpyxl.Workbook()
                ws_dcf = wb.active
                ws_dcf.title = "DCF"
                ws_dcf["A1"] = "Bear TG"
                ws_dcf["B1"] = 0.02
                ws_dcf["A2"] = "Base TG"
                ws_dcf["B2"] = 0.025
                ws_dcf["A3"] = "Bull TG"
                ws_dcf["B3"] = 0.03
                ws_dcf["A4"] = "PV of Terminal Value"
                ws_dcf["B4"] = 600.0
                ws_dcf["A5"] = "Enterprise Value"
                ws_dcf["B5"] = 1000.0
                ws_dcf["A8"] = "PV of Exit Multiple TV"
                ws_dcf["B8"] = 2500.0
                ws_dcf["A12"] = "Sensitivity Analysis"

                ws_wacc = wb.create_sheet("WACC")
                ws_wacc["A1"] = "Bear WACC"
                ws_wacc["B1"] = 0.11
                ws_wacc["A2"] = "Base WACC"
                ws_wacc["B2"] = 0.09
                ws_wacc["A3"] = "Bull WACC"
                ws_wacc["B3"] = 0.08

                wb.defined_names["TerminalGrowth_Bear"] = DefinedName("TerminalGrowth_Bear", attr_text="DCF!$B$1")
                wb.defined_names["TerminalGrowth_Base"] = DefinedName("TerminalGrowth_Base", attr_text="DCF!$B$2")
                wb.defined_names["TerminalGrowth_Bull"] = DefinedName("TerminalGrowth_Bull", attr_text="DCF!$B$3")
                wb.defined_names["WACC_Bear"] = DefinedName("WACC_Bear", attr_text="WACC!$B$1")
                wb.defined_names["WACC_Base"] = DefinedName("WACC_Base", attr_text="WACC!$B$2")
                wb.defined_names["WACC_Bull"] = DefinedName("WACC_Bull", attr_text="WACC!$B$3")
                wb.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="DCF!$B$4")
                wb.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$B$5")
                wb.save(dest)
                wb.close()

            # 3. Plain-text "#N/A" (data_type == "s") in WACC_Bull -> FAIL with non-numeric WACC error
            p_wacc_na = Path(tmp) / "plain_text_na_wacc_bull.xlsx"
            _make_named_rates_wb(p_wacc_na)
            wb_wna = openpyxl.load_workbook(p_wacc_na)
            c_wacc_bull = wb_wna["WACC"]["B3"]
            c_wacc_bull.value = "#N/A"
            c_wacc_bull.data_type = "s"
            wb_wna.save(p_wacc_na)
            wb_wna.close()

            res_wna = validate_dcf_mod.validate_dcf_model(str(p_wacc_na))
            self.assertEqual(res_wna["status"], "FAIL", res_wna)
            self.assertTrue(
                any(
                    "Named WACC (Bull at WACC!B3) has non-numeric value ('#N/A')" in e
                    for e in res_wna["errors"]
                ),
                res_wna["errors"],
            )
            self.assertFalse(
                any("DCF!B3" in i for i in res_wna["info"]),
                f"Should not cross-pair Bull TG against another scenario's WACC: {res_wna['info']}",
            )

            # 4. Plain-text "#N/A" (data_type == "s") in PV_TerminalValue -> clean warning, no Exit Multiple TV mis-pair
            p_tv_na = Path(tmp) / "plain_text_na_pv_tv.xlsx"
            _make_named_rates_wb(p_tv_na)
            wb_tvna = openpyxl.load_workbook(p_tv_na)
            c_pv_tv = wb_tvna["DCF"]["B4"]
            c_pv_tv.value = "#N/A"
            c_pv_tv.data_type = "s"
            wb_tvna.save(p_tv_na)
            wb_tvna.close()

            res_tvna = validate_dcf_mod.validate_dcf_model(str(p_tv_na))
            self.assertEqual(res_tvna["status"], "PASS", res_tvna)
            self.assertTrue(
                any(
                    "Named PV_TerminalValue at DCF!B4 has non-numeric value ('#N/A')" in w
                    for w in res_tvna["warnings"]
                ),
                res_tvna["warnings"],
            )
            self.assertFalse(
                any("typically should be 50-70%" in w or "250.0%" in w for w in res_tvna["warnings"]),
                f"Bogus Exit Multiple TV mis-pairing warning emitted: {res_tvna['warnings']}",
            )

            # 5. Real Excel formula error (=1/0 -> #DIV/0!) in WACC_Bull and PV_TerminalValue ->
            #    deduplicated cleanly via _formula_error_coords without duplicate error/warning spam.
            p_formula_err = Path(tmp) / "formula_err_wacc_and_pv_tv.xlsx"
            _make_named_rates_wb(p_formula_err)
            wb_fe = openpyxl.load_workbook(p_formula_err)
            if has_soffice:
                wb_fe["WACC"]["B3"] = "=1/0"
                wb_fe["DCF"]["B4"] = "=1/0"
            else:
                wb_fe["WACC"]["B3"] = "#DIV/0!"
                wb_fe["DCF"]["B4"] = "#DIV/0!"
            wb_fe.save(p_formula_err)
            wb_fe.close()

            if has_soffice:
                recalc_fe = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(p_formula_err), "30", "--in-place"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_fe.returncode, 1, recalc_fe.stdout + recalc_fe.stderr)

            res_fe = validate_dcf_mod.validate_dcf_model(str(p_formula_err))
            self.assertEqual(res_fe["status"], "FAIL", res_fe)
            self.assertTrue(any("#DIV/0! at WACC!B3" in e for e in res_fe["errors"]), res_fe["errors"])
            self.assertTrue(any("#DIV/0! at DCF!B4" in e for e in res_fe["errors"]), res_fe["errors"])
            self.assertFalse(
                any("WACC!B3" in e and ("non-numeric" in e or "is empty" in e) for e in res_fe["errors"]),
                f"Duplicate non-numeric/empty WACC error emitted for formula error cell: {res_fe['errors']}",
            )
            self.assertFalse(
                any("Named PV_TerminalValue at DCF!B4" in w for w in res_fe["warnings"]),
                f"Redundant PV_TerminalValue warning emitted for formula error cell: {res_fe['warnings']}",
            )
            self.assertFalse(
                any("typically should be 50-70%" in w or "250.0%" in w for w in res_fe["warnings"]),
                f"Bogus Exit Multiple TV mis-pairing warning emitted: {res_fe['warnings']}",
            )

    def test_unnamed_scenario_wacc_and_wacc_sheet_rows_checked_alongside_standard_named_ranges(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Exact repro: build_dcf_skeleton.py --case 2 with Bull TG (DCF!C57) = 10% and an unnamed
            #    "Bull Case WACC" = 9% at DCF!C59 while keeping all standard named ranges intact.
            #    Recalculate with real LibreOffice -> validate_dcf.py MUST return FAIL (g >= WACC).
            if recalc_xlsx_mod.find_soffice() is not None:
                skel_path = Path(tmp) / "repro_unnamed_bull_wacc.xlsx"
                build_proc = subprocess.run(
                    [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(skel_path), "--case", "2"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

                wb = openpyxl.load_workbook(skel_path)
                wb["DCF"]["C57"] = 0.10
                wb["DCF"]["A59"] = "Bull Case WACC"
                wb["DCF"]["C59"] = 0.09
                wb.save(skel_path)
                wb.close()

                recalc_proc = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(skel_path), "30", "--in-place"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_proc.returncode, 0, recalc_proc.stdout + recalc_proc.stderr)

                res_repro = validate_dcf_mod.validate_dcf_model(str(skel_path))
                self.assertEqual(res_repro["status"], "FAIL", res_repro)
                self.assertTrue(
                    any(
                        "Terminal growth (10.00% at DCF!C57) >= WACC (9.00% at DCF!C59)" in e
                        for e in res_repro["errors"]
                    ),
                    res_repro["errors"],
                )

                # 2. Unnamed out-of-range WACC row (25%) on the WACC sheet alongside DCF_WACC MUST warn
                wacc_row_path = Path(tmp) / "repro_unnamed_wacc_sheet_out_of_range.xlsx"
                build_proc2 = subprocess.run(
                    [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(wacc_row_path), "--case", "2"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(build_proc2.returncode, 0, build_proc2.stdout + build_proc2.stderr)

                wb2 = openpyxl.load_workbook(wacc_row_path)
                wb2["WACC"]["A35"] = "WACC"
                wb2["WACC"]["C35"] = 0.25
                wb2.save(wacc_row_path)
                wb2.close()

                recalc_proc2 = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(wacc_row_path), "30", "--in-place"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_proc2.returncode, 0, recalc_proc2.stdout + recalc_proc2.stderr)

                res_wacc_warn = validate_dcf_mod.validate_dcf_model(str(wacc_row_path))
                self.assertEqual(res_wacc_warn["status"], "PASS", res_wacc_warn)
                self.assertTrue(
                    any(
                        "WACC (25.00% at WACC!C35) is outside typical range (5%-20%)" in w
                        for w in res_wacc_warn["warnings"]
                    ),
                    res_wacc_warn["warnings"],
                )

            # 3. Standalone (no-LibreOffice) check: unnamed Bull Case WACC alongside standard named ranges
            #    AND unnamed out-of-range WACC row on WACC sheet both caught deterministically
            static_path = Path(tmp) / "static_named_plus_unnamed_wacc.xlsx"
            wb_s = openpyxl.Workbook()
            ws_dcf = wb_s.active
            ws_dcf.title = "DCF"
            ws_dcf["A2"] = "BEAR CASE"
            ws_dcf["A3"] = "Terminal Growth Rate"
            ws_dcf["C3"] = 0.02
            ws_dcf["A5"] = "BASE CASE"
            ws_dcf["A6"] = "Terminal Growth Rate"
            ws_dcf["C6"] = 0.025
            ws_dcf["A8"] = "BULL CASE"
            ws_dcf["A9"] = "Terminal Growth Rate"
            ws_dcf["C9"] = 0.10
            ws_dcf["A10"] = "Bull Case WACC"
            ws_dcf["C10"] = 0.09
            ws_dcf["A14"] = "PV of Terminal Value"
            ws_dcf["C14"] = 600.0
            ws_dcf["A15"] = "Enterprise Value"
            ws_dcf["C15"] = 1000.0
            ws_dcf["A20"] = "Sensitivity Analysis"

            ws_wacc = wb_s.create_sheet("WACC")
            ws_wacc["A5"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["C5"] = 0.11
            ws_wacc["A8"] = "WACC"
            ws_wacc["C8"] = 0.25

            wb_s.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="WACC!$C$5")
            wb_s.defined_names["TerminalGrowth_Bear"] = DefinedName("TerminalGrowth_Bear", attr_text="DCF!$C$3")
            wb_s.defined_names["TerminalGrowth_Base"] = DefinedName("TerminalGrowth_Base", attr_text="DCF!$C$6")
            wb_s.defined_names["TerminalGrowth_Bull"] = DefinedName("TerminalGrowth_Bull", attr_text="DCF!$C$9")
            wb_s.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="DCF!$C$14")
            wb_s.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$C$15")
            wb_s.save(static_path)
            wb_s.close()

            res_static = validate_dcf_mod.validate_dcf_model(str(static_path))
            self.assertEqual(res_static["status"], "FAIL", res_static)
            self.assertTrue(
                any(
                    "Terminal growth (10.00% at DCF!C9) >= WACC (9.00% at DCF!C10)" in e
                    for e in res_static["errors"]
                ),
                res_static["errors"],
            )
            self.assertTrue(
                any(
                    "WACC (25.00% at WACC!C8) is outside typical range (5%-20%)" in w
                    for w in res_static["warnings"]
                ),
                res_static["warnings"],
            )

    def test_multi_column_unnamed_same_tag_rows_align_by_column_without_cartesian_cross_pairing(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        def _build_multi_col_bull(path: Path) -> openpyxl.Workbook:
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A2"] = "Bull Case Terminal Growth"
            ws_dcf["C2"] = 0.07
            ws_dcf["D2"] = 0.09
            # Row 3 intentionally blank -> separates row 2 and row 4 into distinct blocks
            ws_dcf["A4"] = "Bull Case WACC"
            ws_dcf["C4"] = 0.08
            ws_dcf["D4"] = 0.10
            ws_dcf["A10"] = "PV of Terminal Value"
            ws_dcf["C10"] = 600.0
            ws_dcf["A11"] = "Enterprise Value"
            ws_dcf["C11"] = 1000.0

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A1"] = "Notes"
            ws_wacc["B1"] = "WACC schedule"

            ws_sens = wb.create_sheet("Sensitivity")
            ws_sens["A1"] = "WACC vs Terminal Growth"
            ws_sens["B2"] = 100.0
            return wb

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Valid multi-column same-tag rows separated by blank row MUST PASS with 1-to-1 column
            #    alignment (7% < 8%, 9% < 10%), NO cross-column 9% >= 8% error, NO 7% < 10% info line,
            #    and DCFModelValidator skips buffering the full Sensitivity sheet grid.
            valid_path = Path(tmp) / "multi_col_bull_valid.xlsx"
            wb1 = _build_multi_col_bull(valid_path)
            wb1.save(valid_path)
            wb1.close()

            validator1 = validate_dcf_mod.DCFModelValidator(str(valid_path))
            self.assertIn("Sensitivity", validator1.sheet_names)
            self.assertNotIn("Sensitivity", validator1._scan["sheet_grids"])
            self.assertIn("DCF", validator1._scan["sheet_grids"])
            self.assertIn("WACC", validator1._scan["sheet_grids"])

            res1 = validator1.validate_all()
            self.assertEqual(res1["status"], "PASS", res1)
            self.assertEqual(res1["error_count"], 0, res1["errors"])
            self.assertIn("✓ Terminal growth (7.00% at DCF!C2) < WACC (8.00% at DCF!C4)", res1["info"])
            self.assertIn("✓ Terminal growth (9.00% at DCF!D2) < WACC (10.00% at DCF!D4)", res1["info"])
            self.assertFalse(
                any("DCF!C2" in line and "DCF!D4" in line for line in res1["info"] + res1["errors"]),
                res1["info"],
            )
            self.assertFalse(
                any("DCF!D2" in line and "DCF!C4" in line for line in res1["info"] + res1["errors"]),
                res1["errors"],
            )

            # 2. Named WACC_Bull (0.06) on the WACC tab alongside that multi-column DCF block MUST still
            #    catch 7% >= 6% (and 9% >= 6%) without falsely cross-pairing DCF!D2 against DCF!C4.
            named_wacc_path = Path(tmp) / "multi_col_bull_with_named_wacc_bull.xlsx"
            wb2 = _build_multi_col_bull(named_wacc_path)
            wb2["WACC"]["A5"] = "Bull Rate"
            wb2["WACC"]["C5"] = 0.06
            wb2.defined_names["WACC_Bull"] = DefinedName("WACC_Bull", attr_text="WACC!$C$5")
            wb2.save(named_wacc_path)
            wb2.close()

            res2 = validate_dcf_mod.validate_dcf_model(str(named_wacc_path))
            self.assertEqual(res2["status"], "FAIL", res2)
            self.assertTrue(
                any(
                    "Terminal growth (7.00% at DCF!C2) >= WACC (6.00% at WACC!C5)" in e
                    for e in res2["errors"]
                ),
                res2["errors"],
            )
            self.assertTrue(
                any(
                    "Terminal growth (9.00% at DCF!D2) >= WACC (6.00% at WACC!C5)" in e
                    for e in res2["errors"]
                ),
                res2["errors"],
            )
            self.assertFalse(
                any("DCF!D2" in e and "DCF!C4" in e for e in res2["errors"]),
                res2["errors"],
            )

            # 3. Named WACC_Bull pointing to DCF!C4 (already in closest[3]) MUST NOT cross-pair with DCF!D2.
            named_in_closest_path = Path(tmp) / "multi_col_bull_named_in_closest.xlsx"
            wb3 = _build_multi_col_bull(named_in_closest_path)
            wb3.defined_names["WACC_Bull"] = DefinedName("WACC_Bull", attr_text="DCF!$C$4")
            wb3.save(named_in_closest_path)
            wb3.close()

            res3 = validate_dcf_mod.validate_dcf_model(str(named_in_closest_path))
            self.assertEqual(res3["status"], "PASS", res3)
            self.assertEqual(res3["error_count"], 0, res3["errors"])

            # 4. Unnamed Bull Case WACC (0.06) on the WACC sheet alongside that multi-column DCF block
            #    MUST also be caught without cross-pairing DCF!D2 against DCF!C4.
            wacc_sheet_row_path = Path(tmp) / "multi_col_bull_with_wacc_sheet_row.xlsx"
            wb4 = _build_multi_col_bull(wacc_sheet_row_path)
            wb4["WACC"]["A5"] = "Bull Case WACC"
            wb4["WACC"]["C5"] = 0.06
            wb4.save(wacc_sheet_row_path)
            wb4.close()

            res4 = validate_dcf_mod.validate_dcf_model(str(wacc_sheet_row_path))
            self.assertEqual(res4["status"], "FAIL", res4)
            self.assertTrue(
                any(
                    "Terminal growth (7.00% at DCF!C2) >= WACC (6.00% at WACC!C5)" in e
                    for e in res4["errors"]
                ),
                res4["errors"],
            )
            self.assertFalse(
                any("DCF!D2" in e and "DCF!C4" in e for e in res4["errors"]),
                res4["errors"],
            )

    def test_custom_scenario_tag_pairs_with_matching_wacc_sheet_row(self) -> None:
        def _build_upside_wb(path: Path, upside_tg: float, upside_wacc: float, base_wacc: float) -> None:
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A2"] = "Upside Case Terminal Growth"
            ws_dcf["B2"] = upside_tg
            ws_dcf["A3"] = "Base Case Terminal Growth"
            ws_dcf["B3"] = 0.025
            ws_dcf["A6"] = "PV of Terminal Value"
            ws_dcf["B6"] = 600.0
            ws_dcf["A7"] = "Enterprise Value"
            ws_dcf["B7"] = 1000.0
            ws_dcf["A10"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A2"] = "Upside Case WACC"
            ws_wacc["B2"] = upside_wacc
            ws_wacc["A3"] = "Base Case WACC"
            ws_wacc["B3"] = base_wacc
            wb.save(path)
            wb.close()

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Valid custom scenario: Upside TG (11%) < Upside WACC (12%) PASSES even when Base WACC (10%) < 11%
            valid_path = Path(tmp) / "custom_upside_valid.xlsx"
            _build_upside_wb(valid_path, upside_tg=0.11, upside_wacc=0.12, base_wacc=0.10)
            res_ok = validate_dcf_mod.validate_dcf_model(str(valid_path))
            self.assertEqual(res_ok["status"], "PASS", res_ok)
            self.assertEqual(res_ok["error_count"], 0, res_ok["errors"])
            self.assertIn("✓ Terminal growth (11.00% at DCF!B2) < WACC (12.00% at WACC!B2)", res_ok["info"])
            self.assertIn("✓ Terminal growth (2.50% at DCF!B3) < WACC (10.00% at WACC!B3)", res_ok["info"])

            # 2. Invalid custom scenario: Upside TG (12.5%) >= Upside WACC (12%) FAILS even when Base WACC (15%) > 12.5%
            invalid_path = Path(tmp) / "custom_upside_invalid.xlsx"
            _build_upside_wb(invalid_path, upside_tg=0.125, upside_wacc=0.12, base_wacc=0.15)
            res_bad = validate_dcf_mod.validate_dcf_model(str(invalid_path))
            self.assertEqual(res_bad["status"], "FAIL", res_bad)
            self.assertTrue(
                any(
                    "Terminal growth (12.50% at DCF!B2) >= WACC (12.00% at WACC!B2)" in e
                    for e in res_bad["errors"]
                ),
                res_bad["errors"],
            )

    def test_terminal_growth_label_matcher_rejects_tv_pct_of_ev_and_accepts_valid_labels(self) -> None:
        is_tg = validate_dcf_mod._is_terminal_growth_label
        for positive in (
            "Terminal Growth Rate",
            "Terminal Growth",
            "Terminal Growth Rate (g)",
            "Terminal Growth (%)",
            "Perpetuity Growth Rate",
            "Perpetuity Growth",
            "Long-Term Growth Rate",
            "Long Term Growth Rate",
            "Long-Term Growth",
            "LTG",
            "PGR",
            "Bull Case Terminal Growth",
            "Bear Case Terminal Growth Rate",
            "Base Terminal Growth (g)",
            "Upside Case Terminal Growth",
            "Active Terminal Growth Rate (g)",
            "Memo Long-Term Growth",
        ):
            with self.subTest(positive=positive):
                self.assertTrue(is_tg(positive), f"Expected TG label match for {positive!r}")

        for negative in (
            "Terminal Value % of EV (Perpetuity Growth)",
            "PV of Terminal Value (Perpetuity Growth)",
            "Terminal Value (Perpetuity Growth Method)",
            "Terminal Value % of EV",
            "Implied Terminal Growth Rate",
            "Terminal Growth Sensitivity",
            "Terminal Year Revenue Growth",
            "Terminal Growth vs. GDP",
            "Terminal Growth Spread",
            "Terminal Growth Table",
        ):
            with self.subTest(negative=negative):
                self.assertFalse(is_tg(negative), f"Unexpected TG label match for {negative!r}")

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Without named ranges: TV % of EV (Perpetuity Growth) = 0.65 MUST PASS (0 errors, 0 warnings)
            plain_path = Path(tmp) / "tv_pct_perpetuity_no_names.xlsx"
            self._build_compliant_two_sheet_dcf(plain_path, with_formula=False)
            wb = openpyxl.load_workbook(plain_path)
            wb["DCF"]["A8"] = "Terminal Value % of EV (Perpetuity Growth)"
            wb["DCF"]["B8"] = 0.65
            wb.save(plain_path)
            wb.close()

            res_plain = validate_dcf_mod.validate_dcf_model(str(plain_path))
            self.assertEqual(res_plain["status"], "PASS", res_plain)
            self.assertEqual(res_plain["error_count"], 0, res_plain["errors"])
            self.assertEqual(res_plain["warning_count"], 0, res_plain["warnings"])

            # 2. With full named ranges on build_dcf_skeleton.py (real LibreOffice when available) -> PASS (0 errors, 0 warnings)
            if recalc_xlsx_mod.find_soffice() is not None:
                skel_path = Path(tmp) / "tv_pct_perpetuity_skeleton.xlsx"
                build_proc = subprocess.run(
                    [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(skel_path), "--case", "2"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

                wb_s = openpyxl.load_workbook(skel_path)
                wb_s["DCF"]["A113"] = "Terminal Value % of EV (Perpetuity Growth)"
                wb_s.save(skel_path)
                wb_s.close()

                recalc_proc = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(skel_path), "30", "--in-place"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_proc.returncode, 0, recalc_proc.stdout + recalc_proc.stderr)

                res_skel = validate_dcf_mod.validate_dcf_model(str(skel_path))
                self.assertEqual(res_skel["status"], "PASS", res_skel)
                self.assertEqual(res_skel["error_count"], 0, res_skel["errors"])
                self.assertEqual(res_skel["warning_count"], 0, res_skel["warnings"])

    def test_case_insensitive_dcf_wacc_tabs_and_optional_wacc_sheet_behavior(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Lowercase "dcf" + "wacc" sheets -> PASS (0 errors, 0 warnings)
            lower_path = Path(tmp) / "lowercase_tabs.xlsx"
            wb_l = openpyxl.Workbook()
            ws_dcf = wb_l.active
            ws_dcf.title = "dcf"
            ws_dcf["A2"] = "WACC"
            ws_dcf["B2"] = 0.095
            ws_dcf["A3"] = "Terminal Growth Rate"
            ws_dcf["B3"] = 0.025
            ws_dcf["A5"] = "PV of Terminal Value"
            ws_dcf["B5"] = 600.0
            ws_dcf["A6"] = "Enterprise Value"
            ws_dcf["B6"] = 1000.0
            ws_dcf["A10"] = "Sensitivity Analysis"

            ws_wacc = wb_l.create_sheet("wacc")
            ws_wacc["A2"] = "WEIGHTED AVERAGE COST OF CAPITAL"
            ws_wacc["B2"] = 0.095
            wb_l.save(lower_path)
            wb_l.close()

            res_lower = validate_dcf_mod.validate_dcf_model(str(lower_path))
            self.assertEqual(res_lower["status"], "PASS", res_lower)
            self.assertEqual(res_lower["error_count"], 0, res_lower["errors"])
            self.assertEqual(res_lower["warning_count"], 0, res_lower["warnings"])
            self.assertIn("Found sheet: dcf", res_lower["info"])
            self.assertIn("Found sheet: wacc", res_lower["info"])

            # 2. Single-tab "DCF" workbook with inline WACC (10%) & TG (2.5%) -> PASS (0 errors, 0 warnings)
            inline_path = Path(tmp) / "single_tab_inline_wacc.xlsx"
            wb_i = openpyxl.Workbook()
            ws_i = wb_i.active
            ws_i.title = "DCF"
            ws_i["A2"] = "WACC"
            ws_i["B2"] = 0.10
            ws_i["A3"] = "Terminal Growth Rate"
            ws_i["B3"] = 0.025
            ws_i["A5"] = "PV of Terminal Value"
            ws_i["B5"] = 600.0
            ws_i["A6"] = "Enterprise Value"
            ws_i["B6"] = 1000.0
            ws_i["A10"] = "Sensitivity Analysis"
            wb_i.save(inline_path)
            wb_i.close()

            res_inline = validate_dcf_mod.validate_dcf_model(str(inline_path))
            self.assertEqual(res_inline["status"], "PASS", res_inline)
            self.assertEqual(res_inline["error_count"], 0, res_inline["errors"])
            self.assertEqual(res_inline["warning_count"], 0, res_inline["warnings"])
            self.assertTrue(
                any("No separate WACC sheet found; using WACC from DCF sheet / named range" in i for i in res_inline["info"]),
                res_inline["info"],
            )

            # 3. Single-tab "DCF" + named DCF_WACC on third tab "Assumptions" -> PASS (0 errors, 0 warnings)
            named_wacc_path = Path(tmp) / "single_tab_named_wacc.xlsx"
            wb_n = openpyxl.Workbook()
            ws_nd = wb_n.active
            ws_nd.title = "Dcf"
            ws_nd["A3"] = "Terminal Growth Rate"
            ws_nd["B3"] = 0.025
            ws_nd["A5"] = "PV of Terminal Value"
            ws_nd["B5"] = 600.0
            ws_nd["A6"] = "Enterprise Value"
            ws_nd["B6"] = 1000.0
            ws_nd["A10"] = "Sensitivity Analysis"
            ws_assump = wb_n.create_sheet("Assumptions")
            ws_assump["C4"] = 0.095
            wb_n.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="Assumptions!$C$4")
            wb_n.save(named_wacc_path)
            wb_n.close()

            res_named = validate_dcf_mod.validate_dcf_model(str(named_wacc_path))
            self.assertEqual(res_named["status"], "PASS", res_named)
            self.assertEqual(res_named["error_count"], 0, res_named["errors"])
            self.assertEqual(res_named["warning_count"], 0, res_named["warnings"])

            # 4. Single-tab "DCF" with NO WACC anywhere -> FAIL (missing WACC error + missing WACC sheet warning)
            no_wacc_path = Path(tmp) / "single_tab_no_wacc.xlsx"
            wb_nw = openpyxl.Workbook()
            ws_nw = wb_nw.active
            ws_nw.title = "DCF"
            ws_nw["A3"] = "Terminal Growth Rate"
            ws_nw["B3"] = 0.025
            ws_nw["A5"] = "PV of Terminal Value"
            ws_nw["B5"] = 600.0
            ws_nw["A6"] = "Enterprise Value"
            ws_nw["B6"] = 1000.0
            ws_nw["A10"] = "Sensitivity Analysis"
            wb_nw.save(no_wacc_path)
            wb_nw.close()

            res_nw = validate_dcf_mod.validate_dcf_model(str(no_wacc_path))
            self.assertEqual(res_nw["status"], "FAIL", res_nw)
            self.assertTrue(
                any("Could not locate WACC value" in e for e in res_nw["errors"]),
                res_nw["errors"],
            )
            self.assertTrue(
                any("No dedicated WACC sheet found" in w for w in res_nw["warnings"]),
                res_nw["warnings"],
            )

    def test_supplementary_unnamed_rates_warn_under_full_named_coverage_while_named_and_partial_still_fail(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Skeleton (all 3 scenario TGs named) + extra unnamed memo TG = 15% -> PASS with warning;
            #    making named TerminalGrowth_Bull = 15% -> FAIL!
            if recalc_xlsx_mod.find_soffice() is not None:
                skel_path = Path(tmp) / "skeleton_memo_tg.xlsx"
                build_proc = subprocess.run(
                    [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(skel_path), "--case", "2"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

                wb = openpyxl.load_workbook(skel_path)
                wb["DCF"]["A120"] = "Memo Long-Term Growth"
                wb["DCF"]["C120"] = 0.15
                wb.save(skel_path)
                wb.close()

                recalc_proc = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(skel_path), "30", "--in-place"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_proc.returncode, 0, recalc_proc.stdout + recalc_proc.stderr)

                res_memo = validate_dcf_mod.validate_dcf_model(str(skel_path))
                self.assertEqual(res_memo["status"], "PASS", res_memo)
                self.assertEqual(res_memo["error_count"], 0, res_memo["errors"])
                self.assertTrue(
                    any("Supplementary unnamed terminal growth (15.00% at DCF!C120)" in w for w in res_memo["warnings"]),
                    res_memo["warnings"],
                )

                # Now break named TerminalGrowth_Bull (DCF!C57) = 15% -> MUST still FAIL
                wb_bad = openpyxl.load_workbook(skel_path)
                wb_bad["DCF"]["C57"] = 0.15
                wb_bad.save(skel_path)
                wb_bad.close()

                recalc_bad = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(skel_path), "30", "--in-place"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_bad.returncode, 0, recalc_bad.stdout + recalc_bad.stderr)

                res_bad = validate_dcf_mod.validate_dcf_model(str(skel_path))
                self.assertEqual(res_bad["status"], "FAIL", res_bad)
                self.assertTrue(
                    any("Terminal growth (15.00% at DCF!C57) >= WACC" in e for e in res_bad["errors"]),
                    res_bad["errors"],
                )

            # 2. Full named coverage for BOTH scenario TGs AND scenario WACCs + extra unnamed WACC row = 1% ->
            #    downgrades to warning (PASS with warning), while invalid named WACC_Base still FAILS.
            full_named_path = Path(tmp) / "full_named_with_extra_unnamed_wacc.xlsx"
            wb_f = openpyxl.Workbook()
            ws_d = wb_f.active
            ws_d.title = "DCF"
            ws_d["A1"] = "Bear TG"
            ws_d["B1"] = 0.02
            ws_d["A2"] = "Base TG"
            ws_d["B2"] = 0.025
            ws_d["A3"] = "Bull TG"
            ws_d["B3"] = 0.03
            ws_d["A4"] = "PV of Terminal Value"
            ws_d["B4"] = 600.0
            ws_d["A5"] = "Enterprise Value"
            ws_d["B5"] = 1000.0
            ws_d["A7"] = "Base Case WACC"
            ws_d["B7"] = 0.01  # Extra unnamed memo row (< Base TG 2.5%), while WACC_Bear/Base/Bull are all named
            ws_d["A12"] = "Sensitivity Analysis"

            ws_w = wb_f.create_sheet("WACC")
            ws_w["A1"] = "Bear Rate"
            ws_w["B1"] = 0.11
            ws_w["A2"] = "Base Rate"
            ws_w["B2"] = 0.09
            ws_w["A3"] = "Bull Rate"
            ws_w["B3"] = 0.08

            wb_f.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="WACC!$B$2")
            wb_f.defined_names["TerminalGrowth_Bear"] = DefinedName("TerminalGrowth_Bear", attr_text="DCF!$B$1")
            wb_f.defined_names["TerminalGrowth_Base"] = DefinedName("TerminalGrowth_Base", attr_text="DCF!$B$2")
            wb_f.defined_names["TerminalGrowth_Bull"] = DefinedName("TerminalGrowth_Bull", attr_text="DCF!$B$3")
            wb_f.defined_names["WACC_Bear"] = DefinedName("WACC_Bear", attr_text="WACC!$B$1")
            wb_f.defined_names["WACC_Base"] = DefinedName("WACC_Base", attr_text="WACC!$B$2")
            wb_f.defined_names["WACC_Bull"] = DefinedName("WACC_Bull", attr_text="WACC!$B$3")
            wb_f.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="DCF!$B$4")
            wb_f.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$B$5")
            wb_f.save(full_named_path)
            wb_f.close()

            res_full = validate_dcf_mod.validate_dcf_model(str(full_named_path))
            self.assertEqual(res_full["status"], "PASS", res_full)
            self.assertEqual(res_full["error_count"], 0, res_full["errors"])
            self.assertTrue(
                any("supplementary unnamed WACC (1.00% at DCF!B7)" in w for w in res_full["warnings"]),
                res_full["warnings"],
            )

    def test_sheet_scoped_defined_names_resolve_with_precedence_and_point_capture(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Sheet-scoped DCF_WACC on DCF + unqualified sheet-scoped TerminalGrowth on DCF +
            #    sheet-scoped name pointing to non-target "Assumptions" sheet + lower-precedence "Other" duplicate
            path = Path(tmp) / "sheet_scoped_names.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A5"] = "Custom Discount"
            ws_dcf["C5"] = 0.095
            ws_dcf["A6"] = "Custom Growth"
            ws_dcf["C6"] = 0.025
            ws_dcf["A7"] = "Custom PV TV"
            ws_dcf["C7"] = 600.0
            ws_dcf["A8"] = "Custom EV"
            ws_dcf["C8"] = 1000.0
            ws_dcf["A12"] = "Sensitivity Analysis"

            ws_assump = wb.create_sheet("Assumptions")
            ws_assump["D4"] = 0.03

            ws_other = wb.create_sheet("Other")
            ws_other["B2"] = 0.50  # Must lose to DCF-scoped DCF_WACC

            ws_dcf.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="DCF!$C$5")
            ws_dcf.defined_names["TerminalGrowth_Bear"] = DefinedName("TerminalGrowth_Bear", attr_text="$C$6")
            ws_dcf.defined_names["TerminalGrowth_Base"] = DefinedName("TerminalGrowth_Base", attr_text="DCF!$C$6")
            ws_dcf.defined_names["TerminalGrowth_Bull"] = DefinedName(
                "TerminalGrowth_Bull", attr_text="Assumptions!$D$4"
            )
            ws_dcf.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="DCF!$C$7")
            ws_dcf.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$C$8")
            ws_other.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="Other!$B$2")
            wb.save(path)
            wb.close()

            scan = xlsx_scan_mod.scan_workbook_formulas(
                path,
                capture_sheet_grids=True,
                grid_sheet_names=("DCF", "WACC"),
                capture_defined_names=True,
            )
            self.assertEqual(scan["defined_names"]["DCF_WACC"], ("DCF", 5, 3))
            self.assertEqual(scan["defined_names"]["TerminalGrowth_Bear"], ("DCF", 6, 3))
            self.assertEqual(scan["defined_names"]["TerminalGrowth_Bull"], ("Assumptions", 4, 4))
            # Non-target Assumptions sheet is point-captured cleanly
            self.assertEqual(scan["sheet_grids"]["Assumptions"], {4: {4: (0.03, "D4")}})
            self.assertNotIn("Other", scan["sheet_grids"])

            res_ok = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res_ok["status"], "PASS", res_ok)
            self.assertEqual(res_ok["error_count"], 0, res_ok["errors"])
            self.assertEqual(res_ok["warning_count"], 0, res_ok["warnings"])

            # 2. Workbook-global DCF_WACC takes precedence over DCF-scoped DCF_WACC
            wb2 = openpyxl.load_workbook(path)
            wb2["DCF"]["C9"] = 0.02  # Low global WACC (< 3.0% Bull TG -> FAIL)
            wb2.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="DCF!$C$9")
            wb2.save(path)
            wb2.close()

            scan2 = xlsx_scan_mod.scan_workbook_formulas(
                path,
                capture_sheet_grids=True,
                grid_sheet_names=("DCF", "WACC"),
                capture_defined_names=True,
            )
            self.assertEqual(scan2["defined_names"]["DCF_WACC"], ("DCF", 9, 3))

            res_prec = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res_prec["status"], "FAIL", res_prec)
            self.assertTrue(
                any("WACC (2.00% at DCF!C9)" in e for e in res_prec["errors"]),
                res_prec["errors"],
            )

    def test_leading_equals_defined_names_resolve_single_cells_and_skip_formulas_and_external_refs(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "leading_equals_defined_names.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A5"] = "Custom Rate"
            ws_dcf["C5"] = 0.095
            ws_dcf["A6"] = "Custom Growth"
            ws_dcf["C6"] = 0.025
            ws_dcf["A7"] = "Custom PV TV"
            ws_dcf["C7"] = 600.0
            ws_dcf["A8"] = "Custom EV"
            ws_dcf["C8"] = 1000.0
            ws_dcf["A12"] = "Sensitivity Analysis"

            ws_obrien = wb.create_sheet("O'Brien Inputs")
            ws_obrien["D12"] = 0.03

            # Valid leading-= single-cell refs (sheet-scoped unqualified, qualified, and apostrophe-quoted)
            ws_dcf.defined_names["ScopedAbs"] = DefinedName("ScopedAbs", attr_text="=$C$5")
            ws_dcf.defined_names["ScopedRel"] = DefinedName("ScopedRel", attr_text="=C5")
            ws_dcf.defined_names["ScopedQual"] = DefinedName("ScopedQual", attr_text="=DCF!$C$5")
            ws_dcf.defined_names["ScopedApostrophe"] = DefinedName(
                "ScopedApostrophe", attr_text="='O''Brien Inputs'!$D$12"
            )
            wb.defined_names["DCF_WACC"] = DefinedName("DCF_WACC", attr_text="=DCF!$C$5")
            wb.defined_names["TerminalGrowth_Bear"] = DefinedName(
                "TerminalGrowth_Bear", attr_text="=DCF!$C$6"
            )
            ws_dcf.defined_names["TerminalGrowth_Base"] = DefinedName(
                "TerminalGrowth_Base", attr_text="=$C$6"
            )
            wb.defined_names["TerminalGrowth_Bull"] = DefinedName(
                "TerminalGrowth_Bull", attr_text="='O''Brien Inputs'!$D$12"
            )
            wb.defined_names["PV_TerminalValue"] = DefinedName(
                "PV_TerminalValue", attr_text="=DCF!$C$7"
            )
            wb.defined_names["EnterpriseValue"] = DefinedName(
                "EnterpriseValue", attr_text="=DCF!$C$8"
            )

            # Formula expressions, multi-cell ranges, double-=, unqualified global, and external refs -> MUST skip
            wb.defined_names["FormulaAddRight"] = DefinedName("FormulaAddRight", attr_text="=DCF!$C$5+1")
            wb.defined_names["FormulaAddLeft"] = DefinedName("FormulaAddLeft", attr_text="=1+DCF!$C$5")
            wb.defined_names["FormulaSumRange"] = DefinedName(
                "FormulaSumRange", attr_text="=SUM(DCF!$C$5:$C$10)"
            )
            wb.defined_names["FormulaSumSingle"] = DefinedName(
                "FormulaSumSingle", attr_text="=SUM(DCF!$C$5)"
            )
            ws_dcf.defined_names["FormulaLocalAdd"] = DefinedName("FormulaLocalAdd", attr_text="=A1+B1")
            wb.defined_names["DoubleEquals"] = DefinedName("DoubleEquals", attr_text="==DCF!$C$5")
            wb.defined_names["GlobalUnqualified"] = DefinedName("GlobalUnqualified", attr_text="=$C$5")
            wb.defined_names["ExternalPlain"] = DefinedName("ExternalPlain", attr_text="[1]DCF!$C$5")
            wb.defined_names["ExternalLeadEq"] = DefinedName("ExternalLeadEq", attr_text="=[1]DCF!$C$5")
            wb.defined_names["ExternalQuoted"] = DefinedName(
                "ExternalQuoted", attr_text="='[Book1.xlsx]O''Brien Inputs'!$D$12"
            )
            wb.save(path)
            wb.close()

            scan = xlsx_scan_mod.scan_workbook_formulas(
                path,
                capture_sheet_grids=True,
                grid_sheet_names=("DCF", "WACC"),
                capture_defined_names=True,
            )
            dn = scan["defined_names"]
            self.assertEqual(dn["ScopedAbs"], ("DCF", 5, 3))
            self.assertEqual(dn["ScopedRel"], ("DCF", 5, 3))
            self.assertEqual(dn["ScopedQual"], ("DCF", 5, 3))
            self.assertEqual(dn["ScopedApostrophe"], ("O'Brien Inputs", 12, 4))
            self.assertEqual(dn["DCF_WACC"], ("DCF", 5, 3))
            self.assertEqual(dn["TerminalGrowth_Bear"], ("DCF", 6, 3))
            self.assertEqual(dn["TerminalGrowth_Base"], ("DCF", 6, 3))
            self.assertEqual(dn["TerminalGrowth_Bull"], ("O'Brien Inputs", 12, 4))
            # Non-target apostrophe-named sheet is point-captured cleanly via leading-= defined name
            self.assertEqual(scan["sheet_grids"]["O'Brien Inputs"], {12: {4: (0.03, "D12")}})

            for skipped_name in (
                "FormulaAddRight",
                "FormulaAddLeft",
                "FormulaSumRange",
                "FormulaSumSingle",
                "FormulaLocalAdd",
                "DoubleEquals",
                "GlobalUnqualified",
                "ExternalPlain",
                "ExternalLeadEq",
                "ExternalQuoted",
            ):
                self.assertNotIn(skipped_name, dn, f"{skipped_name} should not resolve as a single cell")

            res = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res["status"], "PASS", res)
            self.assertEqual(res["error_count"], 0, res["errors"])
            self.assertEqual(res["warning_count"], 0, res["warnings"])

    def test_named_negative_tg_with_unnamed_wacc_emits_single_critical_and_no_supplementary_warning(self) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as tmp:
            # 1. Named TerminalGrowth_Bear = -150% (DCF!B2) with Base/Bull TG unnamed (all_scenario_tg_named=False),
            #    all WACC_Bear/Base/Bull named (10%), plus an unnamed "Bear Case WACC" = 12% on DCF ->
            #    MUST FAIL with exactly ONE CRITICAL error and ZERO "Supplementary unnamed terminal growth" warnings.
            path = Path(tmp) / "named_neg_tg_unnamed_wacc.xlsx"
            wb = openpyxl.Workbook()
            ws_dcf = wb.active
            ws_dcf.title = "DCF"
            ws_dcf["A2"] = "Bear Case Terminal Growth Rate"
            ws_dcf["B2"] = -1.50
            ws_dcf["A3"] = "Base Case Terminal Growth Rate"
            ws_dcf["B3"] = 0.025
            ws_dcf["A4"] = "Bull Case Terminal Growth Rate"
            ws_dcf["B4"] = 0.03
            ws_dcf["A5"] = "PV of Terminal Value"
            ws_dcf["B5"] = 600.0
            ws_dcf["A6"] = "Enterprise Value"
            ws_dcf["B6"] = 1000.0
            ws_dcf["A7"] = "Bear Case WACC"
            ws_dcf["B7"] = 0.12
            ws_dcf["A12"] = "Sensitivity Analysis"

            ws_wacc = wb.create_sheet("WACC")
            ws_wacc["A1"] = "Bear Case WACC"
            ws_wacc["B1"] = 0.10
            ws_wacc["A2"] = "Base Case WACC"
            ws_wacc["B2"] = 0.10
            ws_wacc["A3"] = "Bull Case WACC"
            ws_wacc["B3"] = 0.10

            wb.defined_names["TerminalGrowth_Bear"] = DefinedName("TerminalGrowth_Bear", attr_text="DCF!$B$2")
            wb.defined_names["WACC_Bear"] = DefinedName("WACC_Bear", attr_text="WACC!$B$1")
            wb.defined_names["WACC_Base"] = DefinedName("WACC_Base", attr_text="WACC!$B$2")
            wb.defined_names["WACC_Bull"] = DefinedName("WACC_Bull", attr_text="WACC!$B$3")
            wb.defined_names["PV_TerminalValue"] = DefinedName("PV_TerminalValue", attr_text="DCF!$B$5")
            wb.defined_names["EnterpriseValue"] = DefinedName("EnterpriseValue", attr_text="DCF!$B$6")
            wb.save(path)
            wb.close()

            res = validate_dcf_mod.validate_dcf_model(str(path))
            self.assertEqual(res["status"], "FAIL", res)
            neg_tg_errors = [
                e for e in res["errors"] if "CRITICAL: Terminal growth (-150.00% at DCF!B2) is <= -100%" in e
            ]
            self.assertEqual(len(neg_tg_errors), 1, res["errors"])
            self.assertFalse(
                any("Supplementary unnamed terminal growth" in w for w in res["warnings"]),
                res["warnings"],
            )

            # 2. When all_scenario_wacc_named=False and named TerminalGrowth_Bear = -150% pairs against
            #    multiple WACCs (10% + 12%), CRITICAL error is still deduplicated to exactly ONE entry.
            dedupe_path = Path(tmp) / "named_neg_tg_multi_wacc_dedupe.xlsx"
            wb2 = openpyxl.load_workbook(path)
            del wb2.defined_names["WACC_Bull"]
            wb2.save(dedupe_path)
            wb2.close()

            res2 = validate_dcf_mod.validate_dcf_model(str(dedupe_path))
            self.assertEqual(res2["status"], "FAIL", res2)
            neg_tg_errors2 = [
                e for e in res2["errors"] if "CRITICAL: Terminal growth (-150.00% at DCF!B2) is <= -100%" in e
            ]
            self.assertEqual(len(neg_tg_errors2), 1, res2["errors"])
            self.assertFalse(
                any("Supplementary unnamed terminal growth" in w for w in res2["warnings"]),
                res2["warnings"],
            )

            # 3. Direct helper check: unnamed TG <= -100% with all_scenario_tg_named=True paired against
            #    multiple WACCs emits exactly ONE supplementary warning (deduplicated by tg_coord) and zero errors.
            validator = validate_dcf_mod.DCFModelValidator(str(path))
            validator.errors.clear()
            validator.warnings.clear()
            validator.info.clear()
            validator._emit_tg_vs_wacc_results(
                [
                    (-1.50, "DCF!B9", 0.10, "WACC!B1", True, False),
                    (-1.50, "DCF!B9", 0.12, "DCF!B7", True, True),
                ],
                all_scenario_tg_named=True,
                all_scenario_wacc_named=True,
            )
            self.assertEqual(validator.errors, [])
            supp_neg_warnings = [
                w
                for w in validator.warnings
                if "Supplementary unnamed terminal growth (-150.00% at DCF!B9) is <= -100%" in w
            ]
            self.assertEqual(len(supp_neg_warnings), 1, validator.warnings)


@_skip_unless_required(openpyxl is None, "openpyxl is required for DCF skeleton tests")
class TestBuildDcfSkeleton(unittest.TestCase):
    @_skip_unless_required(
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

            # Inspect formulas & DataValidation before recalculation
            wb_f = openpyxl.load_workbook(model_path, data_only=False)
            try:
                ws_f = wb_f["DCF"]
                ws_wacc_f = wb_f["WACC"]
                # B4 DataValidation enforcing whole numbers 1..3
                dvs = list(ws_f.data_validations.dataValidation)
                b4_dvs = [dv for dv in dvs if "B4" in str(dv.sqref)]
                self.assertEqual(len(b4_dvs), 1, f"Expected DataValidation on DCF!B4, got {dvs}")
                self.assertEqual(b4_dvs[0].type, "whole")
                self.assertEqual(b4_dvs[0].operator, "between")
                self.assertEqual(str(b4_dvs[0].formula1), "1")
                self.assertEqual(str(b4_dvs[0].formula2), "3")
                # INT(Case) guard in CHOOSE and ALL Active Case INDEX formulas
                self.assertIn("INT(Case)", str(ws_f["C4"].value))
                for act_row in range(62, 69):
                    for act_col in range(3, 8):
                        self.assertIn("INT(Case)", str(ws_f.cell(row=act_row, column=act_col).value))
                self.assertIn("INT(Case)", str(ws_f["C69"].value))
                self.assertIn("INT(Case)", str(ws_f["C70"].value))
                # Skeleton has zero LibreOffice-lossy features (comment VML must not trigger drawings_or_shapes)
                self.assertEqual(xlsx_scan_mod.inspect_workbook_xml(model_path)["lossy_features"], [])
                # WACC link synchronized with WACC!C30 and Preferred Stock linked from DCF!C16
                self.assertEqual(str(ws_f["C21"].value), "=WACC!C30")
                self.assertEqual(str(ws_f["C22"].value), "=WACC!C14")
                self.assertEqual(str(ws_wacc_f["C22"].value), "=DCF!C16")
                # Perpetuity TV discounted at N - 0.5 ($G$87), Exit Multiple TV at N ($G$87+0.5),
                # and Implied Exit EV/EBITDA compounds mid-year Perpetuity TV half a year to t = N
                self.assertIn("^$G$87", str(ws_f["C96"].value))
                self.assertEqual(str(ws_f["C115"].value), "=C95*(1+$C$21)^0.5/G83")
                self.assertIn("^($G$87+0.5)", str(ws_f["C117"].value))
                # Horizon-independent IF(WACC<=g,"NM", SUMPRODUCT(...)) sensitivity formula
                sens_formula = str(ws_f["E127"].value)
                self.assertTrue(sens_formula.startswith('=IF($B127<=E$124,"NM",'), sens_formula)
                self.assertIn("SUMPRODUCT(", sens_formula)
                self.assertIn("$G$87", sens_formula)
                self.assertNotIn("^0.5", sens_formula)
                self.assertNotIn("^4.5", sens_formula)
                self.assertNotIn("^5", sens_formula)
            finally:
                wb_f.close()

            prices_by_case: dict[int, float] = {}
            expected_labels = {1: "Bear Case", 2: "Base Case", 3: "Bull Case"}

            for case_num in (2, 1, 3):
                if case_num != 2:
                    # Switch named Case cell DCF!B4 directly and re-save
                    wb_edit = openpyxl.load_workbook(model_path)
                    wb_edit["DCF"]["B4"] = case_num
                    wb_edit.save(model_path)
                    wb_edit.close()

                # Recalculate with real LibreOffice (--in-place populates cached values on disk)
                recalc_proc = subprocess.run(
                    [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30", "--in-place"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(recalc_proc.returncode, 0, recalc_proc.stdout + recalc_proc.stderr)
                recalc_payload = json.loads(recalc_proc.stdout)
                self.assertEqual(recalc_payload["status"], "success", recalc_payload)
                self.assertEqual(recalc_payload["mode"], "in_place")
                self.assertTrue(recalc_payload["modified_file"])
                self.assertTrue(recalc_payload["recalculated_with_libreoffice"])
                self.assertEqual(recalc_payload["total_errors"], 0)
                self.assertEqual(recalc_payload["uncached_formulas"], 0)
                self.assertEqual(recalc_payload["data_table_formulas"], 0)
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
                    # Gross-capitalization weights (E + D + P): all non-negative and sum to 100%
                    self.assertGreater(ws_wacc["C24"].value, 0.0)
                    self.assertGreaterEqual(ws_wacc["C25"].value, 0.0)
                    self.assertGreaterEqual(ws_wacc["C26"].value, 0.0)
                    self.assertAlmostEqual(ws_wacc["C27"].value, 1.0, places=6)
                    self.assertAlmostEqual(float(ws_dcf["C21"].value), float(ws_wacc["C30"].value), places=8)
                    # TV % of EV cleanly inside 50%-70%
                    tv_pct = float(ws_dcf["C113"].value)
                    self.assertGreaterEqual(tv_pct, 0.50)
                    self.assertLessEqual(tv_pct, 0.70)
                    # Perpetuity TV DF = 1/(1+WACC)^4.5, Exit Multiple PV = ExitTV/(1+WACC)^5.0,
                    # and Implied Exit Multiple compounds mid-year Perpetuity TV half a year to t = N
                    wacc = float(ws_dcf["C21"].value)
                    self.assertAlmostEqual(float(ws_dcf["C96"].value), 1.0 / ((1.0 + wacc) ** 4.5), places=8)
                    expected_impl_exit = float(ws_dcf["C95"].value) * ((1.0 + wacc) ** 0.5) / float(ws_dcf["G83"].value)
                    self.assertLess(abs(float(ws_dcf["C115"].value) - expected_impl_exit), 1e-6)
                    self.assertAlmostEqual(
                        float(ws_dcf["C117"].value),
                        float(ws_dcf["C116"].value) / ((1.0 + wacc) ** 5.0),
                        places=6,
                    )
                    # All 8 workbook-level named ranges survive LibreOffice recalculation cleanly
                    expected_names = {
                        "Case": ("DCF", "$B$4"),
                        "DCF_WACC": ("WACC", "$C$30"),
                        "TerminalGrowth": ("DCF", "$C$69"),
                        "TerminalGrowth_Bear": ("DCF", "$C$33"),
                        "TerminalGrowth_Base": ("DCF", "$C$45"),
                        "TerminalGrowth_Bull": ("DCF", "$C$57"),
                        "PV_TerminalValue": ("DCF", "$C$99"),
                        "EnterpriseValue": ("DCF", "$C$100"),
                    }
                    for dn_name, expected_dest in expected_names.items():
                        self.assertIn(dn_name, wb_val.defined_names)
                        self.assertEqual(list(wb_val.defined_names[dn_name].destinations), [expected_dest])

                    # Center cell of SUMPRODUCT 5x5 sensitivity matches Active Case Implied Share Price
                    implied_price = float(ws_dcf["C108"].value)
                    sens_center = float(ws_dcf["E127"].value)
                    self.assertLess(abs(implied_price - sens_center), 1e-6)
                    prices_by_case[case_num] = implied_price
                finally:
                    wb_val.close()

                # Re-exported skeleton still has zero lossy_features (vmlDrawing*.vml excluded)
                self.assertEqual(xlsx_scan_mod.inspect_workbook_xml(model_path)["lossy_features"], [])

                # Verify DCFModelValidator resolves core and per-scenario named metrics directly
                validator_obj = validate_dcf_mod.DCFModelValidator(str(model_path))
                self.assertEqual(
                    set(validator_obj._named_metrics.keys()),
                    {"wacc", "terminal_growth", "pv_terminal_value", "enterprise_value"},
                )
                self.assertEqual(
                    set(validator_obj._named_scenario_tg.keys()),
                    {"bear", "base", "bull"},
                )

            self.assertLess(prices_by_case[1], prices_by_case[2])
            self.assertLess(prices_by_case[2], prices_by_case[3])

            # Test fractional B4 = 2.5 does NOT misalign CHOOSE and INDEX (matches Base Case exactly)
            wb_frac = openpyxl.load_workbook(model_path)
            wb_frac["DCF"]["B4"] = 2.5
            wb_frac.save(model_path)
            wb_frac.close()
            recalc_frac = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30", "--in-place"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(recalc_frac.returncode, 0, recalc_frac.stdout + recalc_frac.stderr)
            wb_frac_v = openpyxl.load_workbook(model_path, data_only=True)
            try:
                self.assertEqual(wb_frac_v["DCF"]["C4"].value, "Base Case")
                self.assertAlmostEqual(float(wb_frac_v["DCF"]["C108"].value), prices_by_case[2], places=6)
            finally:
                wb_frac_v.close()

            # Test OTM options ($14 < $15 strike): Unvested RSUs MUST remain 100% included (96.0 + 1.5 = 97.5)
            wb_otm = openpyxl.load_workbook(model_path)
            wb_otm["DCF"]["B4"] = 2
            wb_otm["DCF"]["C8"] = 14.00
            wb_otm.save(model_path)
            wb_otm.close()
            recalc_otm = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30", "--in-place"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(recalc_otm.returncode, 0, recalc_otm.stdout + recalc_otm.stderr)
            wb_otm_v = openpyxl.load_workbook(model_path, data_only=True)
            try:
                ws_otm = wb_otm_v["DCF"]
                self.assertAlmostEqual(float(ws_otm["C13"].value), 97.5, places=6)
            finally:
                wb_otm_v.close()

    @_skip_unless_required(
        recalc_xlsx_mod.find_soffice() is None,
        "LibreOffice (soffice) is required for sensitivity WACC <= g guard integration test",
    )
    def test_sensitivity_grid_wacc_le_g_evaluates_to_nm_without_errors(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "low_wacc_dcf.xlsx"
            build_proc = subprocess.run(
                [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(model_path)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

            # Case 1: WACC = 5.5% and Base Terminal Growth = 3.5%
            # Corner cell G125 has WACC = 4.5% == g = 4.5% -> must evaluate to "NM" (no #DIV/0!)
            wb = openpyxl.load_workbook(model_path)
            ws_wacc = wb["WACC"]
            ws_wacc["C6"] = 0.025   # Rf = 2.5%
            ws_wacc["C7"] = 0.50    # Beta = 0.5 -> Ke = 2.5% + 0.5 * 6.0% = 5.5%
            ws_wacc["C8"] = 0.060   # ERP = 6.0%
            ws_wacc["C13"] = 0.055  # Kd = 5.5%
            ws_wacc["C14"] = 0.0    # Tax = 0% -> after-tax Kd = 5.5%
            ws_wacc["C16"] = 0.055  # Kp = 5.5% -> blended WACC = exactly 5.50%
            wb["DCF"]["C45"] = 0.035  # Base Case Terminal Growth = 3.5%
            wb.save(model_path)
            wb.close()

            recalc_eq = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30", "--in-place"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(recalc_eq.returncode, 0, recalc_eq.stdout + recalc_eq.stderr)
            payload_eq = json.loads(recalc_eq.stdout)
            self.assertEqual(payload_eq["status"], "success", payload_eq)
            self.assertEqual(payload_eq["total_errors"], 0, payload_eq)

            wb_eq_v = openpyxl.load_workbook(model_path, data_only=True)
            try:
                ws_dcf_v = wb_eq_v["DCF"]
                self.assertAlmostEqual(float(ws_dcf_v["C21"].value), 0.055, places=6)
                self.assertEqual(ws_dcf_v["G125"].value, "NM")
                self.assertAlmostEqual(float(ws_dcf_v["E127"].value), float(ws_dcf_v["C108"].value), places=6)
            finally:
                wb_eq_v.close()

            # Case 2: WACC = 5.0% and Base Terminal Growth = 3.5%
            # G125 has WACC 4.0% < g 4.5%, F125 has 4.0% == 4.0%, G126 has 4.5% == 4.5% -> all "NM"
            wb2 = openpyxl.load_workbook(model_path)
            ws_wacc2 = wb2["WACC"]
            ws_wacc2["C6"] = 0.020   # Ke = 2.0% + 0.5 * 6.0% = 5.0%
            ws_wacc2["C13"] = 0.050  # Kd = 5.0%
            ws_wacc2["C16"] = 0.050  # Kp = 5.0% -> blended WACC = exactly 5.00%
            wb2.save(model_path)
            wb2.close()

            recalc_lt = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30", "--in-place"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(recalc_lt.returncode, 0, recalc_lt.stdout + recalc_lt.stderr)
            payload_lt = json.loads(recalc_lt.stdout)
            self.assertEqual(payload_lt["status"], "success", payload_lt)
            self.assertEqual(payload_lt["total_errors"], 0, payload_lt)

            wb_lt_v = openpyxl.load_workbook(model_path, data_only=True)
            try:
                ws_dcf_lt = wb_lt_v["DCF"]
                self.assertAlmostEqual(float(ws_dcf_lt["C21"].value), 0.050, places=6)
                self.assertEqual(ws_dcf_lt["F125"].value, "NM")
                self.assertEqual(ws_dcf_lt["G125"].value, "NM")
                self.assertEqual(ws_dcf_lt["G126"].value, "NM")
                self.assertIsInstance(ws_dcf_lt["E127"].value, (int, float))
                self.assertAlmostEqual(float(ws_dcf_lt["E127"].value), float(ws_dcf_lt["C108"].value), places=6)
            finally:
                wb_lt_v.close()

    @_skip_unless_required(
        recalc_xlsx_mod.find_soffice() is None,
        "LibreOffice (soffice) is required for Preferred Stock WACC integration test",
    )
    def test_wacc_blends_preferred_stock_and_collapses_when_zero(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "pref_wacc_dcf.xlsx"
            build_proc = subprocess.run(
                [sys.executable, str(BUILD_DCF_SKELETON_PATH), "--output", str(model_path)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(build_proc.returncode, 0, build_proc.stdout + build_proc.stderr)

            # 1. Material Preferred Stock ($200M) with Kp = 7.0% blends all 3 capital tranches
            wb = openpyxl.load_workbook(model_path)
            wb["DCF"]["C16"] = 200.0
            wb["WACC"]["C16"] = 0.07
            wb.save(model_path)
            wb.close()

            recalc_pref = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30", "--in-place"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(recalc_pref.returncode, 0, recalc_pref.stdout + recalc_pref.stderr)

            val_pref = validate_dcf_mod.validate_dcf_model(str(model_path))
            self.assertEqual(val_pref["status"], "PASS", val_pref)
            self.assertEqual(val_pref["error_count"], 0, val_pref["errors"])

            wb_v = openpyxl.load_workbook(model_path, data_only=True)
            try:
                ws_w = wb_v["WACC"]
                ws_d = wb_v["DCF"]
                ke = float(ws_w["C9"].value)
                kd_at = float(ws_w["C15"].value)
                kp = float(ws_w["C16"].value)
                e_val = float(ws_w["C20"].value)
                d_val = float(ws_w["C21"].value)
                p_val = float(ws_w["C22"].value)
                tot_cap = e_val + d_val + p_val

                self.assertAlmostEqual(p_val, 200.0, places=6)
                self.assertAlmostEqual(kp, 0.07, places=6)
                self.assertAlmostEqual(float(ws_w["C23"].value), tot_cap, places=6)
                self.assertGreater(float(ws_w["C26"].value), 0.0)
                self.assertAlmostEqual(float(ws_w["C26"].value), p_val / tot_cap, places=8)
                self.assertAlmostEqual(float(ws_w["C27"].value), 1.0, places=8)
                expected_wacc = (e_val / tot_cap) * ke + (d_val / tot_cap) * kd_at + (p_val / tot_cap) * kp
                self.assertAlmostEqual(float(ws_w["C30"].value), expected_wacc, places=8)
                self.assertAlmostEqual(float(ws_d["C21"].value), expected_wacc, places=8)
            finally:
                wb_v.close()

            # 2. Zero Preferred Stock ($0M, Kp = 0.0%) collapses cleanly to 2-tranche E + D WACC
            wb0 = openpyxl.load_workbook(model_path)
            wb0["DCF"]["C16"] = 0.0
            wb0["WACC"]["C16"] = 0.0
            wb0.save(model_path)
            wb0.close()

            recalc_zero = subprocess.run(
                [sys.executable, str(RECALC_XLSX_PATH), str(model_path), "30", "--in-place"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(recalc_zero.returncode, 0, recalc_zero.stdout + recalc_zero.stderr)

            wb0_v = openpyxl.load_workbook(model_path, data_only=True)
            try:
                ws_w0 = wb0_v["WACC"]
                ke0 = float(ws_w0["C9"].value)
                kd_at0 = float(ws_w0["C15"].value)
                e0 = float(ws_w0["C20"].value)
                d0 = float(ws_w0["C21"].value)
                expected_wacc0 = (e0 / (e0 + d0)) * ke0 + (d0 / (e0 + d0)) * kd_at0
                self.assertAlmostEqual(float(ws_w0["C26"].value), 0.0, places=8)
                self.assertAlmostEqual(float(ws_w0["C27"].value), 1.0, places=8)
                self.assertAlmostEqual(float(ws_w0["C30"].value), expected_wacc0, places=8)
            finally:
                wb0_v.close()


if __name__ == "__main__":
    unittest.main()
