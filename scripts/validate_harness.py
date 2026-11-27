#!/usr/bin/env python3
"""Validate the Google Antigravity equity research harness structure and frontmatter.

Usage:
    python3 scripts/validate_harness.py
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
AGENTS_DIR = ROOT / ".agents"

KEBAB_RE = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")
VALID_TRIGGERS = {"always_on", "model_decision", "glob", "manual"}
VALID_MODELS = {"inherit", "flash_lite", "flash", "pro"}
VALID_CMD_POLICIES = {"off", "auto", "eager", "sandbox"}
STALE_PATTERNS = re.compile(
    r"(claude's built-in|mcp__office|cowork plugin|cma mode|office js|(?<![a-z0-9_])recalc\.py)",
    re.IGNORECASE,
)


def parse_frontmatter(path: Path) -> tuple[dict[str, str], str, list[str]]:
    text = path.read_text(encoding="utf-8")
    errors: list[str] = []
    if not text.startswith("---\n"):
        return {}, text, [f"{path.relative_to(ROOT)}: missing YAML frontmatter block"]
    end = text.find("\n---\n", 4)
    if end == -1:
        end = text.find("\n---", 4)
        if end == -1:
            return {}, text, [f"{path.relative_to(ROOT)}: unterminated YAML frontmatter"]
    raw_fm = text[4:end]
    body = text[end + 4 :].strip()

    data: dict[str, str] = {}
    current_key: str | None = None
    for line in raw_fm.splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        if re.match(r"^[A-Za-z0-9_-]+\s*:", line):
            key, val = line.split(":", 1)
            key = key.strip()
            val = val.strip()
            data[key] = val.strip('"').strip("'")
            current_key = key
        elif current_key and (line.startswith(" ") or line.startswith("\t")):
            if line.lstrip().startswith("- "):
                data[current_key] = (data[current_key] + " [LIST_ITEM]").strip()
            else:
                data[current_key] = (data[current_key] + " " + line.strip()).strip()
    return data, body, errors


def main() -> int:
    errors: list[str] = []
    warnings: list[str] = []

    # 1. Root AGENTS.md
    agents_md = ROOT / "AGENTS.md"
    if not agents_md.exists():
        errors.append("Missing root AGENTS.md")
    else:
        size = agents_md.stat().st_size
        if size > 24_000:
            errors.append(f"AGENTS.md exceeds 24,000-byte Antigravity rule cap ({size} bytes)")
        if agents_md.read_text(encoding="utf-8").startswith("---\n"):
            errors.append("AGENTS.md should not contain YAML frontmatter")

    # 2. Rules (.agents/rules/*.md)
    rules_dir = AGENTS_DIR / "rules"
    rule_files = sorted(rules_dir.glob("*.md")) if rules_dir.exists() else []
    if not rule_files:
        errors.append("No rule files found in .agents/rules/")
    for rf in rule_files:
        if rf.stat().st_size > 24_000:
            errors.append(f"{rf.relative_to(ROOT)} exceeds 24,000-byte rule limit")
        fm, body, fm_errs = parse_frontmatter(rf)
        errors.extend(fm_errs)
        trigger = fm.get("trigger", "")
        if trigger not in VALID_TRIGGERS:
            errors.append(f"{rf.relative_to(ROOT)}: invalid trigger '{trigger}' (expected one of {sorted(VALID_TRIGGERS)})")
        if trigger == "model_decision" and not fm.get("description"):
            errors.append(f"{rf.relative_to(ROOT)}: model_decision rule requires 'description'")
        if trigger == "glob":
            globs = fm.get("globs") or fm.get("glob") or ""
            if not globs:
                errors.append(f"{rf.relative_to(ROOT)}: glob rule requires 'globs'")
            if "[LIST_ITEM]" in globs:
                errors.append(f"{rf.relative_to(ROOT)}: 'globs' must be a comma-separated string, not a YAML array")
        if not body:
            errors.append(f"{rf.relative_to(ROOT)}: empty rule body")

    # 3. Skills (.agents/skills/<name>/SKILL.md)
    skills_dir = AGENTS_DIR / "skills"
    skill_dirs = sorted(p for p in skills_dir.iterdir() if p.is_dir()) if skills_dir.exists() else []
    if not skill_dirs:
        errors.append("No skill directories found in .agents/skills/")
    for sd in skill_dirs:
        skill_md = sd / "SKILL.md"
        if not skill_md.exists():
            errors.append(f"{sd.relative_to(ROOT)}: missing SKILL.md")
            continue
        fm, body, fm_errs = parse_frontmatter(skill_md)
        errors.extend(fm_errs)
        name = fm.get("name", "")
        desc = fm.get("description", "")
        if not name or not KEBAB_RE.match(name) or len(name) > 64:
            errors.append(f"{skill_md.relative_to(ROOT)}: invalid skill name '{name}'")
        if name and name != sd.name:
            errors.append(f"{skill_md.relative_to(ROOT)}: frontmatter name '{name}' != folder '{sd.name}'")
        if not desc:
            errors.append(f"{skill_md.relative_to(ROOT)}: missing description")
        elif len(desc) > 1024:
            errors.append(f"{skill_md.relative_to(ROOT)}: description exceeds 1,024 chars ({len(desc)})")
        if not body:
            errors.append(f"{skill_md.relative_to(ROOT)}: empty SKILL.md body")

    # 4. Custom Agents (.agents/agents/*.md)
    custom_agents_dir = AGENTS_DIR / "agents"
    agent_files = sorted(custom_agents_dir.glob("*.md")) if custom_agents_dir.exists() else []
    if not agent_files:
        errors.append("No custom agents found in .agents/agents/")
    for af in agent_files:
        fm, body, fm_errs = parse_frontmatter(af)
        errors.extend(fm_errs)
        name = fm.get("name", "")
        if not name or not KEBAB_RE.match(name):
            errors.append(f"{af.relative_to(ROOT)}: invalid agent name '{name}'")
        if name and name != af.stem:
            errors.append(f"{af.relative_to(ROOT)}: agent name '{name}' != filename stem '{af.stem}'")
        if not fm.get("description"):
            errors.append(f"{af.relative_to(ROOT)}: missing description")
        model = fm.get("model", "inherit")
        if model not in VALID_MODELS:
            errors.append(f"{af.relative_to(ROOT)}: invalid model '{model}'")
        policy = fm.get("commandExecutionPolicy", "sandbox")
        if policy not in VALID_CMD_POLICIES:
            errors.append(f"{af.relative_to(ROOT)}: invalid commandExecutionPolicy '{policy}'")
        if not body:
            errors.append(f"{af.relative_to(ROOT)}: empty agent system prompt")

    # 5. Stale reference scan across .agents/
    for path in sorted(AGENTS_DIR.rglob("*")):
        if path.is_file() and path.suffix in (".md", ".txt", ".py"):
            content = path.read_text(encoding="utf-8", errors="ignore")
            for idx, line in enumerate(content.splitlines(), 1):
                if STALE_PATTERNS.search(line):
                    errors.append(
                        f"{path.relative_to(ROOT)}:{idx}: stale non-Antigravity reference: {line.strip()[:100]}"
                    )

    # 6. Validate MCP example JSON
    mcp_example = ROOT / "config" / "mcp_config.example.json"
    if mcp_example.exists():
        try:
            parsed = json.loads(mcp_example.read_text(encoding="utf-8"))
            if "mcpServers" not in parsed or not isinstance(parsed["mcpServers"], dict):
                errors.append("config/mcp_config.example.json missing 'mcpServers' object")
        except json.JSONDecodeError as exc:
            errors.append(f"config/mcp_config.example.json invalid JSON: {exc}")
    else:
        warnings.append("config/mcp_config.example.json not found")

    print(
        f"Validated Antigravity harness: {len(rule_files)} rules, "
        f"{len(skill_dirs)} skills, {len(agent_files)} custom agents."
    )
    for w in warnings:
        print(f"WARNING: {w}")
    for e in errors:
        print(f"ERROR: {e}")

    if errors:
        print(f"\nFAILED with {len(errors)} error(s).")
        return 1
    print("\nPASSED with 0 errors.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
