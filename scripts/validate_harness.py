#!/usr/bin/env python3
"""Validate the Google Antigravity equity research harness structure and frontmatter.

Usage:
    python3 scripts/validate_harness.py
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any
from urllib.parse import unquote

try:
    import yaml  # type: ignore
except ImportError:
    yaml = None

ROOT = Path(__file__).resolve().parent.parent

KEBAB_RE = re.compile(r"^[a-z][a-z0-9]*(-[a-z0-9]+)*$")
VALID_TRIGGERS = {"always_on", "model_decision", "glob", "manual"}
VALID_MODELS = {"inherit", "flash", "pro"}
VALID_CMD_POLICIES = {"off", "auto", "eager", "sandbox"}
ALLOWED_RULE_KEYS = {"trigger", "description", "globs", "glob"}
ALLOWED_AGENT_KEYS = {
    "name",
    "description",
    "tools",
    "mainAgent",
    "subagent",
    "model",
    "commandExecutionPolicy",
    "mcpServers",
    "skills",
    "plugins",
}
SINGLE_RULE_MAX_BYTES = 24_000
AGGREGATE_ALWAYS_ON_MAX_BYTES = 60_000

STALE_PATTERNS = re.compile(
    r"("
    r"claude's built-in"
    r"|mcp__office"
    r"|cowork plugin"
    r"|cma mode"
    r"|office js"
    r"|(?<![a-z0-9_])recalc\.py"
    r"|ask_user_question"
    r"|\bRead tool\b"
    r"|\bDOCX skill\b"
    r"|\bDOC skill\b"
    r"|\bXLSX skill\b"
    r"|\bXLS skill\b"
    r"|\bxlsx skill\b"
    r"|PowerPoint add-in"
    r"|currently selected range"
    r"|active sheet"
    r")",
    re.IGNORECASE,
)

MD_LINK_RE = re.compile(r"(?<!!)\[[^\]]+\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")
CODE_SPAN_RE = re.compile(r"(?<!`)(`+)(?!`).+?(?<!`)\1(?!`)")


def _strip_inline_comment(raw_val: str) -> str:
    """Strip YAML inline comments (# preceded by whitespace outside quotes) consistently with PyYAML."""
    in_single = False
    in_double = False
    escaped = False
    for idx, ch in enumerate(raw_val):
        if escaped:
            escaped = False
            continue
        if in_double:
            if ch == "\\":
                escaped = True
            elif ch == '"':
                in_double = False
            continue
        if in_single:
            if ch == "'":
                # YAML escapes single quotes inside single-quoted scalars via ''
                if idx + 1 < len(raw_val) and raw_val[idx + 1] == "'":
                    escaped = True
                else:
                    in_single = False
            continue
        if ch == '"':
            in_double = True
        elif ch == "'":
            in_single = True
        elif ch == "#" and (idx == 0 or raw_val[idx - 1] in (" ", "\t")):
            return raw_val[:idx].rstrip()
    return raw_val.rstrip()


def _strip_matching_quotes(raw_val: str, rel_path: str, line_no: int, errors: list[str]) -> str:
    if len(raw_val) >= 1 and raw_val[0] in ("'", '"'):
        quote = raw_val[0]
        if len(raw_val) < 2 or not raw_val.endswith(quote):
            errors.append(f"{rel_path}: frontmatter line {line_no}: unterminated quoted string")
            return raw_val
        return raw_val[1:-1]
    return raw_val


def _parse_scalar_fallback(raw_val: str, rel_path: str, line_no: int, errors: list[str]) -> Any:
    raw_val = _strip_inline_comment(raw_val).strip()
    if not raw_val:
        return ""
    if raw_val[0] in ("'", '"'):
        return _strip_matching_quotes(raw_val, rel_path, line_no, errors)
    if raw_val.startswith("[") and raw_val.endswith("]"):
        inner = raw_val[1:-1].strip()
        if not inner:
            return []
        return [item.strip().strip('"').strip("'") for item in inner.split(",")]
    # Unquoted colon followed by space/EOL inside a scalar value is invalid YAML
    if re.search(r":(\s|$)", raw_val):
        errors.append(
            f"{rel_path}: frontmatter line {line_no}: unquoted colon in scalar value ({raw_val!r})"
        )
    low = raw_val.lower()
    if low == "true":
        return True
    if low == "false":
        return False
    return raw_val


def _parse_yaml_fallback(raw_fm: str, rel_path: str) -> tuple[dict[str, Any], list[str]]:
    errors: list[str] = []
    data: dict[str, Any] = {}
    current_key: str | None = None
    lines = raw_fm.splitlines()

    non_empty = [ln for ln in lines if ln.strip() and not ln.lstrip().startswith("#")]
    if not non_empty:
        return {}, [f"{rel_path}: empty YAML frontmatter"]
    if non_empty[0].lstrip().startswith("- "):
        return {}, [f"{rel_path}: YAML frontmatter must be a mapping, not a list"]

    for idx, line in enumerate(lines, 1):
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        if not line.startswith((" ", "\t")):
            if ":" not in line:
                errors.append(f"{rel_path}: frontmatter line {idx}: invalid YAML syntax ({line.strip()!r})")
                continue
            key_match = re.match(r"^([A-Za-z0-9_-]+)\s*:(.*)$", line)
            if not key_match:
                errors.append(f"{rel_path}: frontmatter line {idx}: invalid key syntax ({line.strip()!r})")
                continue
            key = key_match.group(1).strip()
            raw_val = _strip_inline_comment(key_match.group(2).strip()).strip()
            if raw_val in ("|", ">", "|-", ">-"):
                data[key] = ""
            elif raw_val == "":
                data[key] = ""
            else:
                data[key] = _parse_scalar_fallback(raw_val, rel_path, idx, errors)
            current_key = key
        else:
            if current_key is None:
                errors.append(f"{rel_path}: frontmatter line {idx}: unexpected indentation")
                continue
            stripped = line.lstrip()
            if stripped.startswith("- "):
                item_val = _parse_scalar_fallback(stripped[2:].strip(), rel_path, idx, errors)
                if not isinstance(data[current_key], list):
                    data[current_key] = []
                data[current_key].append(item_val)
            else:
                if isinstance(data[current_key], list):
                    errors.append(
                        f"{rel_path}: frontmatter line {idx}: invalid continuation inside list"
                    )
                else:
                    cleaned_cont = _strip_inline_comment(stripped).strip()
                    prev = str(data[current_key])
                    data[current_key] = f"{prev} {cleaned_cont}".strip() if prev else cleaned_cont

    return data, errors


def parse_frontmatter(path: Path, root: Path = ROOT) -> tuple[dict[str, Any], str, list[str]]:
    rel = str(path.relative_to(root))
    text = path.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\r", "\n")
    if not text.startswith("---\n"):
        return {}, text, [f"{rel}: missing YAML frontmatter block"]
    end = text.find("\n---\n", 4)
    body_start = end + 5
    if end == -1:
        if text.endswith("\n---"):
            end = len(text) - 4
            body_start = len(text)
        else:
            return {}, text, [f"{rel}: unterminated YAML frontmatter"]
    raw_fm = text[4:end]
    body = text[body_start:].strip()

    if yaml is not None:
        try:
            parsed = yaml.safe_load(raw_fm)
        except yaml.YAMLError as exc:
            return {}, body, [f"{rel}: invalid YAML frontmatter: {exc}"]
        if not isinstance(parsed, dict):
            return {}, body, [f"{rel}: YAML frontmatter must be a mapping"]
        return parsed, body, []

    data, errors = _parse_yaml_fallback(raw_fm, rel)
    return data, body, errors


def _iter_markdown_links(text: str) -> list[tuple[int, str]]:
    """Extract (line_no, target) from markdown text, ignoring fenced code blocks and inline code."""
    links: list[tuple[int, str]] = []
    in_fence = False
    for line_no, line in enumerate(text.splitlines(), 1):
        stripped = line.lstrip()
        if stripped.startswith("```") or stripped.startswith("~~~"):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        # Replace inline code spans with a non-empty sentinel so code-formatted link labels
        # like [`guide`](docs/guide.md) are still validated while links wholly inside code spans
        # like `[a](b)` or ``[`a`](b)`` are ignored.
        cleaned = CODE_SPAN_RE.sub("\x00", line)
        for match in MD_LINK_RE.finditer(cleaned):
            target = match.group(1)
            if "\x00" in target:
                continue
            links.append((line_no, target))
    return links


def validate_markdown_links(root: Path) -> list[str]:
    """Validate relative markdown links in AGENTS.md, README.md, research/README.md, docs/*.md, and .agents/."""
    errors: list[str] = []
    root_resolved = root.resolve()
    md_files: list[Path] = []
    for top_file in ("AGENTS.md", "README.md", "research/README.md"):
        p = root / top_file
        if p.exists():
            md_files.append(p)
    docs_dir = root / "docs"
    if docs_dir.exists():
        md_files.extend(sorted(docs_dir.rglob("*.md")))
    agents_dir = root / ".agents"
    if agents_dir.exists():
        md_files.extend(sorted(agents_dir.rglob("*.md")))

    for md_path in md_files:
        text = md_path.read_text(encoding="utf-8", errors="ignore")
        rel_md = md_path.relative_to(root)
        for line_no, raw_target in _iter_markdown_links(text):
            target = raw_target.strip()
            # Ignore URLs, pure anchors, and placeholder/example paths
            if (
                not target
                or target.startswith("#")
                or re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", target)
                or target.startswith("mailto:")
                or any(tok in target for tok in ("<", ">", "[", "]", "{", "}", "..."))
            ):
                continue
            path_part = unquote(target.split("#", 1)[0].split("?", 1)[0])
            if not path_part:
                continue
            if path_part.startswith("/") or Path(path_part).is_absolute():
                errors.append(
                    f"{rel_md}:{line_no}: absolute filesystem path not allowed in markdown link '{raw_target}'"
                )
                continue
            resolved = (md_path.parent / path_part).resolve()
            try:
                resolved.relative_to(root_resolved)
            except ValueError:
                errors.append(
                    f"{rel_md}:{line_no}: markdown link escapes repository root '{raw_target}'"
                )
                continue
            if not resolved.exists():
                errors.append(f"{rel_md}:{line_no}: broken relative link '{raw_target}'")
    return errors


def validate_harness(root: Path = ROOT) -> tuple[list[str], list[str], dict[str, int]]:
    errors: list[str] = []
    warnings: list[str] = []
    agents_dir = root / ".agents"
    always_on_bytes = 0

    # 1. Root AGENTS.md
    agents_md = root / "AGENTS.md"
    if not agents_md.exists():
        errors.append("Missing root AGENTS.md")
    else:
        size = agents_md.stat().st_size
        always_on_bytes += size
        if size > SINGLE_RULE_MAX_BYTES:
            errors.append(
                f"AGENTS.md exceeds {SINGLE_RULE_MAX_BYTES:,}-byte Antigravity rule cap ({size} bytes)"
            )
        agents_text = agents_md.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\r", "\n")
        if agents_text.startswith("---\n"):
            errors.append("AGENTS.md should not contain YAML frontmatter")

    # 2. Rules (.agents/rules/*.md)
    rules_dir = agents_dir / "rules"
    rule_files = sorted(rules_dir.glob("*.md")) if rules_dir.exists() else []
    if not rule_files:
        errors.append("No rule files found in .agents/rules/")
    for rf in rule_files:
        rf_size = rf.stat().st_size
        if rf_size > SINGLE_RULE_MAX_BYTES:
            errors.append(
                f"{rf.relative_to(root)} exceeds {SINGLE_RULE_MAX_BYTES:,}-byte rule limit ({rf_size} bytes)"
            )
        fm, body, fm_errs = parse_frontmatter(rf, root=root)
        errors.extend(fm_errs)
        unknown_rule_keys = sorted(set(fm.keys()) - ALLOWED_RULE_KEYS)
        if unknown_rule_keys:
            errors.append(
                f"{rf.relative_to(root)}: unsupported rule frontmatter key(s): {', '.join(unknown_rule_keys)}"
            )
        raw_trigger = fm.get("trigger", "")
        if not isinstance(raw_trigger, str) or raw_trigger not in VALID_TRIGGERS:
            errors.append(
                f"{rf.relative_to(root)}: invalid trigger {raw_trigger!r} "
                f"(expected one of {sorted(VALID_TRIGGERS)})"
            )
            trigger = ""
        else:
            trigger = raw_trigger
        if trigger == "always_on":
            always_on_bytes += rf_size
        if trigger == "model_decision":
            desc = fm.get("description")
            if not isinstance(desc, str) or not desc.strip():
                errors.append(f"{rf.relative_to(root)}: model_decision rule requires 'description'")
        if trigger == "glob":
            globs = fm.get("globs") if "globs" in fm else fm.get("glob")
            if not globs:
                errors.append(f"{rf.relative_to(root)}: glob rule requires 'globs'")
            elif isinstance(globs, list):
                errors.append(
                    f"{rf.relative_to(root)}: 'globs' must be a comma-separated string, not a YAML array"
                )
            elif not isinstance(globs, str):
                errors.append(
                    f"{rf.relative_to(root)}: 'globs' must be a string, got {type(globs).__name__}"
                )
        if not body:
            errors.append(f"{rf.relative_to(root)}: empty rule body")

    if always_on_bytes > AGGREGATE_ALWAYS_ON_MAX_BYTES:
        errors.append(
            f"Aggregate always-on context (AGENTS.md + always_on rules) is {always_on_bytes:,} bytes, "
            f"exceeding {AGGREGATE_ALWAYS_ON_MAX_BYTES:,}-byte budget"
        )

    # 3. Skills (.agents/skills/<name>/SKILL.md)
    skills_dir = agents_dir / "skills"
    skill_dirs = sorted(p for p in skills_dir.iterdir() if p.is_dir()) if skills_dir.exists() else []
    if not skill_dirs:
        errors.append("No skill directories found in .agents/skills/")
    for sd in skill_dirs:
        skill_md = sd / "SKILL.md"
        if not skill_md.exists():
            errors.append(f"{sd.relative_to(root)}: missing SKILL.md")
            continue
        fm, body, fm_errs = parse_frontmatter(skill_md, root=root)
        errors.extend(fm_errs)
        raw_name = fm.get("name", "")
        raw_desc = fm.get("description", "")
        name = raw_name if isinstance(raw_name, str) else ""
        desc = raw_desc if isinstance(raw_desc, str) else ""
        if not name or not KEBAB_RE.match(name) or len(name) > 64:
            errors.append(
                f"{skill_md.relative_to(root)}: invalid skill name '{raw_name}' "
                "(must start with a lowercase letter and use kebab-case)"
            )
        if name and name != sd.name:
            errors.append(f"{skill_md.relative_to(root)}: frontmatter name '{name}' != folder '{sd.name}'")
        if not desc:
            errors.append(f"{skill_md.relative_to(root)}: missing description")
        elif len(desc) > 1024:
            errors.append(f"{skill_md.relative_to(root)}: description exceeds 1,024 chars ({len(desc)})")
        if not body:
            errors.append(f"{skill_md.relative_to(root)}: empty SKILL.md body")

    # 4. Custom Agents (.agents/agents/*.md)
    custom_agents_dir = agents_dir / "agents"
    agent_files = sorted(custom_agents_dir.glob("*.md")) if custom_agents_dir.exists() else []
    if not agent_files:
        errors.append("No custom agents found in .agents/agents/")
    for af in agent_files:
        fm, body, fm_errs = parse_frontmatter(af, root=root)
        errors.extend(fm_errs)
        unknown_keys = sorted(set(fm.keys()) - ALLOWED_AGENT_KEYS)
        if unknown_keys:
            errors.append(
                f"{af.relative_to(root)}: unsupported custom agent frontmatter key(s): {', '.join(unknown_keys)}"
            )
        raw_name = fm.get("name", "")
        name = raw_name if isinstance(raw_name, str) else ""
        if not name or not KEBAB_RE.match(name):
            errors.append(f"{af.relative_to(root)}: invalid agent name '{raw_name}'")
        if name and name != af.stem:
            errors.append(f"{af.relative_to(root)}: agent name '{name}' != filename stem '{af.stem}'")
        raw_desc = fm.get("description")
        if not isinstance(raw_desc, str) or not raw_desc.strip():
            errors.append(f"{af.relative_to(root)}: missing description")
        for bool_key in ("mainAgent", "subagent"):
            if bool_key in fm and not isinstance(fm[bool_key], bool):
                errors.append(
                    f"{af.relative_to(root)}: '{bool_key}' must be a boolean (true/false), got {fm[bool_key]!r}"
                )
        raw_model = fm.get("model", "inherit")
        if not isinstance(raw_model, str) or raw_model not in VALID_MODELS:
            errors.append(
                f"{af.relative_to(root)}: invalid model '{raw_model}' (expected one of {sorted(VALID_MODELS)})"
            )
        raw_policy = fm.get("commandExecutionPolicy", "sandbox")
        if not isinstance(raw_policy, str) or raw_policy not in VALID_CMD_POLICIES:
            errors.append(f"{af.relative_to(root)}: invalid commandExecutionPolicy '{raw_policy}'")
        if not body:
            errors.append(f"{af.relative_to(root)}: empty agent system prompt")

    # 5. Stale reference scan across .agents/
    if agents_dir.exists():
        for path in sorted(agents_dir.rglob("*")):
            if path.is_file() and path.suffix in (".md", ".txt", ".py"):
                content = path.read_text(encoding="utf-8", errors="ignore")
                for idx, line in enumerate(content.splitlines(), 1):
                    if STALE_PATTERNS.search(line):
                        errors.append(
                            f"{path.relative_to(root)}:{idx}: stale non-Antigravity reference: {line.strip()[:100]}"
                        )

    # 6. Relative markdown link validation
    errors.extend(validate_markdown_links(root))

    # 7. Validate MCP example JSON
    mcp_example = root / "config" / "mcp_config.example.json"
    if mcp_example.exists():
        try:
            parsed = json.loads(mcp_example.read_text(encoding="utf-8"))
            if (
                not isinstance(parsed, dict)
                or "mcpServers" not in parsed
                or not isinstance(parsed["mcpServers"], dict)
            ):
                errors.append(
                    "config/mcp_config.example.json must be a JSON object containing a 'mcpServers' object"
                )
            else:
                for srv_name, srv_cfg in parsed["mcpServers"].items():
                    if not isinstance(srv_cfg, dict):
                        errors.append(
                            f"config/mcp_config.example.json: server '{srv_name}' must be a JSON object"
                        )
                        continue
                    has_remote = any(
                        isinstance(srv_cfg.get(k), str) and srv_cfg.get(k).strip()
                        for k in ("serverUrl", "url")
                    )
                    has_local = isinstance(srv_cfg.get("command"), str) and bool(
                        srv_cfg.get("command").strip()
                    )
                    if has_remote and has_local:
                        errors.append(
                            f"config/mcp_config.example.json: server '{srv_name}' must not specify both "
                            "'command' and 'serverUrl'/'url'"
                        )
                    elif not has_remote and not has_local:
                        errors.append(
                            f"config/mcp_config.example.json: server '{srv_name}' must define 'serverUrl', 'url', or 'command'"
                        )
        except json.JSONDecodeError as exc:
            errors.append(f"config/mcp_config.example.json invalid JSON: {exc}")
    else:
        warnings.append("config/mcp_config.example.json not found")

    counts = {
        "rules": len(rule_files),
        "skills": len(skill_dirs),
        "agents": len(agent_files),
        "always_on_bytes": always_on_bytes,
    }
    return errors, warnings, counts


def main() -> int:
    errors, warnings, counts = validate_harness(ROOT)

    print(
        f"Validated Antigravity harness: {counts['rules']} rules, "
        f"{counts['skills']} skills, {counts['agents']} custom agents "
        f"(always-on context: {counts['always_on_bytes']:,} bytes)."
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
