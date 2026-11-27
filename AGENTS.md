# Equity Research AI Harness (Google Antigravity)

This repository is a Google Antigravity workspace harness for institutional-grade public equity research, financial modeling, sector & competitive analysis, and ongoing coverage management.

## Workspace Architecture

| Layer | Location | How Antigravity Uses It |
|---|---|---|
| **Root Guidance** | `AGENTS.md` | Always loaded at the start of every session |
| **Modular Rules** | `.agents/rules/*.md` | Domain guardrails loaded always (`research-integrity`), on demand (`financial-modeling`), or by path glob (`research-outputs`) |
| **Skills & Slash Commands** | `.agents/skills/<name>/SKILL.md` | Auto-discovered semantically by the agent AND directly invocable via `/<name>` |
| **Custom Agents** | `.agents/agents/<name>.md` | Selectable as primary agent (`/agents` or `agy --agent <name>`) or delegated via `invoke_subagent` |
| **Deliverables** | `research/` | Structured company coverage, sector primers, screens, catalyst calendar, and morning notes |
| **Utilities** | `scripts/` | Excel formula recalculation & error audit (`recalc_xlsx.py`), harness validation (`validate_harness.py`) |

## Task Routing — Skills & Slash Commands

Use the matching skill (or invoke directly with `/<skill-name>`):

### Single-Name Coverage & Earnings
- `/initiating-coverage` — 5-task institutional initiation report (company research → model → valuation → charts → assembly)
- `/earnings-preview` — Pre-earnings setup, consensus expectations, KPIs to watch, bull/base/bear scenarios
- `/earnings-analysis` — Post-earnings quarterly update report (beat/miss, call read, estimate changes, thesis impact)
- `/model-update` — Update an existing coverage model with new quarterly actuals, guidance, or assumptions
- `/thesis-tracker` — Maintain and stress-test a living investment thesis, key pillars, and disconfirming evidence
- `/catalyst-calendar` — Track upcoming earnings, investor days, product launches, regulatory milestones, and macro prints
- `/morning-note` — Concise daily morning meeting note covering overnight news and actionable calls

### Sector, Competitive & Idea Sourcing
- `/sector-overview` — Industry landscape primer (market sizing, value chain, drivers, themes, investability)
- `/competitive-analysis` — Competitive positioning, moat assessment, competitor deep-dives, and market maps
- `/idea-generation` — Quantitative stock screening and thematic long/short idea sourcing

### Valuation, Modeling & Unit Economics
- `/comps-analysis` — Comparable company analysis (operating benchmarks, trading multiples, quartile statistics)
- `/dcf-model` — Institutional DCF model (WACC build, FCF projections, terminal value cross-checked vs. comps, 5×5 sensitivities)
- `/3-statement-model` — Integrated 3-statement financial model (IS, BS, CF, working capital, debt schedule)
- `/unit-economics` — SaaS/subscription & recurring-revenue unit economics (ARR bridge, cohorts, NDR/GRR, LTV/CAC, Rule of 40)

### Quality Assurance & Deliverable Authoring
- `/audit-xls` — Audit any spreadsheet or financial model for formula errors, balance checks, and hardcoded overrides
- `/clean-data-xls` — Clean and normalize raw financial data exports before modeling
- `/xlsx-author` — Build standalone `.xlsx` workbooks programmatically with `openpyxl` and formula verification
- `/pptx-author` — Build `.pptx` research and pitch decks programmatically with `python-pptx` and visual QA

## Custom Agents (Multi-Step Orchestration)

For multi-skill workflows, select a primary agent (`/agents` or `agy --agent <name>`) or delegate in parallel via `invoke_subagent`:
- `equity-analyst` — Lead analyst for full-lifecycle single-name coverage and subagent coordination
- `earnings-reviewer` — End-to-end post-earnings update (`earnings-analysis` → `model-update` → `audit-xls` → `morning-note`)
- `market-researcher` — Sector & thematic primer (`sector-overview` → `competitive-analysis` → `comps-analysis` → `idea-generation`)
- `model-builder` — Scratch financial model build (`3-statement-model` / `dcf-model` / `comps-analysis` / `unit-economics` → `audit-xls`)

## Data Source Hierarchy

Always follow this priority order when gathering financial data:
1. **Connected financial MCP servers** (if configured in `.agents/mcp_config.json`, e.g., FactSet, S&P Global/Kensho, Daloopa, LSEG)
2. **Primary public filings** — SEC EDGAR (`10-K`, `10-Q`, `8-K`, `DEF 14A`), official earnings releases, investor presentations, and call transcripts
3. **Workspace files & connected Google Drive / Sheets** — Existing models, notes, or user-uploaded datasets
4. **Web search** — For latest share prices, news, and consensus context when primary feeds are unavailable; cross-verify against filings and cite source + date

Never invent or silently estimate missing historical financials. Mark unverified figures `[UNSOURCED]` and explicit analyst estimates `[ASSUMPTION]`.

## Default Output Paths

Save persistent research outputs inside `research/`:
- Single-company work: `research/companies/<TICKER>/` (`thesis.md`, `models/`, `notes/`, `reports/`, `data/`)
- Sector & competitive work: `research/sectors/<sector-slug>/` (`overview.md`, `competitive-landscape.md`, `comps/`, `decks/`)
- Stock screens & ideas: `research/screens/`
- Daily morning notes: `research/morning-notes/YYYY-MM-DD.md`
- Coverage catalyst calendar: `research/catalysts.md`

See `docs/data-sources.md` for configuring SEC EDGAR access, Python libraries, and optional institutional MCP connectors.
