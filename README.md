# Equity Research AI Harness for Google Antigravity

An institutional-grade **Google Antigravity** AI harness for public equity research, financial modeling, sector & competitive analysis, and ongoing coverage management — extracted and adapted from [`anthropics/financial-services`](https://github.com/anthropics/financial-services) (`@3e32c22`, Apache-2.0) into native Antigravity **Rules**, **Skills** (auto-discovered + `/<skill>` slash commands), and **Custom Agents**.

## Quickstart

1. **Open this repository in Google Antigravity** (App, CLI `agy`, or IDE extension). Antigravity automatically discovers `AGENTS.md`, `.agents/rules/`, `.agents/skills/`, and `.agents/agents/` — no manual installation required.
2. *(Optional)* **Install Python dependencies** for Excel modeling, charts, and Word/PowerPoint assembly:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```
3. *(Optional)* **Configure institutional MCP connectors** (FactSet, S&P Global, Daloopa, LSEG, etc.) — see [`docs/data-sources.md`](docs/data-sources.md). The harness works out of the box with free public sources (SEC EDGAR, earnings releases, web research, and local/Drive files).
4. **Start researching** by typing a natural prompt or invoking any skill directly as a slash command (e.g., `/sector-overview AI networking`, `/competitive-analysis Cloud Observability`, `/dcf-model MSFT`, `/earnings-analysis NVDA Q3 FY26`), or switch to a specialist agent via `/agents` (or `agy --agent equity-analyst`).

---

## Harness Architecture

```text
equity-research/
├── AGENTS.md                        # Always-on workspace operating guide & canonical file naming
├── .agents/
│   ├── rules/                       # Modular Antigravity rules (trigger-scoped)
│   │   ├── research-integrity.md    # always_on: anti-prompt-injection, citations, drafts-only
│   │   ├── financial-modeling.md    # model_decision: Excel formulas, palette, QA gates
│   │   └── research-outputs.md      # glob (research/**/*): directory & formatting standards
│   ├── agents/                      # 4 custom agents (primary or background subagents)
│   │   ├── equity-analyst.md        # Lead coverage analyst & subagent coordinator (original)
│   │   ├── earnings-reviewer.md     # End-to-end post-earnings print & model reviewer
│   │   ├── market-researcher.md     # Sector primer, competitive landscape, comps & ideas
│   │   └── model-builder.md         # Scratch 3-statement, DCF, comps & unit economics builder
│   └── skills/                      # 18 skills (auto-discovered + /<name> slash commands)
├── config/
│   └── mcp_config.example.json      # Opt-in institutional MCP connector template (illustrative)
├── docs/
│   └── data-sources.md              # Public filings (SEC EDGAR) & optional MCP setup
├── scripts/
│   ├── recalc_xlsx.py               # LibreOffice Excel formula recalculation & error scan
│   └── validate_harness.py          # Automated Antigravity harness schema validator
├── tests/
│   └── test_scripts.py              # Automated regression test suite
└── research/                        # Structured workspace for your coverage deliverables
```

---

## Included Skills & Slash Commands (18)

Every folder in `.agents/skills/` is both **semantically discovered** by Antigravity and directly callable via `/<skill-name>` (grouped identically to [`AGENTS.md`](AGENTS.md)):

### Single-Name Coverage & Earnings (7)
| Slash Command | Purpose |
|---|---|
| `/initiating-coverage` | 5-task institutional initiation draft workflow (company research, 5-tab model, 4-tab valuation, charts where disclosed, DOCX report assembly) with templates and quality checklists |
| `/earnings-preview` | Pre-earnings setup with verified consensus benchmarks, KPIs to watch, and bull/base/bear stock reaction scenarios |
| `/earnings-analysis` | Fast-turnaround post-earnings update draft (timeliness check, strict pre-print consensus variance, call read, revised estimates, thesis impact) |
| `/model-update` | Roll an existing coverage model forward with new quarterly actuals, guidance, and assumption changes |
| `/thesis-tracker` | Maintain a living investment thesis with pillars, bull/base/bear cases, and disconfirming evidence checks |
| `/catalyst-calendar` | Track upcoming earnings dates (`Confirmed` / `Projected [ASSUMPTION]`), investor days, launches, regulatory milestones, and macro prints |
| `/morning-note` | Concise 7am-style morning meeting draft summarizing overnight developments and actionable ideas |

### Sector, Competitive & Idea Sourcing (3)
| Slash Command | Purpose |
|---|---|
| `/sector-overview` | Comprehensive industry primer (market sizing, value chain, structural drivers, and thematic trends) |
| `/competitive-analysis` | Competitive landscape reports & decks, positioning matrices, competitor deep-dives, and moat synthesis |
| `/idea-generation` | Quantitative stock screens (SEC EDGAR XBRL, `yfinance`, Form 4, FINRA short interest) and thematic sweeps (`5–10` candidates) |

### Valuation, Modeling & Unit Economics (4)
| Slash Command | Purpose |
|---|---|
| `/comps-analysis` | Comparable company trading multiples (full EV, LTM/NTM calendarized `EV/Rev`, `EV/EBITDA`, `P/E`, `"NM"` guards) with Max/75th/Median/25th/Min statistics |
| `/dcf-model` | Institutional 2-sheet DCF model (gross-debt WACC, Bear/Base/Bull `INDEX` blocks, mid-year FCF, end-of-year TV, 5×5 live sensitivities) + `build_dcf_skeleton.py` & `validate_dcf.py` |
| `/three-statement-model` | Scratch-built or template-populated integrated Income Statement, Balance Sheet, and Cash Flow model (ASC 740 NOL/DTA, revolver sweep, equity roll-forwards) |
| `/unit-economics` | Public-equity & SaaS recurring-revenue unit economics (ARR/billings bridge, cohorts/proxies, lagged CAC, NDR/GRR, LTV/CAC, Rule of 40) |

### Quality Assurance & Deliverable Authoring (4)
| Slash Command | Purpose |
|---|---|
| `/audit-xls` | Read-only spreadsheet & model audit (`recalc_xlsx.py --check-only`, balance checks, cash tie-out, hardcode detection) |
| `/clean-data-xls` | Clean and normalize raw spreadsheet exports (text-stored numbers, mixed dates, duplicates, errors) |
| `/xlsx-author` | Programmatic institutional `.xlsx` creation via `openpyxl` + automated LibreOffice formula verification |
| `/pptx-author` | Programmatic `.pptx` research & pitch deck creation via `python-pptx` + visual PDF/PNG verification |

### Upstream Slash-Command Mapping (`anthropics/financial-services`)
If you are migrating from Anthropic's upstream Claude plugin commands, use the Antigravity skill equivalents:

| Upstream Command | Antigravity Slash Command |
|---|---|
| `/earnings` | `/earnings-analysis` |
| `/dcf` | `/dcf-model` |
| `/comps` | `/comps-analysis` |
| `/initiate` | `/initiating-coverage` |
| `/3-statement-model` | `/three-statement-model` |
| `/debug-model` | `/audit-xls` |
| `/screen` | `/idea-generation` |
| `/sector` | `/sector-overview` |
| `/thesis` | `/thesis-tracker` |
| `/catalysts` | `/catalyst-calendar` |

---

## Custom Agents (4)

Select any agent as your primary session agent (`/agents` in CLI, `agy --agent <name>`, or in the Antigravity App) or let the main agent delegate tracks in parallel:

- **`equity-analyst`** — Lead equity analyst for initiating coverage, thesis tracking, catalyst monitoring, and coordinating subagents.
- **`earnings-reviewer`** — End-to-end earnings cycle processor (variance read → `model-update` → `audit-xls` → `earnings-analysis` / `morning-note` → `thesis-tracker`).
- **`market-researcher`** — Sector & thematic researcher (`sector-overview` → `competitive-analysis` → `comps-analysis` → `idea-generation` shortlisting `3–5` high-conviction ideas from `5–10` candidates → optional `pptx-author`).
- **`model-builder`** — Institutional modeler (`three-statement-model`, `dcf-model`, `comps-analysis`, `unit-economics`, `xlsx-author` → `audit-xls`, with human review checkpoint).

---

## Validation & Maintenance

Run the built-in validator after editing any rule, skill, or custom agent:

```bash
python3 scripts/validate_harness.py
```

Recalculate and audit any generated Excel model for formula errors (use `--check-only` for read-only audits or `--output <PATH>` for non-destructive writes):

```bash
python3 scripts/recalc_xlsx.py research/companies/MSFT/models/MSFT_DCF_2026-03-15.xlsx 30
python3 .agents/skills/dcf-model/scripts/validate_dcf.py research/companies/MSFT/models/MSFT_DCF_2026-03-15.xlsx
```

Run the automated test suite:

```bash
python3 -m unittest discover -s tests -v
```

## License & Attribution

Licensed under the **Apache License 2.0** (see [`LICENSE`](LICENSE) and [`NOTICE`](NOTICE)). Adapted from [`anthropics/financial-services`](https://github.com/anthropics/financial-services) (`@3e32c22`) for Google Antigravity.