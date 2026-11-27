# Equity Research AI Harness for Google Antigravity

An institutional-grade **Google Antigravity** AI harness for public equity research, financial modeling, sector & competitive analysis, and ongoing coverage management — extracted and adapted from [`anthropics/financial-services`](https://github.com/anthropics/financial-services) (Apache-2.0) into native Antigravity **Rules**, **Skills** (auto-discovered + `/<skill>` slash commands), and **Custom Agents**.

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
├── AGENTS.md                        # Always-on workspace operating guide
├── .agents/
│   ├── rules/                       # Modular Antigravity rules (trigger-scoped)
│   │   ├── research-integrity.md    # always_on: anti-prompt-injection, citations, timeliness
│   │   ├── financial-modeling.md    # model_decision: Excel formulas, color coding, QA
│   │   └── research-outputs.md      # glob (research/**/*): directory & formatting standards
│   ├── agents/                      # 4 custom agents (primary or background subagents)
│   │   ├── equity-analyst.md        # Lead coverage analyst & subagent coordinator
│   │   ├── earnings-reviewer.md     # End-to-end post-earnings print & model reviewer
│   │   ├── market-researcher.md     # Sector primer, competitive landscape, comps & ideas
│   │   └── model-builder.md         # Scratch 3-statement, DCF, comps & unit economics builder
│   └── skills/                      # 18 skills (auto-discovered + /<name> slash commands)
├── config/
│   └── mcp_config.example.json      # Opt-in institutional MCP connector template
├── docs/
│   └── data-sources.md              # Public filings (SEC EDGAR) & optional MCP setup
├── scripts/
│   ├── recalc_xlsx.py               # LibreOffice Excel formula recalculation & error scan
│   └── validate_harness.py          # Automated Antigravity harness schema validator
└── research/                        # Structured workspace for your coverage deliverables
```

---

## Included Skills & Slash Commands (18)

Every folder in `.agents/skills/` is both **semantically discovered** by Antigravity and directly callable via `/<skill-name>`:

### Single-Name Coverage & Earnings (9)
| Slash Command | Purpose |
|---|---|
| `/initiating-coverage` | Complete 5-task institutional initiation report (company research, model, valuation, 25–35 charts, report assembly) with templates and quality checklists |
| `/earnings-analysis` | Fast-turnaround post-earnings update (timeliness check, beat/miss variance, call transcript read, revised estimates, thesis impact) |
| `/earnings-preview` | Pre-earnings setup with consensus expectations, KPIs to watch, and bull/base/bear stock reaction scenarios |
| `/model-update` | Roll an existing coverage model forward with new quarterly actuals, guidance, and assumption changes |
| `/thesis-tracker` | Maintain a living investment thesis with pillars, bull/base/bear cases, and disconfirming evidence checks |
| `/catalyst-calendar` | Track upcoming earnings dates, investor days, product launches, regulatory milestones, and macro prints |
| `/morning-note` | Concise 7am-style morning meeting note summarizing overnight developments and actionable ideas |
| `/sector-overview` | Comprehensive industry primer (market sizing, value chain, structural drivers, and thematic trends) |
| `/idea-generation` | Systematic quantitative stock screens and thematic sweeps for long and short ideas |

### Competitive Analysis, Valuation & Modeling (6)
| Slash Command | Purpose |
|---|---|
| `/competitive-analysis` | Competitive landscape decks, positioning matrices, competitor deep-dives, and moat synthesis |
| `/comps-analysis` | Comparable company trading multiples (`EV/Rev`, `EV/EBITDA`, `P/E`) with Max/75th/Median/25th/Min statistics |
| `/dcf-model` | Institutional DCF model (WACC build, Bear/Base/Bull FCF, comps-checked terminal value, 5×5 sensitivities) + `validate_dcf.py` |
| `/three-statement-model` | Integrated Income Statement, Balance Sheet, and Cash Flow model with SEC filing references |
| `/unit-economics` | SaaS/subscription & recurring-revenue unit economics (ARR bridge, cohort retention, NDR/GRR, LTV/CAC, Rule of 40) |
| `/audit-xls` | Comprehensive formula and accounting integrity audit (balance checks, cash tie-out, hardcode detection) |

### Data Hygiene & Deliverable Authoring (3)
| Slash Command | Purpose |
|---|---|
| `/clean-data-xls` | Clean and normalize raw spreadsheet exports (text-stored numbers, mixed dates, duplicates, errors) |
| `/xlsx-author` | Programmatic institutional `.xlsx` creation via `openpyxl` + automated formula recalculation & audit |
| `/pptx-author` | Programmatic `.pptx` research & pitch deck creation via `python-pptx` + visual PDF/PNG verification |

---

## Custom Agents (4)

Select any agent as your primary session agent (`/agents` in CLI, `agy --agent <name>`, or in the Antigravity App) or let the main agent delegate tracks in parallel:

- **`equity-analyst`** — Lead equity analyst for initiating coverage, thesis tracking, catalyst monitoring, and coordinating subagents.
- **`earnings-reviewer`** — End-to-end earnings cycle processor (`earnings-analysis` → `model-update` → `audit-xls` → `morning-note` → `thesis-tracker`).
- **`market-researcher`** — Sector & thematic researcher (`sector-overview` → `competitive-analysis` → `comps-analysis` → `idea-generation` → optional `pptx-author`).
- **`model-builder`** — Institutional modeler (`three-statement-model`, `dcf-model`, `comps-analysis`, `unit-economics`, `xlsx-author` → `audit-xls`).

---

## Validation & Maintenance

Run the built-in validator after editing any rule, skill, or custom agent:

```bash
python3 scripts/validate_harness.py
```

Recalculate and audit any generated Excel model for formula errors:

```bash
python3 scripts/recalc_xlsx.py research/companies/MSFT/models/MSFT_DCF.xlsx
```

## License & Attribution

Licensed under the **Apache License 2.0** (see [`LICENSE`](LICENSE) and [`NOTICE`](NOTICE)). Adapted from [`anthropics/financial-services`](https://github.com/anthropics/financial-services) for Google Antigravity.