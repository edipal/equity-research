---
name: initiating-coverage
description: Create institutional-grade equity research initiation report drafts through a 5-task workflow executed one task at a time with verified prerequisites — (1) company research, (2) financial modeling, (3) valuation analysis, (4) chart generation, (5) final DOCX assembly. Use for comprehensive first-time coverage of a public company. Not for quarterly earnings updates (use earnings-analysis) or standalone models (use dcf-model or three-statement-model).
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/initiating-coverage/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# Initiating Coverage

Produce comprehensive, first-time equity research coverage drafts through a structured **5-task workflow**. Every deliverable is staged as **`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`** under `research/companies/<TICKER>/` per `.agents/rules/research-integrity.md`.

**Default Font:** Calibri across all documents and charts (unless the user specifies otherwise).

---

## Core Guardrails (Read Before Starting)

1. **One Task at a Time:** Execute exactly **one** task per user request (never chain Tasks 1–5 end-to-end automatically). If the user asks for an entire initiation report or full pipeline, list the 5 tasks and ask which task to start with (default recommendation: Task 1 or Task 2).
2. **Drafts Only & No Fabricated Identities:** Label every report, model summary, and deliverable `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`. Use `[ANALYST NAME — REVIEW REQUIRED]` and `[DATE]` placeholders — **never** invent analyst names, CFA/Ph.D. credentials, email addresses, phone numbers, or firm attributions.
3. **No Fabrication to Hit Quotas:** All row counts, chart counts, word counts, and peer/deal counts are **target ranges where disclosed in primary filings** — never fabricate segment breakdowns, sub-geographies, funding rounds, customer cohorts, or precedent deals to satisfy a quota. Tag every analyst estimate `[ASSUMPTION]` and any unverified figure `[UNSOURCED]`.
4. **Canonical Deliverables Only:** Create only the specified deliverables for the requested task — no unsolicited completion summaries or meta-reports.
5. **Load Only the Active Task Reference:** Each task's full instructions live in `references/task1`–`task5`. Load **only** the reference file for the single task being executed to conserve context.

---

## 5-Task Pipeline & Canonical Deliverables

| Task | Name | Prerequisites | Reference Doc | Canonical Deliverable(s) |
|---|---|---|---|---|
| **1** | Company Research | Ticker / company name (independent) | [references/task1-company-research.md](references/task1-company-research.md) | `research/companies/<TICKER>/reports/<TICKER>_Research_Document_<Date>.md` (target ~6,000–8,000 words) |
| **2** | Financial Modeling | SEC 10-K/10-Q filings or user historicals (Task 1 optional) | [references/task2-financial-modeling.md](references/task2-financial-modeling.md) | `research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx` (5 core tabs: `Revenue Model`, `Income Statement`, `Cash Flow Statement`, `Balance Sheet`, `Scenarios`) + optional `<TICKER>_Historical_Financials_<Date>.xlsx` |
| **3** | Valuation Analysis | Task 2 financial model (verified via `recalc_xlsx.py`) | [references/task3-valuation.md](references/task3-valuation.md) & [references/valuation-methodologies.md](references/valuation-methodologies.md) | `research/companies/<TICKER>/reports/<TICKER>_Valuation_Analysis_<Date>.md` + 4 valuation tabs added to Task 2's `.xlsx`: `WACC`, `DCF` (inline 5×5 sensitivity or optional `Sensitivity` tab), `Comparable Companies`, `Valuation Summary` |
| **4** | Chart Generation | Tasks 1, 2, & 3 + verified market data | [references/task4-chart-generation.md](references/task4-chart-generation.md) | 300 DPI PNGs (target 20–30+ charts where disclosed, prioritizing core financial/valuation charts) + `chart_index.txt` in `research/companies/<TICKER>/reports/charts/` |
| **5** | Report Assembly | Tasks 1, 2, 3, & 4 | [references/task5-report-assembly.md](references/task5-report-assembly.md), [assets/report-template.md](assets/report-template.md), [assets/quality-checklist.md](assets/quality-checklist.md) | `research/companies/<TICKER>/reports/<TICKER>_Initiation_Report_<Date>.docx` (target ~30–50 pages, dense layout, Disclosures & Important Disclaimers appendix) |

*(Tasks 1 and 2 are independent and may be run in either order. Tasks 3, 4, and 5 have strict upstream dependencies — if any prerequisite is missing, STOP and inform the user which task must be completed first.)*

---

## Response Protocol for Full-Pipeline Requests

When a user says *"Create a coverage initiation report for [Company]"* or *"Run all 5 tasks for [Company]"*:
- Do **not** auto-start or chain tasks.
- Present the 5 tasks concisely and ask which task to execute first (suggesting Task 1: Company Research or Task 2: Financial Modeling).

---

## Task Summaries & QA Gates

### Task 1: Company Research
- **Scope:** Company overview, founding/history, executive bios (3–4 executives from DEF 14A/filings), products & pricing, customers & GTM, industry overview, competitive landscape (target 5–10 peers), TAM sizing, and risk assessment (target 8–12 risks across 4 categories) with dated citations.
- **Guardrail:** Document only facts verifiable from SEC filings, company IR, or cited secondary reports; mark estimates `[ASSUMPTION]` and gaps `[UNSOURCED]`.
- **Full workflow:** [references/task1-company-research.md](references/task1-company-research.md).

### Task 2: Financial Modeling
- **Scope:** Extract 3–5 years of historicals (+ latest 10-Q LTM stub if mid-year) and build a live-formula 5-year forecast workbook with **5 canonical tabs**:
  1. `Revenue Model` (product/segment and geographic breakdowns mirroring primary 10-K/10-Q disclosures)
  2. `Income Statement` (full P&L matching filing granularity)
  3. `Cash Flow Statement` (CFO, CFI, CFF, FCF)
  4. `Balance Sheet` (full BS with balance check `= 0` every period)
  5. `Scenarios` (Bear / Base / Bull drivers & outputs)
- **QA Gate (MANDATORY):** Run `python3 scripts/recalc_xlsx.py research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx 30` and confirm exit `0` (`"status": "success"`, zero errors, zero uncached formulas, zero balance-check failures).
- **Full workflow:** [references/task2-financial-modeling.md](references/task2-financial-modeling.md).

### Task 3: Valuation Analysis
- **Scope:** Delegate DCF mechanics to `dcf-model` (gross-debt WACC, mid-year explicit UFCF (`t - 0.5`) and mid-year perpetuity TV (`(1 + WACC)^(N - 0.5)`), strict end-of-year `(1 + WACC)^N` exit-multiple TV discounting, complete equity bridge, TSM diluted shares, SBC as real expense, centered 5×5 live sensitivity grid, TV `%` of EV target `50%–70%`). Add **4 canonical tabs** to Task 2's existing `.xlsx`:
  1. `WACC`
  2. `DCF` *(with inline 5×5 sensitivity table per `dcf-model`, or optional dedicated `Sensitivity` tab)*
  3. `Comparable Companies` *(5–10 peers + Max / 75th / Median / 25th / Min summary)*
  4. `Valuation Summary` *(reconciliation, weighting, football field inputs; Precedent Transactions tab optional and used as cross-check only unless M&A is an active catalyst)*
- **Rating Scale (Standardized):** `Buy` (`>= +15%` 12M expected return), `Hold` (`-10%` to `+15%`), `Sell` (`<= -10%`).
- **QA Gate (MANDATORY):** Run BOTH:
  - `python3 scripts/recalc_xlsx.py research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx 30` (exit `0`, `"status": "success"`)
  - `python3 .agents/skills/dcf-model/scripts/validate_dcf.py research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx` (exit `0`, `"status": "PASS"`)
- **Full workflow:** [references/task3-valuation.md](references/task3-valuation.md) & [references/valuation-methodologies.md](references/valuation-methodologies.md).

### Task 4: Chart Generation
- **Scope:** Generate 300 DPI PNG charts via Python (`matplotlib`, `seaborn`, `Pillow`) saved directly into `research/companies/<TICKER>/reports/charts/` alongside `chart_index.txt`.
- **Core Priority Charts (always build where disclosed):**
  - `chart_03` — Revenue by Product/Segment (stacked area; where multiple segments are disclosed)
  - `chart_04` — Revenue by Geography (stacked bar; where geographic breakdown is disclosed)
  - `chart_28` — DCF Sensitivity 2-Way Heatmap (formatted with floats `fmt='.1f'` / `fmt='.2f'`, never `fmt='d'`)
  - `chart_32` — Valuation Football Field (horizontal bar chart with current & target price markers)
- **No-Fabrication Rule:** Target 20–30+ charts across Tasks 1–3 and verified market data. If a chart's underlying data is not disclosed in primary filings (e.g., undisclosed cohorts or private funding), omit it and document the reason in `chart_index.txt` rather than inventing data.
- **Full workflow:** [references/task4-chart-generation.md](references/task4-chart-generation.md).

### Task 5: Report Assembly
- **Scope:** Assemble `research/companies/<TICKER>/reports/<TICKER>_Initiation_Report_<Date>.docx` programmatically via `python-docx` + `openpyxl` (`data_only=True` after `recalc_xlsx.py`):
  - Page 1: **INITIATING COVERAGE** header (with prominent `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY` banner), rating box, `[ANALYST NAME — REVIEW REQUIRED]`, Figure 1 stock chart, 3–4 quantified `■` bullets, financial summary table.
  - Reuse Task 1 prose verbatim for Company 101 sections, embed Task 4 PNG charts at `<= 6.0"` width (matching default `python-docx` text width), extract Task 2/3 tables directly from Excel, and write deep Projection Assumptions, Scenario Analysis, and Valuation sections.
  - Mandatory appendices: **Data Sources & References** (clickable hyperlinks with dates) and **Disclosures & Important Disclaimers** (draft status, not investment advice, analyst review sign-off).
  - Verify rendered layout and page count via LibreOffice PDF export + `pypdf`.
- **Full workflow:** [references/task5-report-assembly.md](references/task5-report-assembly.md), [assets/report-template.md](assets/report-template.md), [assets/quality-checklist.md](assets/quality-checklist.md).

---

## Canonical Workspace Layout

```text
research/companies/<TICKER>/
├── models/
│   ├── <TICKER>_Historical_Financials_<Date>.xlsx   # Task 2 (optional Step 1 extraction)
│   └── <TICKER>_Financial_Model_<Date>.xlsx         # Task 2 (5 tabs) + Task 3 (4 valuation tabs)
└── reports/
    ├── <TICKER>_Research_Document_<Date>.md         # Task 1
    ├── <TICKER>_Valuation_Analysis_<Date>.md        # Task 3
    ├── charts/                                      # Task 4 (300 DPI PNGs + chart_index.txt)
    └── <TICKER>_Initiation_Report_<Date>.docx       # Task 5 (DRAFT for analyst review)
```
