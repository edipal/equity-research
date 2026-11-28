---
name: earnings-analysis
description: Create professional equity research earnings update reports (8-12 pages, 3,000-5,000 words) analyzing quarterly results for both covered companies and standalone first-time earnings requests. Fast-turnaround format focusing on beat/miss analysis, key metrics, updated estimates/outlook, and thesis impact. Includes 1-3 summary tables and 8-12 charts. Use when user requests "earnings update", "quarterly update", "earnings analysis", "Q1/Q2/Q3/Q4 results", or post-earnings report.
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/earnings-analysis/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# Equity Research Earnings Update

Create institutional **EARNINGS UPDATE DRAFTS** analyzing quarterly results for companies under ongoing coverage **or** standalone first-time earnings requests. Stage all deliverables as **`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`**.

**Key Characteristics:**
- **Length**: 8-12 pages (target range; scale to materiality)
- **Word Count**: 3,000-5,000 words
- **Tables**: 1-3 summary tables (NOT comprehensive)
- **Figures**: 8-12 charts (where disclosed)
- **Turnaround**: Draft within 24-48 hours of earnings release
- **Audience**: Analysts / PMs (concise focus on what's new; add a brief business snapshot when running in standalone first-time mode)
- **Focus**: What's NEW - beat/miss, updated estimates/outlook, thesis impact
- **Font**: Times New Roman or Calibri throughout (consistent per user preference)

## When to Use

Use when the user requests:
- "Create an earnings update for [Company] Q3 2024"
- "Analyze [Company]'s quarterly results"
- "Post-earnings report for [Company]"
- "Q1/Q2/Q3/Q4 update for [Company]"

**Do NOT use if:**
- User requests "initiation report" → Use `/initiating-coverage`
- User requests "flash note" or "quick take" → Use `/morning-note`

## Coverage Modes: Ongoing Coverage vs. Standalone First-Time Earnings

- **Existing coverage (`research/companies/<TICKER>/models/` and/or `thesis.md` present):** Benchmark actuals against verified pre-print consensus and prior internal model estimates, update the coverage model via `model-update`, and revise the living thesis and forward estimates.
- **No prior coverage yet (standalone earnings request, e.g., *"NVDA Q3 earnings analysis"* in a fresh workspace):** Do **not** refuse or redirect into the 5-task `/initiating-coverage` workflow. Build a lightweight prior-quarter/pre-print baseline from the latest `10-K`, prior `10-Q`s, prior management guidance, and verified pre-print consensus (per the consensus source hierarchy below), tag the header and executive summary prominently **`[NO PRIOR COVERAGE MODEL — STANDALONE EARNINGS UPDATE]`**, deliver the full beat/miss analysis, call transcript takeaways, and forward outlook, and optionally offer `/initiating-coverage` afterward if the user wants full ongoing coverage.

## Critical Requirements

### 1. Speed & Timeliness Verification
- **Latest-quarter requests:** Search for `[Company] latest earnings results [current year]`, confirm the earnings release is the most recent quarter (flag if older than 3 months), and verify the earnings call transcript date matches the release date.
- **Historical-quarter requests:** If the user explicitly requests a specific historical quarter (e.g., *"analyze Q2 FY2024"*), do **not** reject it for staleness — verify the requested quarter's primary filings and note clearly that it is a historical retrospective review.
- Draft within 24-48 hours of earnings release (for live coverage)
- Focus on NEW information only; do not rehash company background extensively

### 2. Beat/Miss Analysis & Consensus Source Hierarchy (MANDATORY)
- **Pre-Print Consensus Source Hierarchy:**
  1. Connected financial MCP server (FactSet, S&P Capital IQ, LSEG, Bloomberg)
  2. Pre-release wire articles (Reuters, Bloomberg, WSJ) quoting a named consensus provider dated **BEFORE** the earnings release
  3. Company IR-compiled consensus (where published on IR site)
- Always record **provider, as-of date, and estimate count (`n`)** where available.
- **Prohibited:** NEVER use post-print Yahoo Finance consensus as pre-print consensus (post-print pages overwrite pre-release numbers), and do NOT cite unverified "whisper numbers".
- **Fallback when pre-print consensus is unavailable:** Benchmark actuals against **prior management guidance** or **prior internal model estimates**, and explicitly mark Street consensus **`[UNSOURCED]`**.
- **Basis matching:** Strictly match **GAAP vs. Non-GAAP** between reported actuals and consensus (never compare Non-GAAP EPS actual to GAAP consensus or vice versa).
- **In-line threshold:** Revenue within **`±1.0%`** and EPS within **`±2.0%` or `±$0.01`** = **`INLINE`**; above = **`BEAT`**; below = **`MISS`**.
- Explain **WHY** results differed from expectations (volume, pricing, mix, FX, or one-offs).

### 3. Summary Format
- Keep tables to 1-3 (summary only, not comprehensive)
- No full P&L/Cash Flow/Balance Sheet (just key metrics)
- Keep company background concise (if running in ongoing coverage, assume reader knows the business; if running in standalone first-time mode, include a crisp 2–3 sentence business snapshot and tag `[NO PRIOR COVERAGE MODEL — STANDALONE EARNINGS UPDATE]`)

### 4. Citations & Source Attribution ⭐⭐⭐ MANDATORY

**CRITICAL**: Properly cite all data with SPECIFIC sources and CLICKABLE HYPERLINKS.

**Include specific citations WITH CLICKABLE LINKS in every figure and table:**

```text
Source: Q3 FY2026 Form 10-Q filed 2026-11-06; Q3 FY2026 Earnings Release (2026-10-29)
        [Hyperlink "10-Q" to: https://www.sec.gov/Archives/edgar/data/<CIK>/<accession>/<doc>.htm]
        [Hyperlink "EDGAR company page" to: https://www.sec.gov/edgar/browse/?CIK=<CIK>]
        [Hyperlink "earnings release" to: https://investor.company.com/news/...]
```

**REQUIRED SOURCES LIST:**
- ✅ Earnings release (with date and URL)
- ✅ 10-Q or 10-K filing (with filing date and SEC EDGAR URL: `https://www.sec.gov/edgar/browse/?CIK=...` or `https://www.sec.gov/Archives/edgar/data/...`) — note SEC Form 10-Q is due within **40 days** of quarter end for large accelerated/accelerated filers (**45 days** for non-accelerated filers), though often filed alongside or within days of the release
- ✅ Earnings call transcript (with date matching release date)
- ✅ Investor presentation/supplemental materials (if available)
- ✅ Pre-print consensus estimates source (provider, as-of date, `n` — or `[UNSOURCED]` with guidance/model fallback)
- ✅ Prior guidance (from previous quarter's materials)

### 5. Updated Estimates
- Update forward estimates based on results (defer to `model-update` as single source of truth when running inside `earnings-reviewer`)
- Show old vs. new estimates clearly and tag forward projections `[ASSUMPTION]`
- Explain what changed and why

## High-Level Workflow

The earnings update process follows 5 phases:

### Phase 1: Data Collection (30-60 minutes)

**Verify Timeliness Before Analyzing:**
1. **Write down today's date** and confirm whether the request targets the **latest quarter** or a **specific historical quarter**.
2. **Search primary sources** (MCP / SEC EDGAR / Company IR) — never rely on stale training data.
3. **Verify release & transcript dates match** (for "latest quarter" requests, alert if the most recent release found is >3 months old; for explicitly requested historical quarters, proceed and label as historical).
4. **Pull pre-print consensus** using the strict hierarchy above (record provider, as-of date, `n`, and GAAP/Non-GAAP basis).

**See [references/workflow.md](references/workflow.md)** for detailed search procedures and verification steps.

### Phase 2: Analysis (2-3 hours)
- Beat/miss analysis for each key metric
- Segment/geographic/product breakdown
- Margin and guidance analysis
- Update financial model and estimates

**See [references/workflow.md](references/workflow.md)** for detailed analysis framework.

### Phase 3: Chart Generation (1-2 hours)
Create 8-12 charts focusing on quarterly trends and what's new:
- Quarterly revenue progression
- Quarterly EPS progression
- Quarterly margin trends
- Revenue by segment/geography
- Key operating metrics
- Beat/miss summary
- Estimate revisions
- Valuation charts

**See [references/workflow.md](references/workflow.md)** for chart specifications.

### Phase 4: Report Creation (2-3 hours)
Create 8-12 page DOCX report with specific structure.

**See [references/report-structure.md](references/report-structure.md)** for complete page-by-page templates and formatting requirements.

**High-level structure:**
- Page 1: Earnings summary with rating and price target
- Pages 2-3: Detailed results analysis
- Pages 4-5: Key metrics & guidance
- Pages 6-7: Updated investment thesis
- Pages 8-10: Valuation & estimates
- Pages 11-12: Appendix (optional)

### Phase 5: Quality Check & Delivery (30 minutes)
Verify content, formatting, accuracy, and timeliness before delivery.

**See [references/best-practices.md](references/best-practices.md)** for quality checklist and common mistakes to avoid.

## Output Specification

**Primary Deliverable**: DOCX report (8-12 pages) or Markdown earnings note
**Default Workspace Path**: `research/companies/<TICKER>/notes/<TICKER>_Q[Quarter]FY[Year]_Earnings_YYYY-MM-DD.md` (or `research/companies/<TICKER>/reports/<TICKER>_Q[Quarter]FY[Year]_Earnings_YYYY-MM-DD.docx` for full Word reports)
**Example**: `research/companies/NKE/reports/NKE_Q2FY24_Earnings_2024-01-05.docx`

**Contents:**
- Page 1: Summary with rating, price target, key takeaways
- Pages 2-3: Detailed results analysis
- Pages 4-5: Key metrics and guidance
- Pages 6-7: Updated thesis assessment
- Pages 8-10: Valuation and estimates
- Pages 11-12: Appendix (optional)
- 8-12 embedded charts
- 1-3 summary tables
- Complete sources section with clickable hyperlinks

**Optional Deliverable**: XLS model update (optional for earnings updates)

## Key Differences from Initiation Report

| Aspect | Earnings Update | Initiation Report |
|--------|----------------|-------------------|
| **Length** | 8-12 pages | 30-50 pages |
| **Words** | 3,000-5,000 | 10,000-15,000 |
| **Tables** | 1-3 summary | 12-20 comprehensive |
| **Figures** | 8-12 | 25-35 |
| **Turnaround** | 1-2 days | 3-6 weeks |
| **Scope** | Quarterly results | Complete company |
| **Focus** | What's NEW | Everything |
| **Company Background** | Brief mention | 6-10 pages |
| **XLS Model** | Optional | Required |

## Resources

### references/workflow.md
Detailed Phase 1-5 instructions with step-by-step procedures for data collection, analysis, chart generation, and report creation.

### references/report-structure.md
Complete page-by-page templates, table formats, and formatting requirements for the DOCX report.

### references/best-practices.md
Examples of good/bad headlines, tips for success, common mistakes to avoid, and comprehensive quality checklist.

## Dependencies

**Required:**
- Python (`matplotlib`, `pandas`, `seaborn`) for chart generation
- `python-docx` for `.docx` report creation

**Optional:**
- `xlsx-author` / `openpyxl` for model updates saved under `research/companies/<TICKER>/models/` (not required for earnings updates)
