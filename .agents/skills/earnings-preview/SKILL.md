---
name: earnings-preview
description: Build pre-earnings setup notes with consensus benchmarks, scenario frameworks, and key KPIs to watch before a company reports quarterly earnings. Triggers on "earnings preview", "what to watch for [company] earnings", "pre-earnings", "earnings setup", or "preview Q[X] for [company]". Not for post-earnings print analysis (use earnings-analysis).
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/earnings-preview/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# Earnings Preview

Draft concise pre-earnings setup notes staged as `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`.

## Workflow

### Step 1: Gather Context & Consensus (Strict Source Hierarchy)

- Identify the company, reporting quarter, fiscal year-end month, and confirmed/projected earnings date & time (pre-market vs. after-hours).
- **Consensus Source Hierarchy (MANDATORY):**
  1. Connected financial MCP server (FactSet, S&P Capital IQ, LSEG, Bloomberg)
  2. Pre-release wire articles (Reuters, Bloomberg, WSJ) quoting a named consensus provider dated **prior to** the release
  3. Company IR-compiled consensus (where published on IR site)
  - Always record **provider, as-of date, estimate count (`n`), and accounting basis (GAAP vs. Non-GAAP)**.
  - **NEVER** fabricate or cite unverified "whisper numbers", and **NEVER** reuse post-print Yahoo Finance consensus snapshots as pre-print expectations.
  - **Fallback:** If verified pre-print consensus is unavailable, anchor the preview on **prior management guidance** and **internal model estimates**, marking Street consensus **`[UNSOURCED]`**.
- Review the prior quarter's earnings release, 10-Q/10-K, and earnings call transcript for active guidance ranges and management commentary.
- **Standalone first-time preview (no prior coverage model under `research/companies/<TICKER>/models/`):** Do **not** block or redirect into `/initiating-coverage`; build the baseline directly from the latest `10-K`, prior `10-Q`s, active management guidance, and verified pre-print consensus, and tag the header **`[NO PRIOR COVERAGE MODEL — STANDALONE EARNINGS PREVIEW]`**.

### Step 2: Key Metrics Framework

Build a ranked "what to watch" framework specific to the company, explicitly separating **GAAP** and **Non-GAAP** metrics and defining the **in-line threshold** (`±1.0%` revenue, `±2.0%` or `±$0.01` EPS):

**Financial Metrics:**
- Revenue vs. consensus / prior guidance (total and key segments)
- EPS (GAAP and Non-GAAP matched strictly to consensus basis)
- Margins (gross, operating, EBITDA) — expansion vs. compression drivers
- Free cash flow and capital returns
- Forward quarter / full-year guidance vs. consensus

**Operational Metrics (sector-specific, only where disclosed):**
- Tech/SaaS: ARR, NDR/GRR, cRPO/RPO, large-customer additions
- Retail: Same-store sales, traffic vs. ticket, inventory turns
- Industrials: Backlog, book-to-bill, price vs. volume
- Financials: NIM, NCOs/credit provisions, loan growth, fee income
- Healthcare: Scripts, patient volumes, clinical/regulatory milestones

### Step 3: Scenario Analysis

Build 3 scenarios (tag all analyst estimates `[ASSUMPTION]`):

| Scenario | Revenue | EPS (Basis) | Guidance / Key Driver | Implied Stock Reaction |
|---|---|---|---|---|
| **Bull** | | | | |
| **Base** | | | | |
| **Bear** | | | | |

For each scenario, specify what operational print is required, what management commentary would confirm it, and historical post-earnings move context (cited with dates).

### Step 4: Catalyst Checklist

Identify the 3–5 specific swing factors that will drive the stock reaction:
1. `[Headline KPI]` vs. `[verified consensus / prior guidance]` — why it moves the multiple
2. `[Forward guidance item]` — what the market needs to hear to hold/raise FY+1 estimates
3. `[Strategic / narrative watch item]` — product ramp, margin inflection, capital allocation, or macro demand read-through

### Step 5: Output

- **Default output path:** `research/companies/<TICKER>/notes/<TICKER>_Q[X]FY[YY]_Earnings_Preview_YYYY-MM-DD.md`
- Label header **`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`**.
- Include: company/quarter/date (`Confirmed` vs. `Projected [ASSUMPTION]`), consensus table (provider, as-of date, `n`, GAAP/Non-GAAP basis), ranked KPIs, Bull/Base/Bear scenario table, catalyst checklist, and trading setup (recent price performance, valuation multiple, and options-implied move if sourced from a verified market feed — otherwise omit or mark `[UNSOURCED]`).
