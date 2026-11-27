---
name: unit-economics
description: Analyze unit economics for public equities, SaaS/subscription, and recurring-revenue businesses — ARR bridges, cohort retention, LTV/CAC, net/gross dollar retention, payback periods, revenue quality, and margin waterfalls. Use when evaluating revenue quality, building a cohort analysis, or assessing customer economics. Triggers on "unit economics", "cohort analysis", "ARR analysis", "LTV CAC", "net retention", "revenue quality", or "customer economics".
---

# Unit Economics Analysis

## Workflow

### Step 1: Identify Business Model

Determine the revenue model to tailor the analysis:
- **SaaS / Subscription**: ARR, net retention, cohorts
- **Recurring services**: Contract value, renewal rates, upsell
- **Transaction / usage-based**: Revenue per transaction, volume trends, take rate
- **Hybrid**: Break down by revenue stream

### Step 2: Core Metrics

#### ARR / Revenue Quality
- **ARR bridge**: Beginning ARR → New → Expansion → Contraction → Churn → Ending ARR
- **ARR by cohort**: Vintage analysis — how does each annual cohort retain and grow?
- **Revenue concentration**: Top 10/20/50 customers as % of total
- **Revenue by type**: Recurring vs. non-recurring vs. professional services
- **Contract structure**: ACV distribution, multi-year %, auto-renewal %

#### Customer Economics
- **CAC (Customer Acquisition Cost)**: Total S&M spend / new customers acquired
- **LTV (Lifetime Value)**: (ARPU × Gross Margin) / Churn Rate
- **LTV:CAC ratio**: Target >3x for healthy businesses
- **CAC payback period**: Months to recover acquisition cost
- **Blended vs. segmented**: Break down by customer segment (enterprise vs. SMB vs. mid-market)

#### Retention & Expansion
- **Gross retention**: % of beginning ARR retained (excludes expansion)
- **Net retention (NDR)**: % of beginning ARR retained including expansion
- **Logo churn**: % of customers lost
- **Dollar churn**: % of revenue lost (often different from logo churn)
- **Expansion rate**: Upsell + cross-sell as % of beginning ARR

#### Cohort Analysis
Build a cohort matrix showing:

| Cohort | Year 0 | Year 1 | Year 2 | Year 3 | Year 4 |
|--------|--------|--------|--------|--------|--------|
| 2020 | $1.0M | $1.1M | $1.2M | $1.1M | |
| 2021 | $1.5M | $1.7M | $1.8M | | |
| 2022 | $2.0M | $2.3M | | | |
| 2023 | $3.0M | | | | |

Show both absolute $ and indexed (Year 0 = 100%) views.

#### Margin Waterfall
- Revenue → Gross Profit → Contribution Margin → EBITDA
- Fully loaded unit economics: what does it cost to acquire, serve, and retain a customer?
- Gross margin by revenue stream (subscription vs. services vs. other)

### Step 3: Benchmarking

Compare unit economics to relevant benchmarks:
- **SaaS Rule of 40**: Growth rate + EBITDA margin > 40%
- **SaaS Magic Number**: Net new ARR / prior period S&M spend > 0.75x
- **NDR benchmarks**: Best-in-class >120%, good >110%, concerning <100%
- **LTV:CAC**: Best-in-class >5x, good >3x, concerning <2x
- **Gross retention**: Best-in-class >95%, good >90%, concerning <85%
- **CAC payback**: Best-in-class <12mo, good <18mo, concerning >24mo

### Step 4: Revenue Quality Score

Synthesize into a revenue quality assessment:

| Factor | Score (1-5) | Notes |
|--------|-------------|-------|
| Recurring % | | |
| Net retention | | |
| Customer concentration | | |
| Cohort stability | | |
| Growth durability | | |
| Margin profile | | |
| **Overall** | | |

### Step 5: Output

Stage all outputs under `research/companies/<TICKER>/`:
- **Model workbook**: `research/companies/<TICKER>/models/<TICKER>_UnitEconomics_YYYY-MM-DD.xlsx` via `xlsx-author` (ARR/billings bridge, cohort matrix or public-filing proxy panel, unit economics dashboard)
- **Analytical note / summary**: `research/companies/<TICKER>/notes/<TICKER>_UnitEconomics_YYYY-MM-DD.md` (with optional slide via `pptx-author` under `research/companies/<TICKER>/reports/`) highlighting key benchmarks, red flags, and diligence questions

## Public-Filing Proxies & Important Notes

- **Public-equity data reality**: Unlike private equity data rooms, public companies rarely disclose raw customer-level tables. When granular cohorts are unavailable, build transparent proxies from SEC filings (`10-K`, `10-Q`), earnings releases, and investor day decks — explicitly labeling derived figures `[ASSUMPTION]`:
  - **RPO & cRPO** (Remaining Performance Obligations / current RPO) to track contracted backlog and forward revenue visibility
  - **Calculated billings** (`Revenue + ΔDeferred Revenue`) and **deferred revenue** trends to cross-check demand vs. GAAP revenue recognition
  - **Disclosed NRR / DBNRR, customer-count tiers** (e.g., `$100K+` / `$1M+` ARR logos), and **implied ACV** (`Subscription Revenue / Ending Customers`)
  - **Public-filing S&M efficiency**: Implied Magic Number (`ΔSubscription Revenue × 4 / Prior-Quarter S&M` or `Net New ARR / Prior-Year S&M`), implied gross-margin-adjusted CAC payback, and S&M % of revenue
- If the user uploads customer-level exports or investor-day cohort disclosures, prioritize those primary datasets and document exact source dates
- NDR above 100% can mask elevated gross churn if expansion is strong enough — flag when GRR is undisclosed
- Differentiate clearly between contracted ARR, billings, and GAAP recognized revenue
- For usage-based models, focus on consumption trends, RPO drawdowns, and expansion patterns rather than traditional seat-based ARR metrics
- Evaluate professional services revenue separately — it is non-recurring and carries structurally lower gross margins
