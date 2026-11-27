---
name: unit-economics
description: Analyze unit economics for public equities, SaaS/subscription, and recurring-revenue businesses — ARR/billings bridges, cohort retention, lagged CAC, LTV/CAC, GRR/NDR, CAC payback, Rule of 40, and margin waterfalls. Triggers on "unit economics", "cohort analysis", "ARR analysis", "LTV CAC", "net retention", "revenue quality", or "customer economics".
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/private-equity/skills/unit-economics/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# Unit Economics Analysis

Evaluate revenue quality, customer retention, acquisition efficiency, and contribution margins for public equities and recurring-revenue businesses. Stage outputs as `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY`.

## Workflow

### Step 1: Identify Business Model
Tailor KPIs to the monetization model:
- **SaaS / Subscription:** ARR, RPO/cRPO, calculated billings, GRR/NDR, CAC payback, Rule of 40
- **Recurring Services:** Contract value, renewal rates, expansion attach
- **Transaction / Usage-Based:** GMV/GPV, take rate, active buyers/sellers, repeat volume
- **Hybrid:** Separate recurring subscription lines from lower-margin professional services / hardware

### Step 2: Core Metrics & Exact Institutional Formulas

#### ARR & Revenue Quality
- **ARR Bridge:** `Ending ARR = Beginning ARR + New Logo ARR + Expansion ARR - Contraction ARR - Churned ARR`
- **Calculated Billings (Public-Filing Proxy):** `Calculated Billings = GAAP Revenue + ΔDeferred Revenue` (verify short-term + long-term deferred revenue)
- **Revenue Mix & Concentration:** Recurring vs. services `%`; top 10/20 customers `%` (where disclosed in 10-K)

#### Retention & Expansion (Explicit Formulas)
- **Gross Dollar Retention (GRR):**
  $$\text{GRR} = \frac{\text{Beginning ARR} - \text{Contraction ARR} - \text{Churned ARR}}{\text{Beginning ARR}} \quad (\le 100\%)$$
- **Net Dollar Retention (NDR / DBNRR):**
  $$\text{NDR} = \frac{\text{Beginning ARR} + \text{Expansion ARR} - \text{Contraction ARR} - \text{Churned ARR}}{\text{Beginning ARR}}$$
- **Logo Churn:** `Churned Customers_t / Beginning Customers_t`
- **Dollar Churn:** `(Contraction ARR + Churned ARR) / Beginning ARR`

#### Customer Acquisition & Lifetime Value (Lagged CAC Basis)
- **Lagged CAC (MANDATORY for multi-month sales cycles):**
  $$\text{CAC}_t = \frac{\text{Sales \& Marketing Spend}_{t-1}}{\text{New Customers Acquired}_t}$$
  *(Match the lag `t-1` to the company's sales cycle — typically 1 quarter for mid-market/enterprise SaaS — rather than dividing contemporaneous S&M by current-period closes. Explicitly note when contemporaneous S&M is used for PLG / zero-lag self-serve models.)*
- **CAC Payback Period (Months):**
  $$\text{CAC Payback (months)} = \frac{\text{CAC}_t}{\text{Monthly ARPU} \times \text{Subscription Gross Margin \%}}$$
  *(Public-filing macro proxy when customer counts are undisclosed: `Prior-Period S&M / (Net New ARR × Subscription Gross Margin %) × 12`.)*
- **Lifetime Value (LTV):**
  $$\text{LTV} = \frac{\text{Annual ARPU} \times \text{Subscription Gross Margin \%}}{\text{Annual Churn Rate}}$$
  *(Use logo churn for per-customer LTV or gross dollar churn for revenue-weighted LTV; state basis explicitly and note that LTV explodes artificially if churn approaches 0% — cap assumed customer lifetime at 7–10 years when modeling.)*
- **LTV : CAC Ratio:** `LTV / Lagged CAC` (target `> 3.0x`)

#### Cohort Matrix & Margin Waterfall
- Build a cohort retention table (`Cohort × Year 0..Year 4`) showing both absolute `$` and indexed (`Year 0 = 100%`) views where cohorts are disclosed (never fabricate undisclosed cohorts!).
- **Margin Waterfall:** `Revenue → Subscription Gross Profit → Contribution Margin (less S&M/Support) → Operating Income / EBITDA`. Always separate GAAP vs. Non-GAAP (SBC treatment).

### Step 3: Benchmarking (Aligned with `comps-analysis`)
- **Rule of 40:** `YoY Revenue Growth % + FCF Margin %` *(or `YoY Revenue Growth % + EBITDA Margin %` — **ALWAYS label explicitly** which margin metric is used so benchmarks match `comps-analysis` cleanly!)* Target `>= 40%`.
- **SaaS Magic Number:** `Net New ARR / Prior-Period S&M Spend` (quarterly: `(Subscription Rev_t - Subscription Rev_{t-1}) × 4 / S&M_{t-1}`). Healthy `> 0.75x`; efficient `> 1.0x`.
- **NDR:** Best-in-class `> 120%`, good `> 110%`, concerning `< 100%`.
- **GRR:** Best-in-class `> 95%`, good `> 90%`, concerning `< 85%`.
- **LTV : CAC:** Best-in-class `> 5.0x`, good `> 3.0x`, concerning `< 2.0x`.
- **CAC Payback:** Best-in-class `< 12 mo`, good `< 18 mo`, concerning `> 24 mo`.

### Step 4: Revenue Quality Scorecard

| Factor | Score (1–5) | Evidence & Primary Source |
|---|---|---|
| Recurring Revenue `%` | | |
| Net & Gross Dollar Retention | | |
| Customer Concentration | | |
| Cohort / RPO Visibility | | |
| S&M Efficiency (Magic # / Payback) | | |
| Subscription Gross Margin & Rule of 40 | | |
| **Overall Revenue Quality** | | |

### Step 5: Canonical `.xlsx` Workbook Layout & Output Paths

Stage deliverables under `research/companies/<TICKER>/` (aligned with `model-builder`):
- **Model workbook:** `research/companies/<TICKER>/models/<TICKER>_UnitEconomics_YYYY-MM-DD.xlsx` built via `xlsx-author` with live formulas, institutional palette (`#1F4E79` / `#D9E1F2` / `#F2F2F2` / `#BDD7EE`), source comments on all blue inputs, and **5 canonical tabs**:
  1. `Summary Dashboard` — Headline KPIs (ARR, Growth, GRR/NDR, Lagged CAC, Payback, LTV/CAC, Magic Number, Rule of 40 labeled FCF vs. EBITDA basis, Revenue Quality Scorecard)
  2. `ARR & Billings Bridge` — Beginning ARR → New / Expansion / Contraction / Churn → Ending ARR, plus GAAP Revenue + ΔDeferred Revenue calculated billings & RPO/cRPO table
  3. `Cohort & Proxy Panel` — Disclosed customer cohorts (absolute `$` & `%` indexed) OR transparent public-filing proxies (customer tiers `$100K+`/`$1M+`, implied ACV, DBNRR) tagged `[ASSUMPTION]` where derived
  4. `Unit Economics & Waterfall` — Lagged CAC, Payback, LTV/CAC, subscription vs. services gross margin, contribution margin waterfall
  5. `Sources & Notes` — Filing dates, GAAP/Non-GAAP reconciliations, SBC treatment, and proxy methodology
- **Mandatory QA Gate:** Run `python3 scripts/recalc_xlsx.py research/companies/<TICKER>/models/<TICKER>_UnitEconomics_YYYY-MM-DD.xlsx 30` (exit `0`, `"status": "success"`).
- **Analytical summary note:** `research/companies/<TICKER>/notes/<TICKER>_UnitEconomics_YYYY-MM-DD.md` (with optional slides via `pptx-author` under `research/companies/<TICKER>/reports/`).

## Public-Filing Data Reality

- Public companies rarely disclose raw customer-level tables. When granular cohorts are unavailable in 10-K/10-Q/Investor Day materials, build transparent proxies from RPO/cRPO, calculated billings, disclosed DBNRR, customer tiers, and lagged S&M efficiency — explicitly tagging derived proxies `[ASSUMPTION]` and unavailable figures `[UNSOURCED]`.
- Flag whenever NDR `> 100%` is reported without GRR (strong expansion can mask elevated gross churn).
- Evaluate professional services separately from subscription software.
