---
name: idea-generation
description: Systematic stock screening and investment idea sourcing combining quantitative screens, thematic sweeps, and pattern recognition to surface long and short ideas. Triggers on "idea generation", "stock screen", "find ideas", "what looks interesting", "screen for", "new ideas", or "pitch me something". Not for overnight news recaps (use morning-note) or single-company valuation models (use dcf-model or comps-analysis).
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/idea-generation/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# Idea Generation

Run reproducible stock screens and thematic idea sweeps staged as `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`.

## Workflow

### Step 1: Define Screening Universe & Search Criteria

- **Interactive mode:** Confirm parameters with the user before screening:
  - **Universe (explicit):** e.g., S&P 500, Russell 1000/2000, a specific GICS sector ticker list, or a user-provided CSV/Excel export.
  - **Direction:** Long, short, or both
  - **Market cap & liquidity:** Large / mid / small / micro + minimum ADV
  - **Style:** Value, growth, quality, special situation, event-driven
  - **Geography & Theme:** US / international / global; optional thematic angle
- **Subagent mode (e.g., invoked by `market-researcher`):** Do not stall for input. Use the caller's specified sector/universe and criteria, default to US-listed liquid public equities (market cap `>= $1B`) if unspecified, tag inferred filters `[ASSUMPTION]`, and report exact screen parameters in the handoff.

### Step 2: Quantitative Screens & Verified Free-Data Sources

Prioritize connected financial MCP servers when configured; otherwise use verified free public data sources matched to each criterion:

| Screening Criterion | Verified Free-Data Source (No MCP Required) |
|---|---|
| Financials, revenue growth, margins, ROIC, ROE, FCF, debt/equity, AR/inventory trends | **SEC EDGAR XBRL** (`companyfacts` / `frames` REST API, `10-K`/`10-Q`) |
| Current price, market cap, LTM multiples (`EV/EBITDA`, `P/E`, `EV/Rev`, `P/B`), volume | **`yfinance`** cross-checked against latest 10-Q/10-K share counts & balance sheet |
| Insider buying / selling (last 90 days) & insider ownership | **SEC EDGAR Form 4** (`edgar` submissions API) & **DEF 14A** |
| Short interest & days-to-cover | **FINRA bi-monthly equity short interest** reports |
| Spin-offs, restructurings, 13D activists, auditor changes | **SEC EDGAR** Forms `8-K`, `10-12B`, `SC 13D`, `NT 10-K/Q` |

- **LTM Multiple Default:** Default all quantitative screens to **LTM multiples** computed from verified SEC filings and current market prices. Only screen on **NTM / forward consensus multiples** when a connected financial MCP server provides verified forward consensus — never fabricate forward consensus.
- **Mandatory Origin Tag:** Every candidate produced by an executed screen must record the exact filter pass/fail values and as-of date. Any qualitative or thematic idea surfaced **without** running a quantitative screen MUST be explicitly tagged **`[NOT SCREENED — ANALYST IDEA]`**.

**Screen Archetypes:**
- **Value:** LTM P/E below sector median, LTM EV/EBITDA below 5Y average, FCF yield `> 5%`, P/B `< 1.5x`, Form 4 insider buying (90d), dividend yield above median.
- **Growth:** Revenue growth `> 15%` YoY, operating income/EPS growth `> 20%` YoY, revenue acceleration, margin expansion, ROIC `> 15%`, NDR `> 110%` (where disclosed for SaaS).
- **Quality:** Consistent 5Y revenue growth, stable/expanding margins, ROE `> 15%`, low net debt/EBITDA, high FCF conversion (`FCF / Net Income > 80%`), insider ownership `> 5%`.
- **Short:** Decelerating/declining revenue, gross/operating margin compression, AR or inventory outgrowing revenue, Form 4 insider selling, unjustified peer multiple premium, rising FINRA short interest with weakening fundamentals, Form 8-K auditor/restatement red flags.
- **Special Situation:** Recent spin-offs (12m), post-restructuring exits, SC 13D activist campaigns, C-suite transitions at underperforming franchises.

### Step 3: Thematic Sweep

1. Define the thematic thesis and investment horizon.
2. Map the value chain (direct vs. second-order beneficiaries and bottlenecks).
3. Distinguish pure-play vs. diversified revenue exposure (citing 10-K segment `%`).
4. Flag which ideas passed a quantitative screen vs. **`[NOT SCREENED — ANALYST IDEA]`**.

### Step 4: Idea Presentation

For each shortlisted candidate (target **5–10 candidates**, from which `market-researcher` shortlists 3–5 highest-conviction ideas):

**[Company Name] ([TICKER]) — [Long/Short] — [Screened | `[NOT SCREENED — ANALYST IDEA]`]**

| Metric | Value (As-Of Date) | vs. Peer Median | Source |
|---|---|---|---|
| Market Cap / EV | | | |
| EV/EBITDA (LTM; or NTM if MCP) | | | |
| P/E (LTM; or NTM if MCP) | | | |
| Revenue Growth (YoY) | | | |
| EBITDA Margin | | | |
| FCF Yield | | | |

- **Thesis (3–5 bullets):** Mispricing, variant view, dated catalyst.
- **Key Risks & Disconfirming Evidence:** What breaks the thesis.
- **Suggested Next Steps:** Deep-dive (`initiating-coverage`), `dcf-model`, `comps-analysis`, or `unit-economics`.

### Step 5: Output

Save as an internal draft under `research/screens/`:
- **Primary screen memo:** `research/screens/YYYY-MM-DD_<screen-name>.md` — explicit universe, filter rules, data sources/dates, shortlist of 5–10 candidates (with `[NOT SCREENED — ANALYST IDEA]` tags where applicable), cross-idea comparison table, and prioritized research queue.
- **Optional quantitative screen export:** `research/screens/YYYY-MM-DD_<screen-name>.xlsx` via `xlsx-author` with live screening formulas and verified via `scripts/recalc_xlsx.py`.
