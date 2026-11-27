---
name: comps-analysis
description: Build institutional comparable company analyses with operating metrics, LTM/NTM trading valuation multiples (EV/Revenue, EV/EBITDA, P/E), and statistical benchmarking (Max, 75th, Median, 25th, Min) in Excel (.xlsx). Use for relative public equity valuation, peer benchmarking, and DCF exit-multiple cross-checks. Not for qualitative moat/positioning decks (use competitive-analysis) or intrinsic DCF valuation (use dcf-model).
---

<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/comps-analysis/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# Comparable Company Analysis

## When to Use

**Ideal for:**
- Public company valuation and relative multiple analysis
- Benchmarking growth, margins, and efficiency vs. industry peers
- Informing DCF terminal exit multiples and sensitivity ranges
- Identifying valuation outliers (premium/discount analysis)
- Supporting initiating coverage, sector overviews, and investment memos

**Less suitable for:**
- Private companies without comparable public peers
- Highly diversified conglomerates (use sum-of-the-parts instead)
- Distressed/bankrupt companies or pre-revenue businesses with unique models

## ⚠️ Data Source Priority

Follow this data source hierarchy:

1. **Connected financial MCPs** (if configured in `.agents/mcp_config.json`, e.g., FactSet, S&P Global/Kensho, Daloopa, LSEG) — use first for verified historicals, consensus, and trading multiples.
2. **Primary filings & structured data** — SEC EDGAR (10-K, 10-Q, 8-K), company investor relations releases, or user-provided spreadsheets in the workspace / Google Drive.
3. **Market data & web search** — Use for current share prices, diluted share counts, or recent announcements when MCPs are unavailable; verify against filings and cite every source and date. Never fabricate figures — mark unverified fields `[UNSOURCED]`.

---

## Overview

Build institutional-grade comparable company analyses that combine operating metrics, valuation multiples, and statistical benchmarking in a structured Excel workbook (`.xlsx`) that enables informed investment decisions through peer comparison.

**ALWAYS confirm context first:**
1. **"Do you have a preferred template or should I use the standard institutional layout?"**
2. **"Who is the audience?"** (Investment committee, coverage initiation, sector primer, quick reference)
3. **"What's the key question?"** (Valuation, growth analysis, competitive positioning, capital efficiency)
4. **"What's the context?"** (Single-name valuation, sector benchmarking, DCF cross-check, portfolio review)

**Adapt based on specifics:**
- **Industry context**: Big tech mega-caps need different metrics than emerging SaaS companies
- **Sector-specific needs**: Add relevant KPIs early (e.g., ARR, NDR, Rule of 40 for SaaS; SSS for retail; ROE/P/TBV for financials; R&D/Rev for biotech)
- **Company familiarity**: Well-known names need less background and more focus on delta and multiple dispersion

**Core principle:** Use institutional structure (clear headers, transparent formulas, statistical rigor, audit trails) while adapting metrics to the sector. User-provided examples and explicit preferences always take precedence over defaults.

## Core Philosophy

**"Build the right structure first, then let the data tell the story."**

Start with headers that force strategic thinking about what matters, input clean sourced data, build transparent formulas, and let statistics emerge automatically. A good comp should be immediately readable by someone who didn't build it.

---

## ⚠️ CRITICAL: Formulas Over Hardcodes + Step-by-Step Verification

**Execution Environment (Antigravity):**
- Build standalone `.xlsx` workbooks programmatically with Python (`openpyxl`). Write formula strings (`cell.value = "=E7/C7"`), never pre-computed results.
- **Default output paths** (aligned with `.agents/rules/research-outputs.md`): save single-company peer comps to `research/companies/<TICKER>/models/<TICKER>_Comps_YYYY-MM-DD.xlsx` and sector-level peer comps to `research/sectors/<sector-slug>/comps/<sector-slug>_comps_YYYY-MM-DD.xlsx`.
- After generating the workbook, recalculate and verify zero formula errors by running `python3 scripts/recalc_xlsx.py <path-to-workbook.xlsx>` and checking against `audit-xls`.
- If the user wants the comps table synced to Google Sheets and the `gsheets` plugin is connected, export or write the structured workbook cleanly after verifying formulas locally.

**Formulas, not hardcodes:**
- Every derived value (margin, multiple, statistic) MUST be an Excel formula referencing input cells — never a pre-computed number pasted in
- When using Python/openpyxl to build the sheet: write `cell.value = "=E7/C7"` (formula string), NOT `cell.value = 0.687` (computed result)
- The only hardcoded values should be raw input data (revenue, EBITDA, share price, etc.) — and every one of those gets a cell comment with its source
- Why: the model must update automatically when an input changes. A hardcoded margin is a silent bug waiting to happen.

**Verify step-by-step (Interactive vs. Subagent Mode):**
- **Interactive mode:** Confirm at key checkpoints — (1) header & peer-set structure, (2) raw sourced inputs & fiscal calendarization, (3) operating margins sanity check, (4) valuation multiples & quartile statistics — before finalizing.
- **Subagent mode (e.g., invoked by `market-researcher` or `model-builder`):** Do not stall for mid-build confirmation. Build the complete workbook end-to-end, tag inferred assumptions `[ASSUMPTION]` and missing feeds `[UNSOURCED]`, run `python3 scripts/recalc_xlsx.py <file.xlsx> 30` (confirming exit `0`), and surface the full peer-set rationale, calendarization notes, and outlier checks in the handoff.

---

## Section 1: Document Structure & Setup

### Header Block (Rows 1-3)
```
Row 1: [ANALYSIS TITLE] - COMPARABLE COMPANY ANALYSIS
Row 2: [List of Companies with Tickers] • [Company 1 (TICK1)] • [Company 2 (TICK2)] • [Company 3 (TICK3)]
Row 3: As of [Period] | All figures in [USD Millions/Billions] except per-share amounts and ratios
```

**Why this matters:** Establishes context immediately. Anyone opening this file knows what they're looking at, when it was created, and how to interpret the numbers.

### Visual Convention Standards (OPTIONAL - User preferences and uploaded templates always override)

**IMPORTANT: These are suggested defaults only. Always prioritize:**
1. User's explicit formatting preferences
2. Formatting from any uploaded template files
3. Company/team style guides
4. These defaults (only if no other guidance provided)

**Suggested Font & Typography:**
- **Font family**: Times New Roman (professional, readable, industry standard)
- **Font size**: 11pt for data cells, 12pt for headers
- **Bold text**: Section headers, company names, statistic labels

**Default Color & Shading — Professional Blue/Grey Palette (minimal is better):**
- **Keep it restrained** — only blues and greys. Do NOT introduce greens, oranges, reds, or multiple accent colors. A clean comps sheet uses 3-4 colors total.
- **Section headers** (e.g., "OPERATING STATISTICS & FINANCIAL METRICS"):
  - Dark blue background (`#1F4E79` or `#17365D` navy)
  - White bold text
  - Full row shading across all columns
- **Column headers** (e.g., "Company", "Revenue", "Margin"):
  - Light blue background (`#D9E1F2` or similar pale blue)
  - Black bold text
  - Centered alignment
- **Data rows**:
  - White background for company data
  - Black text for formulas; blue text for hardcoded inputs
- **Statistics rows** (Maximum, 75th Percentile, etc.):
  - Light grey background (`#F2F2F2`)
  - Black text, left-aligned labels
- **That's the whole palette**: dark blue + light blue + light grey + white. Nothing else unless the user's template says otherwise.

**Suggested Formatting Conventions:**
- **Decimal precision**:
  - Percentages: 1 decimal (12.3%)
  - Multiples: 1 decimal (13.5x)
  - Dollar amounts: No decimals, thousands separator (69,632)
  - Margins shown as percentages: 1 decimal (68.7%)
- **Borders**: No borders (clean, minimal appearance)
- **Alignment**: All metrics center-aligned for clean, uniform appearance
- **Cell dimensions**: All column widths should be uniform/even, all row heights should be consistent (creates clean, professional grid)

**Note:** If the user provides a template file or specifies different formatting, use that instead.

---

## Section 2: Operating Statistics & Financial Metrics

### Core Columns (Start with these)
1. **Company** - Names with consistent formatting
2. **Revenue** - Size metric (can be LTM, quarterly, or annual depending on context)
3. **Revenue Growth** - Year-over-year percentage change
4. **Gross Profit** - Revenue minus cost of goods sold
5. **Gross Margin** - GP/Revenue (fundamental profitability)
6. **EBITDA** - Earnings before interest, tax, depreciation, amortization
7. **EBITDA Margin** - EBITDA/Revenue (operating efficiency)

### Optional Additions (Choose based on industry/purpose)
- **Quarterly vs LTM** - Include both if seasonality matters
- **Free Cash Flow** - For capital-intensive or SaaS businesses
- **FCF Margin** - FCF/Revenue (cash generation efficiency)
- **Net Income** - For mature, profitable companies
- **Operating Income** - For businesses with varying D&A
- **CapEx metrics** - For asset-heavy industries
- **Rule of 40** - Specifically for SaaS (Growth % + Margin %)
- **FCF Conversion** - For quality of earnings analysis (advanced)

### Formula Examples (Using Row 7 as example)
```excel
// Core ratios - these are always calculated
Gross Margin (F7): =E7/C7
EBITDA Margin (H7): =G7/C7

// Optional ratios - include if relevant
FCF Margin: =[FCF]/[Revenue]
Net Margin: =[Net Income]/[Revenue]
Rule of 40: =[Growth %]+[FCF Margin %]
```

**Golden Rule:** Every ratio should be [Something] / [Revenue] or [Something] / [Something from this sheet]. Keep it simple.

### Statistics Block (After company data)

**CRITICAL: Add statistics formulas for all comparable metrics (ratios, margins, growth rates, multiples).**

```
[Leave one blank row for visual separation]
- Maximum: =MAX(B7:B9)
- 75th Percentile: =QUARTILE(B7:B9,3)
- Median: =MEDIAN(B7:B9)
- 25th Percentile: =QUARTILE(B7:B9,1)
- Minimum: =MIN(B7:B9)
```

**Columns that NEED statistics (comparable metrics):**
- Revenue Growth %, Gross Margin %, EBITDA Margin %, EPS
- EV/Revenue, EV/EBITDA, P/E, Dividend Yield %, Beta

**Columns that DON'T need statistics (size metrics):**
- Revenue, EBITDA, Net Income (absolute size varies by company scale)
- Market Cap, Enterprise Value (not comparable across different-sized companies)

**Note:** Add one blank row between company data and statistics rows for visual separation. Do NOT add a "SECTOR STATISTICS" or "VALUATION STATISTICS" header row.

**Why quartiles matter:** They show distribution, not just average. A 75th percentile multiple tells you what "premium" companies trade at.

---

## Section 3: Valuation Multiples & Investment Metrics

### Core Valuation Columns (Start with these)
1. **Company** — Same order as operating section
2. **Market Cap** — Current Share Price × Fully Diluted Shares (TSM)
3. **Enterprise Value** — Full institutional formula (strictly apples-to-apples with the `dcf-model` equity bridge):
   `Enterprise Value = Market Cap + Total Gross Debt + Preferred Stock + Non-Controlling Interest - Cash & Short-Term Investments - Equity Investments`
   *(Subtract non-operating unconsolidated equity-method investments so EV/EBITDA exit multiples are apples-to-apples with the DCF bridge; exclude ASC 842 operating lease liabilities unless peer EBITDA is consistently adjusted on an IFRS 16 / EBITDAR basis.)*
4. **EV/Revenue (LTM & NTM / FY+1)** — How much the market pays per dollar of sales
5. **EV/EBITDA (LTM & NTM / FY+1)** — How much the market pays per dollar of operating cash flow proxy
6. **P/E Ratio (LTM & NTM / FY+1)** — Price relative to net earnings (match GAAP vs. Non-GAAP strictly across all peers!)

### Calendarization & Forward Multiples (NTM / FY+1)
Public peers often have different fiscal year-ends (e.g., Jan, Mar, Jun, Sep, Dec). Always align forward multiples to a **common 12-month window (NTM or CY+1)** using weighted fiscal-year interpolation when consensus estimates are available (e.g., `NTM Metric = (m/12) * FY1E + ((12 - m)/12) * FY2E`, where `m` is months remaining in `FY1`). If forward consensus is unavailable via connected MCPs, default to **LTM** from primary SEC filings and disclose that NTM multiples are `[UNSOURCED]`.

### Optional Valuation Metrics (Choose based on context)
- **FCF Yield** — `FCF / Market Cap` (for cash-focused analysis)
- **PEG Ratio** — `P/E / (EPS Growth Rate in percentage points)` (when growth is stored as a decimal like `0.15`, divide by `Growth * 100`!)
- **Price/Book** — Market value vs. book value (for financials/asset-heavy businesses)
- **ROE/ROA** — Return metrics
- **Rule of 40** — `YoY Revenue Growth % + FCF Margin %` (or `+ EBITDA Margin %` — ALWAYS label explicitly whether FCF or EBITDA margin is used!)

**Key Principle:** Include 3–5 core multiples that matter for your industry.

### Formula Examples (Always Guard Negative/Zero Denominators with `"NM"`)
```excel
// Enterprise Value (full formula, matching DCF equity bridge)
Enterprise Value: =MarketCap + TotalGrossDebt + PreferredStock + NCI - CashAndSTI - EquityInvestments

// Core multiples — return "NM" (NOT 0!) for non-meaningful/negative/zero denominators
// Returning "NM" ensures MEDIAN / QUARTILE / AVERAGE ignore invalid rows instead of skewing toward zero
EV/Revenue:  =IF(Revenue<=0, "NM", IFERROR(EV/Revenue, "NM"))
EV/EBITDA:   =IF(EBITDA<=0, "NM", IFERROR(EV/EBITDA, "NM"))
P/E Ratio:   =IF(NetIncome<=0, "NM", IFERROR(MarketCap/NetIncome, "NM"))

// Optional multiples
FCF Yield:   =IFERROR(LTM_FCF/MarketCap, "NM")
// PEG Ratio (decimal growth rate e.g. 0.15 for 15% MUST be multiplied by 100!)
PEG Ratio:   =IF(OR(PE="NM", Growth<=0), "NM", IFERROR(PE/(Growth*100), "NM"))
```

### Cross-Reference Rule
**CRITICAL:** Valuation multiples MUST reference the operating metrics section. Never input the same raw data twice. If revenue is in C7, then EV/Revenue formula should reference C7.

### Statistics Block
Same structure as operating section: Max, 75th, Median, 25th, Min for every comparable metric (Excel `MEDIAN`, `QUARTILE`, `MAX`, `MIN` automatically ignore `"NM"` text strings). Add one blank row for visual separation between company data and statistics. Do NOT add a "VALUATION STATISTICS" header row.

---

## Section 4: Notes & Methodology Documentation

### Required Components

**Data Sources & Quality:**
- Where did the data come from? (Connected MCP servers — FactSet, S&P Global/Kensho, Daloopa, LSEG — SEC filings, verified market feeds)
- What period does it cover? (LTM / calendarized NTM, as-of date)
- How was it verified? (Cross-checked against 10-K/10-Q; GAAP vs. Non-GAAP basis reconciled)

**Key Definitions:**
- EBITDA calculation method (`Operating Income (EBIT) + D&A`, noting explicitly whether SBC is excluded or included)
- Free Cash Flow formula (`Operating CF - CapEx`)
- Rule of 40 basis (explicitly state `Rev Growth % + FCF Margin %` OR `Rev Growth % + EBITDA Margin %`)
- Calendarization formula for off-calendar fiscal year-ends

**Valuation Methodology:**
- Enterprise Value = `Market Cap + Total Gross Debt + Preferred Stock + Non-Controlling Interest - Cash & Short-Term Investments - Equity Investments`
- Diluted share count basis (Treasury Stock Method + if-converted test)
- Outlier exclusions and adjustments documented transparently

**Analysis Framework:**
- What's the investment thesis? (Cloud/SaaS efficiency)
- What metrics matter most? (Cash generation, capital efficiency)
- How should readers interpret the statistics? (Quartiles provide context)

---

## Section 5: Choosing the Right Metrics (Decision Framework)

### Start with "What question am I answering?"

**"Which company is undervalued?"**
→ Focus on: EV/Revenue, EV/EBITDA, P/E, Market Cap
→ Skip: Operational details, growth metrics

**"Which company is most efficient?"**
→ Focus on: Gross Margin, EBITDA Margin, FCF Margin, Asset Turnover
→ Skip: Size metrics, absolute dollar amounts

**"Which company is growing fastest?"**
→ Focus on: Revenue Growth %, EBITDA CAGR, User/Customer Growth
→ Skip: Margin metrics, leverage ratios

**"Which is the best cash generator?"**
→ Focus on: FCF, FCF Margin, FCF Conversion, CapEx intensity
→ Skip: EBITDA, P/E ratios

### Industry-Specific Metric Selection

**Software/SaaS:**
Must have: Revenue Growth, Gross Margin, Rule of 40
Optional: ARR, Net Dollar Retention, CAC Payback
Skip: Asset Turnover, Inventory metrics

**Manufacturing/Industrials:**
Must have: EBITDA Margin, Asset Turnover, CapEx/Revenue
Optional: ROA, Inventory Turns, Backlog
Skip: Rule of 40, SaaS metrics

**Financial Services:**
Must have: ROE, ROA, Efficiency Ratio, P/E
Optional: Net Interest Margin, Loan Loss Reserves
Skip: Gross Margin, EBITDA (not meaningful for banks)

**Retail/E-commerce:**
Must have: Revenue Growth, Gross Margin, Inventory Turnover
Optional: Same-Store Sales, Customer Acquisition Cost
Skip: Heavy R&D or CapEx metrics

### The "5-10 Rule"

**5 operating metrics** - Revenue, Growth, 2-3 margins/efficiency metrics
**5 valuation metrics** - Market Cap, EV, 3 multiples
**= 10 total columns** - Enough to tell the story, not so many you lose the thread

If you have more than 15 metrics, you're probably including noise. Edit ruthlessly.

---

## Section 6: Best Practices & Quality Checks

### Before You Start
1. **Define the peer group** - Companies must be truly comparable (similar business model, scale, geography)
2. **Choose the right period** - LTM smooths seasonality; quarterly shows trends
3. **Standardize units upfront** - Millions vs. billions decision affects everything
4. **Map data sources** - Know where each number comes from

### As You Build
1. **Input all raw data first** - Complete the blue text before writing formulas
2. **Add cell comments to ALL hard-coded inputs** - Attach `openpyxl.comments.Comment(text, author)` to `cell.comment` documenting the source OR assumption

   **For sourced data, cite exactly where it came from:**
   - Example: "Bloomberg Terminal - MSFT Equity DES, accessed 2024-10-02"
   - Example: "Q4 2024 10-K filing, page 42, line item 'Total Revenue'"
   - Example: "FactSet consensus estimate as of 2024-10-02"
   - **Include hyperlinks when possible**: Set `cell.hyperlink = "<url>"` (with standard hyperlink styling) pointing to the SEC filing, data source, or report

   **For assumptions, explain the reasoning:**
   - Example: "Assumed 15% EBITDA margin based on peer median, company does not disclose"
   - Example: "Estimated Enterprise Value as Market Cap + $50M net debt (from Q3 balance sheet, Q4 not yet available)"
   - Example: "Forward P/E based on street consensus EPS of $3.45 (average of 12 analyst estimates)"

   **Why this matters**: Enables audit trails, data verification, assumption transparency, and future updates
3. **Build formulas row by row** - Test each calculation before moving on
4. **Use absolute references for headers** - $C$6 locks the header row
5. **Format consistently** - Percentages as percentages, not decimals
6. **Add conditional formatting** - Highlight outliers automatically

### Sanity Checks
- **Margin test**: Gross margin > EBITDA margin > Net margin (always true by definition)
- **Multiple reasonableness**: 
  - EV/Revenue: typically 0.5-20x (varies widely by industry)
  - EV/EBITDA: typically 8-25x (fairly consistent across industries)
  - P/E: typically 10-50x (depends on growth rate)
- **Growth-multiple correlation**: Higher growth usually means higher multiples
- **Size-efficiency trade-off**: Larger companies often have better margins (scale benefits)

### Common Mistakes to Avoid
❌ Mixing market cap and enterprise value in formulas
❌ Using different time periods for numerator and denominator (LTM vs quarterly)
❌ Hardcoding numbers into formulas instead of cell references
❌ **Hard-coded inputs without cell comments citing the source OR explaining the assumption**
❌ Missing hyperlinks to SEC filings or data sources when available
❌ Including too many metrics without clear purpose
❌ Including non-comparable companies (different business models)
❌ Using outdated data without disclosure
❌ Calculating averages of percentages incorrectly (should be median)

---

## Section 7: Advanced Features

### Dynamic Headers
For columns showing calculations, use clear unit labels:
```
Revenue Growth (YoY) % | EBITDA Margin | FCF Margin | Rule of 40
```

### Quartile Analysis Benefits
Instead of just mean/median, quartiles show:
- **75th percentile** = "Premium" companies trade here
- **Median** = Typical market valuation
- **25th percentile** = "Discount" territory

This helps answer: "Is our target company trading rich or cheap vs. peers?"

### Industry-Specific Modifications

**Software/SaaS:**
- Add: ARR, Net Dollar Retention, CAC Payback Period
- Emphasize: Rule of 40, FCF margins, gross margins >70%

**Healthcare:**
- Add: R&D/Revenue, Pipeline value, Regulatory status
- Emphasize: EBITDA margins, growth rates, reimbursement risk

**Industrials:**
- Add: Backlog, Order book trends, Geographic mix
- Emphasize: ROIC, asset turnover, cyclical adjustments

**Consumer:**
- Add: Same-store sales, Customer acquisition cost, Brand value
- Emphasize: Revenue growth, gross margins, inventory turns

---

## Section 8: Workflow & Practical Tips

### Step-by-Step Process
1. **Set up structure** (30 minutes)
   - Create all headers
   - Format cells (blue for inputs, black for formulas)
   - Lock in units and date references

2. **Gather data** (60-90 minutes)
   - Pull from primary sources (S&P Kensho MCP, FactSet MCP, Daloopa MCP if available; otherwise Bloomberg, SEC)
   - Input all raw numbers in blue
   - Document sources in notes section

3. **Build formulas** (30 minutes)
   - Start with simple ratios (margins)
   - Progress to multiples (EV/Revenue)
   - Add cross-checks (do margins make sense?)

4. **Add statistics** (15 minutes)
   - Copy formula structure for all columns
   - Verify ranges are correct (B7:B9, not B7:B10)
   - Check quartile logic

5. **Quality control** (30 minutes)
   - Run sanity checks
   - Verify formula references
   - Check for #DIV/0! or #REF! errors
   - Compare against known benchmarks

6. **Documentation** (15 minutes)
   - Complete notes section
   - Add data sources
   - Define methodologies
   - Date-stamp the analysis

### Pro Tips
- **Save templates**: Build once, reuse forever
- **Color-code outliers**: Conditional formatting for values >2 standard deviations
- **Link to source files**: Hyperlink to Bloomberg screenshots or SEC filings
- **Version control**: Save as "Comps_v1_2024-12-15" with clear dating
- **Collaborative reviews**: Have someone else check your formulas

### Excel Formatting Checklist (Aligned with `.agents/rules/financial-modeling.md`)
- [ ] Font set to user's preferred style (default: Calibri or Times New Roman, 10–11pt data, 11–12pt headers)
- [ ] Section headers formatted per palette (default: dark navy `#1F4E79` with white bold text)
- [ ] Column headers formatted per palette (default: light blue `#D9E1F2` with black bold text)
- [ ] Statistics rows formatted per palette (default: light grey `#F2F2F2`; highlight target/median row `#BDD7EE` if helpful)
- [ ] Column widths set cleanly with zero `###` truncation; consistent row heights
- [ ] Numbers formatted with proper decimal precision (`0.0%`, `0.0"x"`, `#,##0`)
- [ ] One blank row separating company rows from statistics rows
- [ ] Every hard-coded input cell has a comment citing exact source + date (or `[ASSUMPTION]` explanation)
- [ ] Hyperlinks added to SEC filings or data sources where applicable

---

## Section 9: Illustrative Placeholder Template Layout

*(Illustrative placeholder figures only — verify live numbers from primary filings/MCP before building. Notice `Gross Margin > EBITDA Margin` strictly holds for every peer!)*

```text
┌────────────────────────────────────────────────────────────────────────────┐
│ TECHNOLOGY — COMPARABLE COMPANY ANALYSIS (ILLUSTRATIVE PLACEHOLDER LAYOUT) │
│ Microsoft (MSFT) • Alphabet (GOOGL) • Amazon (AMZN)                        │
│ As of [Date] | All dollar figures in USD Millions                          │
├────────────────────────────────────────────────────────────────────────────┤
│ OPERATING METRICS (LTM)                                                    │
├──────────┬─────────┬─────────┬──────────┬─────────┬──────────┬─────────────┤
│ Company  │ Revenue │ Growth  │ Gross    │ EBITDA  │ EBITDA   │ FCF         │
│          │ (LTM)   │ (YoY)   │ Margin   │ (LTM)   │ Margin   │ Margin      │
├──────────┼─────────┼─────────┼──────────┼─────────┼──────────┼─────────────┤
│ MSFT     │ 261,800 │ 15.0%   │ 69.4%    │ 139,300 │ 53.2%    │ 27.5%       │
│ GOOGL    │ 350,000 │ 13.9%   │ 58.2%    │ 127,800 │ 36.5%    │ 20.8%       │
│ AMZN     │ 638,000 │ 11.0%   │ 48.9%    │ 141,000 │ 22.1%    │  6.0%       │
│          │         │         │          │         │          │             │
│ Max      │   —     │ =MAX()  │ =MAX()   │   —     │ =MAX()   │ =MAX()      │
│ 75th %   │   —     │ =QUART  │ =QUART   │   —     │ =QUART   │ =QUART      │
│ Median   │   —     │ =MEDIAN │ =MEDIAN  │   —     │ =MEDIAN  │ =MEDIAN     │
│ 25th %   │   —     │ =QUART  │ =QUART   │   —     │ =QUART   │ =QUART      │
│ Min      │   —     │ =MIN()  │ =MIN()   │   —     │ =MIN()   │ =MIN()      │
├────────────────────────────────────────────────────────────────────────────┤
│ VALUATION MULTIPLES (LTM & NTM)                                            │
├──────────┬──────────┬──────────┬─────────┬──────────┬──────────┬───────────┤
│ Company  │ Mkt Cap  │ EV*      │ EV/Rev  │ EV/EBITDA│ EV/EBITDA│ P/E       │
│          │          │          │ (LTM)   │ (LTM)    │ (NTM)    │ (NTM)     │
├──────────┼──────────┼──────────┼─────────┼──────────┼──────────┼───────────┤
│ MSFT     │3,100,000 │3,085,000 │ 11.8x   │ 22.1x    │ 19.6x    │ 31.0x     │
│ GOOGL    │2,200,000 │2,120,000 │  6.1x   │ 16.6x    │ 14.8x    │ 21.5x     │
│ AMZN     │2,150,000 │2,190,000 │  3.4x   │ 15.5x    │ 13.2x    │ 33.5x     │
│          │          │          │         │          │          │           │
│ Max      │   —      │   —      │ =MAX()  │ =MAX()   │ =MAX()   │ =MAX()    │
│ 75th %   │   —      │   —      │ =QUART  │ =QUART   │ =QUART   │ =QUART    │
│ Median   │   —      │   —      │ =MEDIAN │ =MEDIAN  │ =MEDIAN  │ =MEDIAN   │
│ 25th %   │   —      │   —      │ =QUART  │ =QUART   │ =QUART   │ =QUART    │
│ Min      │   —      │   —      │ =MIN()  │ =MIN()   │ =MIN()   │ =MIN()    │
└──────────┴──────────┴──────────┴─────────┴──────────┴──────────┴───────────┘
*EV = Market Cap + Total Gross Debt + Preferred + NCI - Cash & ST Investments - Equity Investments.
```

**Add complexity only when needed:**
- Include quarterly AND LTM if seasonality matters
- Add FCF metrics if cash generation is key story
- Include industry-specific metrics (Rule of 40 for SaaS, etc.)
- Add more statistics rows if you have >5 companies

---

## Section 10: Industry-Specific Additions (Optional)

Only add these if they're critical to your analysis. Most comps work fine with just core metrics.

**Software/SaaS:**
Add if relevant: ARR, Net Dollar Retention, Rule of 40

**Financial Services:**
Add if relevant: ROE, Net Interest Margin, Efficiency Ratio

**E-commerce:**
Add if relevant: GMV, Take Rate, Active Buyers

**Healthcare:**
Add if relevant: R&D/Revenue, Pipeline Value, Patent Timeline

**Manufacturing:**
Add if relevant: Asset Turnover, Inventory Turns, Backlog

---

## Section 11: Red Flags & Warning Signs

### Data Quality Issues
🚩 Inconsistent time periods (mixing quarterly and annual)  
🚩 Missing data without explanation  
🚩 Significant differences between data sources (>10% variance)

### Valuation Red Flags
🚩 Negative EBITDA companies being valued on EBITDA multiples (use revenue multiples instead)  
🚩 P/E ratios >100x without hypergrowth story  
🚩 Margins that don't make sense for the industry

### Comparability Issues
🚩 Different fiscal year ends (causes timing problems)  
🚩 Mixing pure-play and conglomerates  
🚩 Materially different business models labeled as "comps"

**When in doubt, exclude the company.** Better to have 3 perfect comps than 6 questionable ones.

---

## Section 12: Formulas Reference Guide

### Essential Excel Formulas
```excel
// Statistical Functions
=AVERAGE(range)          // Simple mean
=MEDIAN(range)           // Middle value
=QUARTILE(range, 1)      // 25th percentile
=QUARTILE(range, 3)      // 75th percentile
=MAX(range)              // Maximum value
=MIN(range)              // Minimum value
=STDEV.P(range)          // Standard deviation

// Financial Calculations
=B7/C7                                  // Simple ratio (Margin)
=SUM(B7:B9)/3                           // Average of multiple companies
=IF(B7<=0, "NM", IFERROR(C7/B7, "NM"))  // Guard zero/negative multiples with "NM" (NEVER 0!)
=IFERROR(C7/D7, "NM")                   // Handle divide by zero cleanly with "NM"

// Cross-Sheet References
='Sheet1'!B7             // Reference another sheet
=VLOOKUP(A7, Table1, 2)  // Lookup from data table
=INDEX(MATCH())          // Advanced lookup

// Formatting
=TEXT(B7, "0.0%")        // Format as percentage
=TEXT(C7, "#,##0")       // Thousands separator
```

### Common Ratio Formulas
```excel
Gross Margin = Gross Profit / Revenue
EBITDA Margin = EBITDA / Revenue
FCF Margin = Free Cash Flow / Revenue
FCF Conversion = FCF / Operating Cash Flow
ROE = Net Income / Shareholders' Equity
ROA = Net Income / Total Assets
Asset Turnover = Revenue / Total Assets
Debt/Equity = Total Debt / Shareholders' Equity
```

---

## Key Principles Summary

1. **Structure drives insight** - Right headers force right thinking
2. **Less is more** - 5-10 metrics that matter beat 20 that don't
3. **Choose metrics for your question** - Valuation analysis ≠ efficiency analysis
4. **Statistics show patterns** - Median/quartiles reveal more than average
5. **Transparency beats complexity** - Simple formulas everyone understands
6. **Comparability is king** - Better to exclude than force a bad comp
7. **Document your choices** - Explain which metrics and why in notes section

---

## Output Checklist

Before delivering a comp analysis, verify:
- [ ] All companies are truly comparable
- [ ] Data is from consistent time periods
- [ ] Units are clearly labeled (millions/billions)
- [ ] Formulas reference cells, not hardcoded values
- [ ] **All hard-coded input cells have comments with either: (1) exact data source with citation, OR (2) clear assumption with explanation**
- [ ] **Hyperlinks added where relevant** (SEC EDGAR filings, Bloomberg pages, research reports)
- [ ] Statistics include at least 5 metrics (Max, 75th, Med, 25th, Min)
- [ ] Notes section documents sources and methodology
- [ ] Visual formatting follows conventions (blue = input, black = formula)
- [ ] Sanity checks pass (margins logical, multiples reasonable)
- [ ] Date stamp is current ("As of [Date]")
- [ ] Formula auditing shows no errors (#DIV/0!, #REF!, #N/A)

---

## Continuous Improvement

After completing a comp analysis, ask:
1. Did the statistics reveal unexpected insights?
2. Were there any data gaps that limited analysis?
3. Did stakeholders ask for metrics you didn't include?
4. How long did it take vs. how long should it take?
5. What would make this more useful next time?

The best comp analyses evolve with each iteration. Save templates, learn from feedback, and refine the structure based on what decision-makers actually use.
