<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/earnings-analysis/references/workflow.md), Apache-2.0; modified for Google Antigravity. -->

# Detailed Workflow for Earnings Updates

This document provides detailed step-by-step instructions for each phase of the earnings update process (`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY`).

## ⚠️ Timeliness Verification (Latest vs. Historical Quarter)

Training data may be outdated. Actively retrieve primary earnings materials from connected MCP servers, SEC EDGAR, or company IR.

- **When analyzing the "latest quarter":**
  1. Write down today's date.
  2. Search for the most recent earnings release and matching transcript.
  3. Flag and re-search if the release date is older than **3 months** (90 days).
- **When the user explicitly requests a historical quarter** (e.g., *"analyze Q2 FY2024"*): Do **not** reject it as stale — retrieve that exact quarter's primary filings, match its pre-release consensus (or prior guidance/model fallback), and label the note clearly as a historical retrospective.

## Phase 1: Earnings Data Collection (30-60 minutes)

### Step 1: Identify the Latest Earnings Period

**CRITICAL**: ALWAYS SEARCH FOR THE LATEST EARNINGS - DO NOT RELY ON KNOWLEDGE CUTOFF.
**CRITICAL**: NEVER USE EARNINGS DATA FROM TRAINING - IT IS OUTDATED.

**Step 1a: Search for Latest Earnings Release**

**🚨 ACTIVELY SEARCH - training data is outdated. 🚨**

**MANDATORY STEP 1: CHECK TODAY'S DATE**
- **Write down today's date explicitly**: [Month] [Day], [Year]
- **Use this to verify** that any earnings found are within 3 months
- **Example**: "Today is October 29, 2024"

**MANDATORY STEP 2: SEARCH FOR "LATEST EARNINGS"**
- **Use web search** with queries like:
  - `[Company name] latest earnings results`
  - `[Company name] most recent quarterly earnings`
  - `[Ticker symbol] earnings latest quarter`
- **OR search company investor relations site**:
  - Go to `investor.[company].com` or `[company].com/investors`
  - Navigate to "Press Releases", "News", or "Earnings" section
  - **Sort by date to find MOST RECENT release**
  - Look for keywords: "earnings", "results", "financial results", "quarterly results"

**MANDATORY STEP 3: VERIFY THE RELEASE DATE**
- **Look at the date of the earnings release found**
- **Calculate**: Is this date within the last 3 months from today?
- **If YES** → Proceed to next step
- **If NO (older than 3 months)** → 🚨 WRONG QUARTER - Search again for more recent

**❌ COMMON MISTAKES TO AVOID:**
- ❌ Using earnings data from training without searching
- ❌ Assuming "Q3 2024" is latest based on expectations
- ❌ Grabbing the first earnings release found without checking the date
- ❌ Not comparing the release date to today's date
- ❌ Proceeding when the release is 4+ months old

**✅ CORRECT APPROACH:**
- ✅ Check today's date first
- ✅ Search explicitly for "latest" or "most recent"
- ✅ Read the actual release date on the materials
- ✅ Confirm release date is within 3 months of today
- ✅ If unsure, search again with different terms

**MANDATORY STEP 4: IDENTIFY THE QUARTER**
- **Read the title/headline** to identify the quarter (Q1, Q2, Q3, Q4 or fiscal quarter)
- **Read the release date** on the document itself
- **Verify both the quarter name AND the date are recent**

3. **Alternative search methods if IR site is unclear:**
   - Web search: `[Company name] latest earnings results`
   - Web search: `[Company name] most recent quarterly earnings`
   - Web search: `[Ticker symbol] earnings latest quarter`
   - SEC EDGAR: Search for company and look at most recent 10-Q or 10-K filing date

**Example searches that find latest data:**
- "Nike latest earnings results" → Returns most recent quarter reported
- "AAPL most recent quarterly earnings" → Shows latest Apple earnings
- "Tesla Q3 2024 earnings" → Results confirm Q3 2024 exists

**Step 1b: Understand Company's Fiscal Calendar**

After identifying the latest quarter from search, understand the company's fiscal year to interpret it correctly:

**Common fiscal year patterns:**
- **Calendar year (CY)**: Q1=Jan-Mar, Q2=Apr-Jun, Q3=Jul-Sep, Q4=Oct-Dec
- **Nike fiscal**: Q1=Jun-Aug, Q2=Sep-Nov, Q3=Dec-Feb, Q4=Mar-May (May fiscal year-end)
- **Apple fiscal**: Q1=Oct-Dec, Q2=Jan-Mar, Q3=Apr-Jun, Q4=Jul-Sep (September fiscal year-end)
- **Walmart fiscal**: Q1=Feb-Apr, Q2=May-Jul, Q3=Aug-Oct, Q4=Nov-Jan (January fiscal year-end)

Many companies state their fiscal year in the earnings release header. Search `[company] fiscal year calendar` if needed.

**Step 1c: MANDATORY VERIFICATION — Confirm Quarter & Dates**

- [ ] ✅ **Today's date written down**: `[YYYY-MM-DD]`
- [ ] ✅ **Target scope confirmed**: Latest quarter OR user-specified historical quarter
- [ ] ✅ **Earnings release date found**: `[YYYY-MM-DD]` (for "latest quarter" requests, verify within 3 months of today; skip staleness rejection if a historical quarter was explicitly requested)
- [ ] ✅ **Opened and read** the actual primary earnings release / 8-K / 10-Q confirming reported figures

**RED FLAGS (for "latest quarter" requests):** Release date > 90 days old without user requesting a historical period, mismatched release vs. transcript quarter, or relying on memory instead of primary filings.

**Step 1d: Handle Naming Variations**

Companies use different terminology - recognize these patterns:

**Quarter terminology:**
- "Q1 2024", "Q1 FY24", "First Quarter 2024", "1Q24"
- "Third Quarter Fiscal 2024", "Q3 FY2024", "3Q FY24"

**Earnings release titles:**
- "[Company] Reports Q3 2024 Results"
- "[Company] Announces Third Quarter Fiscal 2024 Financial Results"
- "[Company] Q3 Revenue Grew 15% Year-over-Year"

**SEC filing searches:**
- Company name may differ from common name (e.g., "Meta Platforms, Inc." vs "Facebook")
- Search by ticker symbol to find filings reliably
- Look for most recent 10-Q (quarterly) or 10-K (annual if Q4)

### Step 2: Gather Earnings Materials

After SEARCHING FOR and confirming the latest quarter, collect the following:

**⚠️ IMPORTANT: SEARCH for and ACCESS actual documents - do not rely on training data.**

**Primary Materials (REQUIRED):**
- **Earnings press release** - Usually on company investor relations site under "Press Releases" or "News"
  - Navigate to IR site and find the actual press release
  - Search patterns: "[Company name] latest earnings", "[Company name] Q[X] [Year] earnings results"
  - Look for PDF or HTML version
  - **Verify the date matches what was found in Step 1** (should be within last 1-3 months)
  - **Read the actual document** to confirm the quarter and get reported numbers

- **10-Q or 10-K filing** — On SEC EDGAR (`https://www.sec.gov/edgar/browse/?CIK=<CIK>`):
  - Search by ticker symbol or CIK
  - For quarters 1–3: Look for the corresponding Form `10-Q`
  - For Q4: Look for Form `10-K` (annual report)
  - Note: SEC Form `10-Q` is due within **40 days** after quarter end for large accelerated and accelerated filers (**45 days** for non-accelerated filers), though many issuers file alongside or within days of the earnings release
  - Direct EDGAR URL format: `https://www.sec.gov/edgar/browse/?CIK=<CIK>` or `https://www.sec.gov/Archives/edgar/data/<CIK>/<accession-no-hyphens>/<filename>.htm`

- **Earnings call transcript** — Verify the transcript date matches the earnings release date (±1 day) and same fiscal quarter (sources: company IR site, Seeking Alpha, AlphaStreet).

**Supplemental Materials (if available):**
- Investor presentation / earnings slides (PDF on IR site)
- Supplemental financial data schedules

**Reference Materials & Consensus Source Hierarchy (MANDATORY):**
- **Prior quarter & prior-year same quarter results** (for QoQ and YoY comparisons)
- **Prior guidance & internal model estimates** (from previous quarter's release and workspace model)
- **Pre-Print Consensus Estimates (STRICT HIERARCHY):**
  1. Connected financial MCP server (FactSet, S&P Capital IQ, LSEG, Bloomberg)
  2. Pre-release wire articles (Reuters, Bloomberg, WSJ) quoting a named consensus provider dated **BEFORE** the earnings release
  3. Company IR-compiled consensus (if published on IR site)
  - Always record **provider, as-of date, estimate count (`n`), and accounting basis (GAAP vs. Non-GAAP)**.
  - **NEVER** use post-print Yahoo Finance consensus as pre-print consensus (post-print pages overwrite pre-release estimates), and **NEVER** fabricate or cite unverified "whisper numbers".
  - **Fallback:** If verified pre-print consensus is unavailable, benchmark reported results against **prior management guidance** or **prior internal model estimates**, and explicitly mark Street consensus **`[UNSOURCED]`**.
- **Basis & In-Line Rules:** Match GAAP vs. Non-GAAP basis identically between reported actuals and consensus. Treat Revenue within **`±1.0%`** and EPS within **`±2.0%` or `±$0.01`** as **`INLINE`** (above = `BEAT`, below = `MISS`).

**VERIFICATION before proceeding to Step 3:**
- [ ] All primary materials (release, transcript, 10-Q/10-K if filed) verified for the **exact same quarter**
- [ ] Release date verified (`<= 90 days` for "latest quarter" requests; or labeled historical retrospective if explicitly requested)
- [ ] Pre-print consensus verified via hierarchy with provider + as-of date + `n` + GAAP/Non-GAAP basis (or `[UNSOURCED]` fallback documented)

### Step 3: Extract Key Metrics

Create a structured summary:

```
REPORTED RESULTS vs. ESTIMATES:
─────────────────────────────────────────────────
                    Reported    Our Est    Consensus    Beat/(Miss)
Revenue             $X,XXX      $X,XXX     $X,XXX       $XX (X%)
Gross Margin        XX.X%       XX.X%      XX.X%        XXbps
EBITDA              $XXX        $XXX       $XXX         $XX (X%)
Operating Profit    $XXX        $XXX       $XXX         $XX (X%)
EPS (Adjusted)      $X.XX       $X.XX      $X.XX        $X.XX
EPS (GAAP)          $X.XX       $X.XX      $X.XX        $X.XX

KEY BUSINESS METRICS:
─────────────────────────────────────────────────
[Metric 1]          XXX         XXX        XXX          +X% YoY
[Metric 2]          XXX         XXX        XXX          +X% YoY
[Metric 3]          XXX         XXX        XXX          +X% YoY
```

### Step 4: Identify Key Themes from Call

Listen to or read earnings call transcript and note:
- Management's tone (confident, cautious, defensive?)
- Key topics emphasized (product launches, geographic trends, competition)
- Questions from analysts (what are investors concerned about?)
- Guidance provided (raised, lowered, maintained, introduced?)
- Any surprises or unexpected commentary

## Phase 2: Analysis (2-3 hours)

### Step 5: Beat/Miss Analysis

For EACH key metric that beat or missed, explain:

**If BEAT:**
- What drove the outperformance?
- Was it one-time or sustainable?
- Did management guide higher going forward?
- How does this impact our thesis?

**If MISS:**
- What went wrong?
- Was it company-specific or industry-wide?
- Is management taking corrective action?
- How does this impact our thesis?

**Example Format:**
```
■ **Revenue Beat by 3% Driven by Strong DTC Performance**

Revenue of $13.5B exceeded our estimate of $13.1B by $400M (3%) and consensus
of $13.2B by $300M (2%). The outperformance was driven primarily by Direct-to-
Consumer channels, which grew 18% YoY (vs. our 12% estimate), offsetting
weaker-than-expected wholesale (-5% vs. flat estimate). Management cited strong
digital demand and successful product launches (Pegasus 40 running shoe, new
Jordan colorways) as key drivers. DTC now represents 42% of total revenue vs.
38% a year ago, demonstrating successful channel shift strategy.
```

### Step 6: Segment/Geographic/Product Analysis

Analyze performance by:
- Business segment (if multi-segment company)
- Geography (North America, Europe, China, etc.)
- Product category
- Channel (retail, wholesale, e-commerce)

Identify:
- What outperformed expectations?
- What underperformed?
- Trends vs. prior quarters
- Management commentary on outlook for each area

### Step 7: Margin Analysis

Analyze profitability:
- Gross margin: up or down? why?
- Operating margin: up or down? why?
- Key drivers (pricing, mix, costs, leverage)
- Outlook going forward

### Step 8: Guidance Analysis

If company provided guidance:
- Compare new guidance to prior guidance
- Compare to internal estimates and Street estimates
- Assess credibility (does company have track record of sandbagging? beating?)
- Identify key assumptions behind guidance

If company did NOT provide guidance:
- Note this explicitly
- Provide independent outlook based on results and commentary

### Step 9: Update Financial Model

Update estimates for:
- Current year (remaining quarters)
- Next year
- Potentially year after

**Show clearly:**
```
UPDATED ESTIMATES:
─────────────────────────────────────────────────
                        Old Est     New Est     Change      Reason
FY2024E Revenue         $XX.XB      $XX.XB      +X.X%      [Brief reason]
FY2024E EBITDA          $X.XB       $X.XB       +X.X%      [Brief reason]
FY2024E EPS             $X.XX       $X.XX       +X.X%      [Brief reason]

FY2025E Revenue         $XX.XB      $XX.XB      +X.X%      [Brief reason]
FY2025E EBITDA          $X.XB       $X.XB       +X.X%      [Brief reason]
FY2025E EPS             $X.XX       $X.XX       +X.X%      [Brief reason]
```

### Step 10: Update Valuation & Price Target

Based on updated estimates:
- Recalculate DCF (use updated cash flows)
- Update comparable company multiples (if peer group has reported)
- Determine new fair value
- Decide if price target changes

**Price Target Decision:**
- If estimates changed significantly (>5%) → Usually change price target
- If estimates changed marginally (<5%) → May maintain price target
- If thesis strengthened/weakened → May change even without estimate change

### Step 11: Assess Rating Impact

Decide whether to change rating:
- If results significantly better than expected + guidance raised → Consider upgrade
- If results significantly worse + guidance cut → Consider downgrade
- If inline or mixed → Usually maintain rating

**Consider:**
- Stock reaction (up/down/flat?)
- Valuation (expensive/cheap relative to new estimates?)
- Risk/reward (asymmetry shifted?)

## Phase 3: Chart Generation (1-2 hours)

### Step 12: Generate 8-12 Charts

Create charts focusing on QUARTERLY TRENDS and WHAT'S NEW.

**REQUIRED CHARTS (8-12 total):**

1. **Quarterly Revenue Progression** (Bar chart)
   - Last 8-12 quarters
   - Show beat/miss vs. estimates each quarter
   - Highlight current quarter

2. **Quarterly EPS Progression** (Bar chart)
   - Last 8-12 quarters
   - Show beat/miss vs. estimates
   - Adjusted and GAAP

3. **Quarterly Margin Trend** (Line chart)
   - Gross margin, EBIT margin, net margin
   - Last 8-12 quarters
   - Show trajectory

4. **Revenue by Segment/Geography** (Stacked bar OR table)
   - Current quarter vs. YoY
   - Growth rates by segment

5. **Key Operating Metrics** (Multi-line chart)
   - Customer count, ARPU, units sold, etc. (whatever is relevant)
   - Last 8-12 quarters

6. **Beat/Miss Summary** (Waterfall or table)
   - Show components of beat/miss
   - What drove variance from estimates

7. **Estimate Revision Chart** (Before/after comparison)
   - Old FY estimates vs. new FY estimates
   - Bar chart showing change

8. **Valuation Chart** (P/E or EV/EBITDA multiple)
   - Historical multiple range
   - Current multiple
   - Fair value multiple

**OPTIONAL CHARTS (if space allows):**
- Peer comparison (if peers have reported)
- Guidance vs. Street comparison
- Cash flow metrics
- Balance sheet highlights (if notable)

**Chart Style Guidelines:**
- Focus on TRENDS (quarterly progression)
- Highlight CHANGES (beat/miss, estimate revisions)
- Keep simple and clear (this is a fast-turnaround report)

## Phase 4: Report Creation (2-3 hours)

### Step 13: Create DOCX Report

Use `python-docx` to create the 8-12 page Word report (saved under `research/companies/<TICKER>/reports/`, or draft as Markdown under `research/companies/<TICKER>/notes/` if requested).

See [report-structure.md](report-structure.md) for complete page-by-page templates and formatting requirements.

**Key Steps:**
1. Create Page 1 with earnings summary and quick takeaways
2. Add detailed results analysis (Pages 2-3)
3. Include key metrics and guidance (Pages 4-5)
4. Update investment thesis (Pages 6-7)
5. Provide valuation and estimates (Pages 8-10)
6. Add appendix if needed (Pages 11-12)
7. Embed all 8-12 charts throughout
8. Add 1-3 summary tables
9. Include complete sources section with clickable hyperlinks

### Step 14: Optional - Update XLS Model

If a full financial model exists for this company (from initiation), update it with:
- Actual Q[X] results
- Revised estimates for future quarters
- Updated valuation

**Note**: For earnings updates, a full XLS file is OPTIONAL (not required like in initiation reports). The DOCX report is the primary deliverable.

If creating XLS, include:
- Quarterly model tab
- Updated annual projections
- Revised DCF
- Updated comps analysis

## Phase 5: Quality Check & Delivery (30 minutes)

### Step 15: Quality Checklist

Before staging the draft, verify:

**Content:**
- [ ] Beat/miss clearly stated and quantified
- [ ] Key drivers explained (not just "strong performance")
- [ ] Updated estimates provided (old vs. new shown)
- [ ] Price target updated or explicitly maintained
- [ ] Rating confirmed or changed with rationale
- [ ] Guidance analyzed (if provided)
- [ ] Thesis impact assessed

**Formatting:**
- [ ] Page 1 has summary box and key bullets
- [ ] All tables have source lines
- [ ] All figures numbered and captioned
- [ ] Estimates table shows old vs. new
- [ ] 8-12 charts embedded throughout
- [ ] Report is 8-12 pages (not too long, not too short)

**Accuracy:**
- [ ] Numbers match company's reported results exactly
- [ ] Math checks out (estimates, valuation)
- [ ] No typos in ticker, company name, numbers
- [ ] Charts match text descriptions
- [ ] Date is current

**Citations:** ⭐ MANDATORY
- [ ] Every figure has specific source with document and date
- [ ] Every table has specific source with document reference
- [ ] Beat/miss analysis cites consensus source with date
- [ ] Guidance changes cite current and prior guidance sources
- [ ] Key statistics have footnotes with specific page/slide references
- [ ] Sources section lists all materials with URLs
- [ ] ALL URLs are CLICKABLE HYPERLINKS (not plain text)
- [ ] Hyperlinks tested and working (Ctrl+Click opens correct page)
- [ ] All SEC filings hyperlinked to EDGAR viewer
- [ ] All earnings materials hyperlinked (release, transcript, presentation)
- [ ] Prior guidance hyperlinked to prior quarter's materials
- [ ] No raw URLs displayed - all formatted as clickable links
- [ ] Earnings call quotes cite specific speaker and approximate timestamp

**Timeliness & Compliance:**
- [ ] Draft completed within 24–48 hours of earnings release (for live coverage; or labeled historical retrospective if requested)
- [ ] Report labeled `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`
- [ ] Pre-print consensus verified via hierarchy with provider, as-of date, `n`, and matching GAAP/Non-GAAP basis (or `[UNSOURCED]` fallback documented)
- [ ] Beat/inline/miss classifications match thresholds (`±1.0%` revenue, `±2.0%` or `±$0.01` EPS)

### Step 16: Deliver Report

Provide user with:

1. **DOCX file**: `research/companies/<TICKER>/reports/<TICKER>_Q[Quarter]FY[Year]_Earnings_YYYY-MM-DD.docx`
2. **Chart files**: All PNG/JPG charts (for reference)
3. **Optional XLS**: Updated financial model if maintained

**Brief summary for user:**
```
[Company] Q[X] [Year] Earnings Update Complete

Results: [BEAT / INLINE / MISS]
- Revenue: $X.XB ([beat/missed] by $XXM or X%)
- EPS: $X.XX ([beat/missed] by $X.XX)

Key Takeaways:
■ [Takeaway 1]
■ [Takeaway 2]
■ [Takeaway 3]

Updated Estimates:
- FY[Year]E Revenue: $XX.XB (prior: $XX.XB, [+/-]X%)
- FY[Year]E EPS: $X.XX (prior: $X.XX, [+/-]X%)

Rating: [MAINTAINED / RAISED / LOWERED] [RATING]
Price Target: $XXX (prior: $XXX) - [+/-]XX% upside

Deliverable: 8-12 page earnings update report with updated estimates and valuation.
```
