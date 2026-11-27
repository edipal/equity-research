<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/initiating-coverage/assets/report-template.md), Apache-2.0; modified for Google Antigravity. -->

# Equity Research Initiation Report Template

This template defines the layout for assembling a comprehensive equity research initiation report draft (`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`).

**CRITICAL ASSEMBLY REQUIREMENTS:**
1. **Generate chart PNGs (target 20–30+ where disclosed)** using Python (`matplotlib` / `seaborn` / `Pillow`) BEFORE assembling the Word document.
2. **Use `python-docx`** to build the formatted `.docx` file programmatically (never deliver raw markdown in place of `.docx`).
3. **Embed actual 300 DPI PNGs** from `research/companies/<TICKER>/reports/charts/` sized to `<= 6.0"` width so they fit within default page margins cleanly.
4. **Draft Banner & Placeholders:** Include `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)` in the Page 1 header and document footer. Use `[ANALYST NAME — REVIEW REQUIRED]` and `[DATE]` — **never** invent analyst identities, credentials, emails, or phone numbers.
5. **Dynamic Fiscal Years:** Replace `[FY-3]A`, `[FY-2]A`, `[FY-1]A`, `[FY]A` (or `LTM`), `[FY+1]E`..`[FY+5]E` with the company's actual fiscal calendar verified from its latest 10-K and 10-Q.

---

## PAGE 1: INITIATING COVERAGE (MOST IMPORTANT PAGE)

**CRITICAL:** Page 1 is an institutional **INITIATING COVERAGE** tearsheet — NOT a generic "Executive Summary" and NOT a "Company Update".

**Structural Checklist:**
- Prominent draft notice banner: `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`
- Section banner: **`INITIATING COVERAGE`** (NEVER `"Company Update"`)
- Top-left rating box using standardized **`Buy / Hold / Sell`** scale
- Analyst review placeholder block (no invented credentials or contact info)
- Top-right Figure 1 stock price performance chart
- 3–4 quantified `■` bullets (**bold topic header** + 3–5 concise sentences each)
- Bottom financial & valuation metrics table (`A` suffix for actuals, `E` suffix for estimates)

### Layout Structure

**HEADER BANNER:**
```text
DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)
INITIATING COVERAGE | [Company Name] ([EXCHANGE]: [TICKER]) | [DATE]
```

**TOP LEFT — RATING BOX:**
```text
Rating:             [Buy / Hold / Sell]
Price ([Date]):     $[XX.XX]
12M Price Target:   $[XX.XX] ([+/-XX.X%])
52-Week Range:      $[XX.XX] - $[XX.XX]
Market Cap:         $[XX.X]B
Enterprise Value:   $[XX.X]B
```
*(Standardized rating thresholds: Buy `>= +15%`, Hold `-10%` to `+15%`, Sell `<= -10%` 12-month expected return.)*

**TOP LEFT — ANALYST REVIEW BLOCK:**
```text
Prepared for Review By: [ANALYST NAME — REVIEW REQUIRED]
Review Date:            [DATE]
Status:                 DRAFT — INTERNAL REVIEW ONLY
```
*(Never fabricate names, CFA/Ph.D. credentials, email addresses, or phone numbers.)*

**TOP RIGHT — STOCK PRICE PERFORMANCE:**
```text
Figure 1 - [Company Name] Stock Price Performance (12-24M vs. Benchmark)
[Embedded 300 DPI PNG chart]
Source: [Market Data Provider], as of [Date].
```

**MAIN CONTENT — SECTION BAR:**
```text
INITIATING COVERAGE — [BUY / HOLD / SELL] | [THESIS-FOCUSED HEADLINE]
```

**MAIN CONTENT — DETAILED BULLETS (3–4 bullets):**
```text
■ **[Bold Topic Header capturing key thesis pillar].** 3–5 sentences leading with quantified metrics, primary filing citations, and model projections. Use "vs." rather than "versus". Tag explicit forward estimates [ASSUMPTION] where appropriate in draft notes.

■ **[Second Topic Header].** [3–5 sentences...]

■ **[Third Topic Header].** [3–5 sentences...]

■ **[Fourth Topic Header — Risks / Valuation].** [3–5 sentences...]
```

**BOTTOM SECTION — FINANCIAL AND VALUATION METRICS TABLE:**
```text
                            [FY-2]A   [FY-1]A   [FY]A     [FY+1]E   [FY+2]E
Revenue ($M)                [X]       [X]       [X]       [X]       [X]
Revenue Growth (%)          X.X%      X.X%      X.X%      X.X%      X.X%
Gross Margin (%)            X.X%      X.X%      X.X%      X.X%      X.X%
EBITDA ($M)                 [X]       [X]       [X]       [X]       [X]
EBITDA Margin (%)           X.X%      X.X%      X.X%      X.X%      X.X%
Diluted EPS ($)             X.XX      X.XX      X.XX      X.XX      X.XX
P/E (x)                     XX.Xx     XX.Xx     XX.Xx     XX.Xx     XX.Xx
EV/Revenue (x)              X.Xx      X.Xx      X.Xx      X.Xx      X.Xx
EV/EBITDA (x)               XX.Xx     XX.Xx     XX.Xx     XX.Xx     XX.Xx

Note: "A" = actual per SEC filings; "E" = draft model estimates [ASSUMPTION]. State GAAP vs. Non-GAAP basis explicitly.
Source: Company SEC filings (FY[Current] 10-K / latest 10-Q), internal draft model.
```

---

## FIGURE NUMBERING AND FORMATTING STANDARDS

Every chart and table must include:
1. **Sequential numbering:** `Figure 1`, `Figure 2`, `Figure 3`, ... (no gaps).
2. **Descriptive title:** `Figure X - [Company] [Specific Metric] ([Period])`.
3. **Specific source line below:** `Source: [Company] FY[Year] 10-K (filed [Date]), [Market Data Source], internal draft estimates.` (Never generic `"Company data"` alone.)

---

## PAGE 2: TABLE OF CONTENTS

```text
Initiating Coverage Summary..........................................1
Investment Thesis & Risks............................................3
Company Overview (Company 101).......................................6
  Business Description & History.....................................6
  Management & Governance............................................8
  Products & Technology..............................................9
  Customers & Go-to-Market..........................................11
Industry Overview & Competitive Landscape...........................13
Financial Analysis, Projections & Scenarios.........................18
  Historical Performance............................................18
  Projection Assumptions............................................21
  Scenario Analysis (Bear / Base / Bull)............................26
Valuation Analysis..................................................31
Appendices, Data Sources & Disclosures..............................41
```

---

## PAGES 3–5: INVESTMENT THESIS & RISKS

- **Investment Thesis:** 3–5 pillars (quantified opportunity, competitive advantage, financial impact, catalyst timeline), interspersed with charts (e.g., TAM evolution, competitive matrix, margin pathway).
- **Investment Risks:** Target 8–12 distinct risks across 4 categories (Company-Specific, Industry/Market, Financial, Macroeconomic), noting severity, likelihood, and mitigants. Present disconfirming evidence honestly per `.agents/rules/research-integrity.md`.

---

## PAGES 6–17: COMPANY 101

Incorporate Task 1 prose verbatim (reformatted to DOCX styles) with inline charts:
1. **Company Overview** (plain-English description, business model, monetization, scale)
2. **Company History & Milestones** (founding story, timeline, strategic pivots, key acquisitions)
3. **Management & Governance** (3–4 executive bios sourced from DEF 14A/filings, board independence, insider ownership)
4. **Products & Technology** (core offerings, pricing model, differentiation, disclosed traction)
5. **Customers & Go-to-Market** (customer segments, channels, unit economics where disclosed)
6. **Industry Overview & Competitive Landscape** (industry scope, TAM/SAM/SOM, Porter's Five Forces, 5–10 peers where applicable, market share where disclosed)

---

## PAGES 18–30: FINANCIAL ANALYSIS, PROJECTIONS & SCENARIOS

- **Historical Performance (3–5 years + LTM stub if mid-year):** Full Income Statement, Revenue by Product/Segment (where disclosed in 10-K), Revenue by Geography (where disclosed in 10-K), Cash Flow Statement, Balance Sheet highlights.
- **Projection Assumptions (`[FY+1]E–[FY+5]E`):** Product/segment drivers, geographic drivers, gross margin bridge, R&D/S&M/G&A operating leverage (`% of Revenue`), D&A, CapEx, NWC (DSO/DIO/DPO), tax rate, SBC treatment. Tag all forward drivers `[ASSUMPTION]`.
- **Scenario Analysis (Bear / Base / Bull):** Quantified revenue CAGR, terminal EBITDA margin, FCF, key catalysts/triggers, probability weights, and comparison table extracted from the `Scenarios` tab.

---

## PAGES 31–40: VALUATION ANALYSIS

Aligned strictly with `dcf-model` and `references/task3-valuation.md`:
- **DCF Analysis:** Gross-debt WACC table (`WACC` tab), 5Y mid-year explicit UFCF (`t - 0.5`) and mid-year perpetuity TV (`(1 + WACC)^(N - 0.5)`, `g <= 3.5%`), cross-checked against strict end-of-year `(1 + WACC)^N` exit-multiple TV discounting, complete equity bridge (`EV - Gross Debt - Preferred - NCI + Cash + Equity Investments`), TSM diluted shares, SBC kept as real expense, centered odd 5×5 live sensitivity heatmap (`chart_28`), and TV `%` of EV check (`50%–70%` target).
- **Comparable Companies (`Comparable Companies` tab):** Two-part table (5–10 peers + target, followed by mandatory Max / 75th / Median / 25th / Min statistical summary) with NTM/LTM multiples and premium/discount justification.
- **Precedent Transactions (Optional Cross-Check):** Include only where relevant M&A exists; keep at `0%` target weight (cross-check only) unless an active M&A catalyst is part of the thesis.
- **Valuation Summary & Football Field (`Valuation Summary` tab + `chart_32`):** Weighted valuation summary, football field chart, 12-month Price Target, standardized `Buy / Hold / Sell` rating, 3–5 catalysts, and key risks to target.

---

## PAGES 41+: APPENDICES, DATA SOURCES & DISCLOSURES (MANDATORY)

1. **Appendix A: Detailed Financial Model Tables** (extended IS/CF/BS extracts reconciled to `<TICKER>_Financial_Model_<Date>.xlsx`)
2. **Appendix B: Data Sources & References** (organized by SEC Filings, Earnings Transcripts, Company IR, Market Data, Industry Reports — all public URLs formatted as clickable hyperlinks with dates; paid feeds noted as subscription required)
3. **Appendix C: Disclosures & Important Disclaimers (MANDATORY)**
   - **Draft Status Notice:** *"DRAFT — FOR INTERNAL ANALYST REVIEW ONLY. This document is an unapproved internal research draft generated with AI assistance and has not been reviewed, verified, or signed off by a licensed research analyst or supervisory analyst."*
   - **Not Investment Advice:** *"This draft does not constitute personalized investment advice, a solicitation, or an offer to buy or sell any security."*
   - **Data & Assumption Notice:** *"Historical figures are sourced from primary public filings cited herein. Forward projections represent illustrative draft assumptions (`[ASSUMPTION]`) subject to analyst verification. Any unverified figure is marked `[UNSOURCED]`."*
   - **Analyst Sign-Off Block:** `[ANALYST NAME — REVIEW REQUIRED]` | `[SUPERVISORY REVIEW — PENDING]` | `[DATE]`
