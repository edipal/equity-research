<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/initiating-coverage/assets/quality-checklist.md), Apache-2.0; modified for Google Antigravity. -->

# Quality Control Checklist for Initiation Report Drafts

Verify all applicable items below before handing off a `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY` initiation report package.

> **No-Fabrication Rule:** Target counts below (pages, charts, tables, rows, words) are **target ranges where disclosed in primary filings**. NEVER fabricate segments, geographies, cohorts, precedent deals, or analyst credentials to hit a numerical target. Document any primary-filing disclosure gaps explicitly and tag analyst estimates `[ASSUMPTION]` and unverified items `[UNSOURCED]`.

---

## 1. Compliance, Draft Status & Integrity Checks (MANDATORY)

- [ ] Report prominently labeled **`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`** on Page 1 and in footer
- [ ] **Page 1 header reads `INITIATING COVERAGE`** (NOT `"Company Update"` or `"Investment Update"`)
- [ ] **No fabricated analyst identities:** Uses `[ANALYST NAME — REVIEW REQUIRED]` and `[DATE]` placeholders (no invented names, CFA/Ph.D. credentials, emails, phone numbers, or firm names)
- [ ] **Disclosures & Important Disclaimers appendix included** (draft notice, not investment advice, `[ASSUMPTION]`/`[UNSOURCED]` notice, analyst review sign-off block)
- [ ] **GAAP vs. Non-GAAP clearly separated** and reconciled wherever both appear
- [ ] **Timeliness verified:** Latest reported 10-K/10-Q quarter, fiscal year-end month, and share price as-of date explicitly stated
- [ ] **Balanced objectivity:** Bull, Base, and Bear cases + disconfirming evidence and risks honestly presented

---

## 2. Deliverables & File Paths

- [ ] DOCX draft report: `research/companies/<TICKER>/reports/<TICKER>_Initiation_Report_<Date>.docx`
- [ ] Excel model (Task 2 + Task 3 tabs): `research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx`
- [ ] Charts directory: `research/companies/<TICKER>/reports/charts/` (300 DPI PNGs + `chart_index.txt`)
- [ ] `scripts/recalc_xlsx.py` passed with exit `0` (`"status": "success"`, `0` formula errors, `0` uncached formulas)
- [ ] `.agents/skills/dcf-model/scripts/validate_dcf.py` passed with exit `0` (`"status": "PASS"`)

---

## 3. DOCX Report — Content, Length & Visual Density

**Target Length & Density (adjust proportionally to primary-filing complexity — never pad or fabricate):**
- [ ] Target length: ~30–50 pages (verified via LibreOffice PDF export + `pypdf` page count: `_____` pages)
- [ ] Target word count: ~10,000–15,000 words (`_____` words)
- [ ] Target charts: ~20–30+ embedded 300 DPI PNG images (`_____` charts; NO text-only chart placeholders)
- [ ] Target tables: ~12–20 structured tables (`_____` tables)
- [ ] Charts sized `<= 6.0"` wide (fitting within standard margins) and interspersed with text (no orphaned headers or chart-dump pages)

**Core Visuals (verify all that are supported by primary disclosures):**
- [ ] Figure 1: Stock price performance chart on Page 1
- [ ] Revenue by product/segment (stacked area) — *where multiple products/segments are disclosed in 10-K/10-Q*
- [ ] Revenue by geography (stacked bar) — *where geographic breakdown is disclosed in 10-K/10-Q*
- [ ] DCF sensitivity 2-way heatmap (formatted with float annotations, e.g., `$XX.XX`, centered on Base WACC & `g`)
- [ ] Valuation football field horizontal bar chart (showing current price & target price)

**Page 1 & Structure Checks:**
- [ ] Standardized rating scale (`Buy / Hold / Sell`) with explicit upside/downside thresholds
- [ ] Rating box: Rating, Price (as-of date), 12M Price Target, Upside/Downside `%`, 52-week range, Market Cap, EV
- [ ] 3–4 quantified `■` bullets with bold headers
- [ ] Financial summary table with `A` (actual) and `E` (estimate) suffixes matching the company's fiscal calendar
- [ ] Table of Contents (Page 2)
- [ ] Investment Thesis & Risks (8–12 distinct risks across 4 categories)
- [ ] Company 101 (overview, history, 3–4 executive bios from DEF 14A, products, GTM, industry, 5–10 peers where applicable)
- [ ] Financial Analysis, Projection Assumptions (product/geography/margin/OpEx/CapEx/NWC), and Bear/Base/Bull Scenarios
- [ ] Valuation Analysis (DCF, Comps with statistical summary, Precedents cross-check if applicable, Reconciliation, Catalysts & Risks)
- [ ] Appendices: Extended model tables, Data Sources & References, and Disclosures & Important Disclaimers

---

## 4. Citations & Hyperlinks

- [ ] Every figure and table has a specific source line (filing name + date, not generic `"Company data"`)
- [ ] All public URLs (SEC EDGAR, IR press releases, transcripts) are clickable hyperlinks in `.docx` (no bare raw URLs pasted in prose)
- [ ] Subscription sources (FactSet, Bloomberg, S&P Capital IQ) explicitly noted as `(subscription required)`
- [ ] Data Sources & References appendix lists all sources by category with dates

---

## 5. Excel Financial Model — Structure & Integrity

**Canonical Tab Architecture (9 core tabs + optional tabs):**
- [ ] Task 2 operating tabs: `Revenue Model`, `Income Statement`, `Cash Flow Statement`, `Balance Sheet`, `Scenarios`
- [ ] Task 3 valuation tabs: `WACC`, `DCF` *(inline 5×5 sensitivity table per `dcf-model`, or optional `Sensitivity` tab)*, `Comparable Companies`, `Valuation Summary`
- [ ] Optional tabs (only where disclosed/applicable): `Precedent Transactions` (cross-check)

**Model QA Checks:**
- [ ] `python3 scripts/recalc_xlsx.py <model.xlsx> 30` → `"status": "success"`, `0` errors, `0` uncached formulas
- [ ] `python3 .agents/skills/dcf-model/scripts/validate_dcf.py <model.xlsx>` → `"status": "PASS"`
- [ ] Balance Sheet balances (`Assets - (Liabilities + Equity) = 0`) in every historical and projected period
- [ ] Gross-debt WACC weights (`E/(E+D)` and `D/(E+D)`, `0%–100%`), mid-year explicit UFCF (`t - 0.5`) and mid-year perpetuity TV (`(1 + WACC)^(N - 0.5)`), strict end-of-year `(1 + WACC)^N` exit-multiple TV discounting, `g <= 3.5%` and `< WACC`
- [ ] Complete equity bridge (Gross Debt, Preferred, NCI, Cash, Equity Investments; no ASC 842 operating lease double-counting), TSM diluted shares, SBC kept as real expense
- [ ] Institutional styling per `.agents/rules/financial-modeling.md` (blue inputs with source comments, black formulas, green cross-sheet links, `#1F4E79` / `#D9E1F2` / `#F2F2F2` / `#BDD7EE` fill palette)

---

## 6. Cross-File Consistency & Writing Quality

- [ ] Spot-check 10–15 key metrics between `.docx` and `.xlsx` (Revenue, EBITDA, Margins, Diluted EPS, WACC, DCF Price, Comps Price, Price Target, Rating) — exact match
- [ ] Lead with numbers (`"Revenue grew 15% YoY to $1.2B"` vs. vague adjectives); use `"vs."` not `"versus"`
- [ ] Zero visible markdown syntax (`#`, `**`, backticks) in `.docx`; Calibri font, headers/footers, and page numbers verified in rendered PDF
- [ ] All disclosure omissions documented cleanly in `chart_index.txt` / model notes rather than fabricated
