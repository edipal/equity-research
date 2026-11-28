<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/initiating-coverage/references/task2-financial-modeling.md), Apache-2.0; modified for Google Antigravity. -->

# Task 2: Financial Modeling - Detailed Workflow

Step-by-step instructions for executing Task 2 (Financial Modeling) of the `initiating-coverage` skill. All deliverables are staged as `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY`.

## Task Overview

**Purpose:** Extract 3–5 years of historical financials (+ latest 10-Q LTM stub when mid-fiscal-year) and build an integrated, live-formula 5-year Excel financial model with Bear/Base/Bull scenarios.

**Prerequisites (verify before starting):**
- **Required:** Access to company financial data — latest `10-K` and recent `10-Q` filings from SEC EDGAR (public companies), verified financial statements (private companies), or pre-extracted user historicals.
- **Optional:** Task 1 (`<TICKER>_Research_Document_<Date>.md`) for business context.

**Canonical Task 2 Tabs (5 core tabs — Task 3 adds 4 valuation tabs to this same workbook):**
1. `Revenue Model` — Product/segment and geographic revenue builds mirroring primary filing disclosures
2. `Income Statement` — Full P&L (3–5 years historical + LTM stub if mid-year + 5 years projected)
3. `Cash Flow Statement` — Operating, Investing, Financing activities + Free Cash Flow
4. `Balance Sheet` — Full Assets, Liabilities, Equity + Balance Check (`= 0` every period)
5. `Scenarios` — Single workbook-wide `Case` selector (`1=Bear, 2=Base, 3=Bull`, named range `Case`, `1..3` `DataValidation`), Bear / Base / Bull assumption drivers, and summary outputs

*(All assumption inputs are placed directly as blue hardcoded cells with source comments on their respective tabs — do NOT assume a separate `Assumptions` tab unless explicitly requested. `Scenarios` holds the single workbook `Case` selector referenced across Tasks 2 and 3; Unlevered FCF lines and balance-sheet bridge items flow via live cross-sheet formulas directly from `Income Statement`, `Cash Flow Statement`, and `Balance Sheet` into Task 3's `DCF` and `WACC` tabs — never duplicated as standalone driver blocks.)*

> **No-Fabrication & Tagging Rule:** Match the granularity disclosed in primary 10-K/10-Q filings. Target ranges below (e.g., ~20–30 product rows, ~15–20 geography rows, ~40–50 P&L rows) apply **where disclosed in primary filings** — never invent sub-products, countries, or headcount splits that are not reported. Tag every forward estimate `[ASSUMPTION]` and any unverified historical/private-company input `[UNSOURCED]`.

---

## Model Structure, Styling & Formula Standards

Follow `.agents/rules/financial-modeling.md` and `xlsx-author`:
- **Font colors:** Blue (`#0000FF`) = hardcoded inputs; Black (`#000000`) = same-sheet formulas; Green (`#008000`) = cross-sheet links.
- **Fill palette:** `#1F4E79` dark navy section headers (white bold text), `#D9E1F2` column headers (black bold text), `#F2F2F2` input cell shading, `#BDD7EE` key total / check highlights.
- **Live formulas only:** Every total, subtotal, margin, growth rate, cash flow line, balance sheet roll-forward, and balance check MUST be a live Excel formula written via `openpyxl`.
- **Cell comments on creation:** Attach `"Source: [Filing/Document], [Date], [Note/Page]"` (or `[ASSUMPTION]`) to every blue input cell as it is written.
- **Dynamic fiscal years:** Verify fiscal year-end month and latest reported quarter from SEC filings before creating columns. Use `[FY-3]A`, `[FY-2]A`, `[FY-1]A`, `[FY]A` (plus `LTM` / `YTD` stub column when mid-fiscal-year) and `[FY+1]E` through `[FY+5]E`.

---

## Step-by-Step Modeling Workflow

### Step 1: Extract Historical Financials (Skip if User Provided Clean Historicals)

**For Public Companies:**
1. Pull latest `10-K` and any subsequent `10-Q` filings via connected MCP servers or SEC EDGAR (`https://www.sec.gov/edgar/browse/?CIK=<CIK>` / `https://www.sec.gov/edgar/searchedgar/companysearch`).
2. Optionally save raw extracted statements to `research/companies/<TICKER>/models/<TICKER>_Historical_Financials_<Date>.xlsx` with 3–5 years of historicals across:
   - Income Statement (Revenue, COGS, Gross Profit, R&D, S&M, G&A, D&A, Operating Income/EBIT, Interest, Other, Pre-Tax Income, Taxes, Net Income, Basic & Diluted Shares, Basic & Diluted EPS, plus GAAP-to-Non-GAAP reconciliations if management reports Non-GAAP metrics)
   - Cash Flow Statement (Net Income, D&A, SBC, Working Capital changes, CFO, CapEx, Acquisitions, CFI, Debt issuance/repayment, Share repurchases, Dividends, CFF, Beginning & Ending Cash)
   - Balance Sheet (Cash & Marketable Securities, AR, Inventory, Prepaids, PP&E, Intangibles, Goodwill, AP, Accrued Expenses, Deferred Revenue, Short-Term & Long-Term Debt, Finance Leases, Common Stock/APIC, Retained Earnings, Treasury Stock, AOCI)
   - Segment / Product / Geographic revenue footnote tables and filing notes (FYE month, filing date, one-time items).
3. If the company is mid-fiscal-year (e.g., Q1–Q3 reported), build an **LTM stub** (`Latest FY + Current YTD - Prior-Year YTD`) so projections anchor on the latest actual run-rate.

**For Private Companies:**
- Collect disclosed revenue, funding announcements, and audited/management statements where available.
- Never silently fabricate historical margins from peers — if peer medians are used as illustrative benchmarks for missing private-company lines, mark every inferred cell `[ASSUMPTION]` and every unverified historical baseline `[UNSOURCED]` with explicit source notes.

### Step 2: Build `Revenue Model` Tab

Mirror the company's actual 10-K segment footnote structure (ASC 280):
- **A. Revenue by Product / Segment** (target ~10–30 rows where disclosed):
  - List each reported product line / segment (historical actuals `[FY-3]A`–`[FY]A` + projections `[FY+1]E`–`[FY+5]E`), `% of Total Revenue`, and `YoY Growth %`.
  - Place projection growth-rate drivers in blue input rows directly above or below each product line (all tagged `[ASSUMPTION]` with rationale in cell comments).
- **B. Revenue by Geography** (target ~8–20 rows where disclosed):
  - List each reported geographic region (`North America`, `EMEA`, `APAC`, etc.) with `% of Total` and `YoY Growth %`.
  - Verify `Revenue by Product Total = Revenue by Geography Total = Total Revenue` (add an explicit reconciliation check row `= 0`).
- **C. Revenue by Channel / KPIs** *(only if disclosed in filings, e.g., Direct vs. Partner, Store Count × Sales/Store, or ARR / Seat Count × ARPU)*.

### Step 3: Build `Income Statement` Tab

Build a full GAAP P&L (target ~35–50 line items matching 10-K structure) linking Total Revenue directly from `Revenue Model` (green cross-sheet link):
- Revenue & `YoY Growth %`
- Cost of Revenue (linked to Gross Margin / COGS `% of Revenue` blue driver rows) → Gross Profit & `Gross Margin %`
- Operating Expenses modeled as `% of Revenue` (NEVER `% of Gross Profit`): R&D, S&M, G&A (and D&A clearly identified — note explicitly whether D&A and SBC are embedded inside COGS/OpEx per GAAP filings so neither is double-counted!)
- EBIT (Operating Income) & `EBIT Margin %`
- Memo: EBITDA (`= EBIT + Total D&A`) & `EBITDA Margin %` (keep GAAP vs. Non-GAAP / SBC-adjusted EBITDA explicitly separated and labeled!)
- Interest Expense (linked to beginning-of-period debt balance to avoid circularity), Interest Income, Other Income/(Expense)
- Pre-Tax Income (EBT), Income Tax Expense & Effective Tax Rate `%`
- Net Income & `Net Margin %`
- Basic & Diluted Shares Outstanding (incorporating disclosed buybacks/dilution), Basic & Diluted EPS

### Step 4: Build `Cash Flow Statement` Tab

Link Net Income from `Income Statement` (green link) and working-capital/PP&E/debt schedules from `Balance Sheet`:
- **Operating Activities:** Net Income + D&A + Stock-Based Compensation (ties to Balance Sheet APIC increase!) + Deferred Taxes / Other Non-Cash + Working Capital Changes (`-ΔAR`, `-ΔInventory`, `-ΔPrepaids`, `+ΔAP`, `+ΔAccrued`, `+ΔDeferred Revenue`) = **Cash from Operations (CFO)**
- **Investing Activities:** Capital Expenditures (`% of Revenue` driver), Acquisitions, Marketable Securities / Other = **Cash from Investing (CFI)**
- **Memo: Levered Free Cash Flow** (`= CFO - CapEx`) & `FCF Margin %`
- **Financing Activities:** Debt Issuance, Debt Repayment, Share Repurchases, Dividends Paid, Equity Issuance / Other = **Cash from Financing (CFF)**
- **Cash Roll-Forward:** Beginning Cash + Net Change in Cash = Ending Cash (links directly to `Balance Sheet` Cash).

### Step 5: Build `Balance Sheet` Tab

Build an integrated Balance Sheet (target ~30–45 line items matching 10-K structure):
- **Current Assets:** Cash & Equivalents (`= Cash Flow Statement` Ending Cash), Short-Term Investments, Accounts Receivable (DSO driver), Inventory (DIO driver), Prepaids & Other
- **Non-Current Assets:** Gross PP&E + CapEx, Accumulated Depreciation (`-D&A`), Net PP&E, Goodwill & Intangibles, Other Non-Current Assets
- **Current Liabilities:** Accounts Payable (DPO driver), Accrued Expenses, Deferred Revenue, Short-Term / Current Debt, Other Current Liabilities
- **Non-Current Liabilities:** Long-Term Debt (beginning balance + issuances - repayments), Deferred Tax Liabilities, Other Non-Current Liabilities
- **Shareholders' Equity:** Common Stock & APIC (prior APIC + SBC + equity issuance), Retained Earnings (prior RE + Net Income - Dividends), Treasury Stock (prior Treasury - Share Repurchases), AOCI
- **Balance Check Row (MANDATORY):** `= Total Assets - (Total Liabilities + Total Equity)` — MUST equal `0.00` in every historical and projected period.

### Step 6: Build `Scenarios` Tab

Create the **single workbook-wide scenario control** and structured **Bear / Base / Bull** parameter and output comparisons that drive Task 2's active projections AND are referenced via live cross-sheet link by Task 3's `DCF` tab (ensuring there is never a second independent scenario switch or duplicated operating driver block in Task 3):
- **Single `Case` selector:** Place `Case` (`1=Bear, 2=Base, 3=Bull`, guarded by whole-number `1..3` Excel `DataValidation`, workbook defined name `Case`) with an active-label formula (`=CHOOSE(INT(Case),"Bear Case","Base Case","Bull Case")`) driving active model assumptions across `Revenue Model`, `Income Statement`, `Cash Flow Statement`, and `Balance Sheet`.
- Key driver comparison table: `[FY+1]E–[FY+5]E` Revenue CAGR, Terminal Gross Margin, Terminal EBITDA Margin, CapEx `% of Revenue`, NWC `% of ΔRev`.
- Key output comparison table: Terminal-Year Revenue, EBITDA, EBITDA Margin, Net Income, Diluted EPS, Free Cash Flow, Cumulative 5Y FCF.
- Concise written scenario rationale (bull catalysts, base trajectory, bear downside triggers — all tagged `[ASSUMPTION]`).

### Step 7: Recalculate & Audit Workbook (MANDATORY QA GATE)

Before handing off Task 2, run the LibreOffice recalculation & formula audit gate so cached formula values are populated directly into the canonical versioned workbook on disk:

```bash
python3 scripts/recalc_xlsx.py research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx 30 --in-place
```
*(Or if staging is needed, write the uncached `openpyxl` output to `/tmp/<TICKER>_Financial_Model_<Date>_staged.xlsx`, run `--output` directly to `research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx`, and delete the temp file; default mode without flags is non-destructive `--check-only` and does not populate cached values on disk.)*

- **Required result:** Exit code `0`, `"status": "success"`, `"total_errors": 0`, `"uncached_formulas": 0`.
- If `"status": "errors_found"` or `"error"` (exit `1`), fix every error and re-run.
- If `"status": "unverified"` (exit `2`, LibreOffice absent), install LibreOffice or disclose clearly that cached values are unverified before proceeding.
- Confirm Balance Sheet check `= 0` across every period and Revenue by Product `= Revenue by Geography = Income Statement Revenue`.

---

## Deliverable Path & Success Criteria

- **Primary deliverable:** `research/companies/<TICKER>/models/<TICKER>_Financial_Model_YYYY-MM-DD.xlsx`
- **Optional intermediate extraction:** `research/companies/<TICKER>/models/<TICKER>_Historical_Financials_YYYY-MM-DD.xlsx`

**Checklist before proceeding to Task 3:**
- [ ] All 5 canonical tabs present: `Revenue Model`, `Income Statement`, `Cash Flow Statement`, `Balance Sheet`, `Scenarios`
- [ ] 3–5 years of verified historical actuals (+ LTM stub if mid-year) + 5 years of live-formula projections
- [ ] Zero fabricated segment/geography lines; all estimates tagged `[ASSUMPTION]` with source comments on blue inputs
- [ ] Balance sheet balances (`0.00` difference) every year; cash ties between CF and BS
- [ ] `scripts/recalc_xlsx.py` passes with exit `0` (`"status": "success"`)
