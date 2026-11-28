<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/3-statement-model/references/formulas.md), Apache-2.0; modified for Google Antigravity. -->

# Formula Reference

**IMPORTANT:** Use the formulas outlined in this reference document unless otherwise specified by the user.

---

## Core Linkages

```text
Balance Sheet:           Assets = Liabilities + Equity
Net Income:              IS Net Income -> CF Operations (starting point)
Cash Flow:               ΔCash = CFO + CFI + CFF
Cash Tie-Out:            Ending Cash (CF) = Cash (BS Asset)
Retained Earnings:       Prior RE + Net Income - Dividends = Ending RE
Common Stock / APIC:     Prior APIC + Equity Issuance (CFF) + SBC (IS) = Ending APIC
Treasury Stock:          Prior Treasury Stock - Share Repurchases (CFF) = Ending Treasury Stock
AOCI:                    Prior AOCI + Other Comprehensive Income (OCI) = Ending AOCI
Revolver / Min Cash:     Cash Before Revolver < Min Cash -> Revolver Draw; Excess Cash -> Revolver Paydown
(Startup/Monthly only):  Closing Cash (Monthly) = Closing Cash (Annual); Year 0 Equity = Beg Equity (Y1)
```

## Gross Profit Calculation

**IMPORTANT:** Gross Profit must be calculated from Net Revenue, not Gross Revenue.

```
Net Revenue - Cost of Revenue = Gross Profit
```

| Term | Definition |
|------|------------|
| Gross Revenue | Total revenue before any deductions |
| Net Revenue | Gross Revenue - Returns - Allowances - Discounts |
| Cost of Revenue | Direct costs attributable to production of goods/services sold |
| Gross Profit | Net Revenue - Cost of Revenue |

**Note:** Always use Net Revenue (also called "Net Sales" or simply "Revenue" on most financial statements) as the starting point for profitability calculations. Gross Revenue overstates the true top-line performance.

## Margin Formulas

```
Gross Margin %      = Gross Profit / Net Revenue
EBITDA              = EBIT + D&A  (or = Gross Profit - OpEx)
EBITDA Margin %     = EBITDA / Net Revenue
EBIT Margin %       = EBIT / Net Revenue
Net Income Margin % = Net Income / Net Revenue
```

## Credit Metric Formulas

```
Total Debt            = Current Portion of Debt + Long-Term Debt
Net Debt              = Total Debt - Cash
Total Debt / EBITDA   = Total Debt / EBITDA (from IS)
Net Debt / EBITDA     = Net Debt / EBITDA (from IS)
Interest Coverage     = EBITDA / Interest Expense (from IS)
Net Int Exp % Debt    = Net Interest Expense / Long-Term Debt
Debt / Total Cap      = Total Debt / (Total Debt + Total Equity)
Debt / Equity         = Total Debt / Total Equity
Current Ratio         = Total Current Assets / Total Current Liabilities
Quick Ratio           = (Total Current Assets - Inventory) / Total Current Liabilities
```

## Forecast Formulas (% of Net Revenue Method)

```
Cost of Revenue (Forecast) = Net Revenue × Cost of Revenue % Assumption
S&M (Forecast)             = Net Revenue × S&M % Assumption
G&A (Forecast)             = Net Revenue × G&A % Assumption
R&D (Forecast)             = Net Revenue × R&D % Assumption
SBC (Forecast)             = Net Revenue × SBC % Assumption
```

## Working Capital Formulas

```
Accounts Receivable
  Prior AR
  + Revenue (from IS)
  - Cash Collections (plug)
  = Ending AR
  DSO = (AR / Revenue) × 365

Inventory
  Prior Inventory
  + Purchases (plug)
  - COGS (from IS)
  = Ending Inventory
  DIO = (Inventory / COGS) × 365

Accounts Payable
  Prior AP
  + Purchases (from Inventory calc)
  - Cash Payments (plug)
  = Ending AP
  DPO = (AP / COGS) × 365

Net Working Capital = AR + Inventory - AP
ΔWC = Current NWC - Prior NWC
```

## D&A Schedule Formulas

```
Beginning PP&E (Gross)
+ CapEx
= Ending PP&E (Gross)

Beginning Accumulated Depreciation
+ Depreciation Expense
= Ending Accumulated Depreciation

PP&E (Net) = Gross PP&E - Accumulated Depreciation
```

## Debt & Revolver Sweep Schedule Formulas

```text
Term Debt:
  Beginning Term Debt
  + New Term Borrowings
  - Mandatory / Optional Term Repayments
  = Ending Term Debt

Revolver / Minimum Cash Sweep (prevents negative cash balances):
  Cash Available for Debt Service = Beginning Cash + CFO + CFI + Term Debt Flow + Equity/Div Flow
  Excess / (Shortfall) Cash       = Cash Available for Debt Service - Minimum Cash Balance
  Revolver Draw / (Paydown)       = -MIN(Beginning Revolver, Excess Cash)  [if Excess > 0, pay down up to Beg Revolver]
                                    or ABS(Shortfall)                       [if Shortfall < 0, draw to maintain Min Cash]
  Ending Revolver                 = Beginning Revolver + Revolver Draw / (Paydown)

Interest Expense (Default Non-Circular Method):
  Interest Expense = (Beginning Term Debt × Term Rate) + (Beginning Revolver × Revolver Rate)
  (Use beginning-of-period balances by default to avoid circularity; only use average balances if
   wb.calculation.iterate = True and a circuit-breaker toggle is enabled.)
```

## Equity Roll-Forward Formulas (RE, APIC, Treasury Stock, AOCI)

```text
Retained Earnings:
  Beginning Retained Earnings + Net Income (IS) - Dividends Paid (CFF) = Ending Retained Earnings

Common Stock / Additional Paid-In Capital (APIC):
  Beginning Common Stock / APIC + Equity Issuance (CFF) + Stock-Based Compensation (IS) = Ending APIC

Treasury Stock (Share Repurchases):
  Beginning Treasury Stock (negative equity) - Share Repurchases (CFF cash outflow) = Ending Treasury Stock

Accumulated Other Comprehensive Income (AOCI):
  Beginning AOCI + Other Comprehensive Income (FX / hedging / unrealized gains) = Ending AOCI
```

## NOL (Net Operating Loss) Schedule Formulas

```text
NOL CARRYFORWARD SCHEDULE

Beginning NOL Balance:
  - Established public company: pull Ending NOL carryforward from latest 10-K Income Tax footnote
  - New-formation / startup model only: Beginning NOL (Year 1) = 0
+ NOL Generated (if EBT < 0, then ABS(EBT), else 0)
- NOL Utilized (limited by taxable income and utilization cap)
= Ending NOL Balance

NOL UTILIZATION CALCULATION

Pre-Tax Income (EBT)
  If EBT > 0:
    NOL Available = Beginning NOL Balance
    Utilization Limit = EBT × 80%  (post-2017 federal limit)
    NOL Utilized = MIN(NOL Available, Utilization Limit)
    Taxable Income = EBT - NOL Utilized
  If EBT ≤ 0:
    NOL Utilized = 0
    Taxable Income = 0
    NOL Generated = ABS(EBT)

TAX CALCULATION WITH NOL (GAAP / ASC 740)

Current Tax Expense (Taxes Payable) = MAX(0, Taxable Income × Tax Rate)
  (Current cash taxes payable cannot be negative; losses create NOL carryforwards)

DEFERRED TAX ASSET (DTA) FOR NOL

DTA - NOL Carryforward = Ending NOL Balance × Tax Rate
  (If a 100% valuation allowance applies, Net DTA = 0 on BS and ΔDTA = 0 on IS & CF)
ΔDTA = Current DTA - Prior DTA
Deferred Tax Expense / (Benefit) = -ΔDTA
Total Income Tax Expense / (Benefit) = Current Tax Expense + Deferred Tax Expense / (Benefit)
                                     = Current Tax Expense - ΔDTA
  (Increase in DTA = non-cash tax benefit on IS; subtract ΔDTA from Net Income on CFO)
  (Decrease in DTA = non-cash deferred tax expense on IS; add back |ΔDTA| on CFO)
```

## Balance Sheet Structure

```text
ASSETS
  Cash (from CF ending cash; >= Minimum Cash via revolver sweep)
  Accounts Receivable (from WC)
  Inventory (from WC)
  Total Current Assets

  PP&E, Net (from DA)
  Deferred Tax Asset - NOL (from NOL schedule)
  Total Non-Current Assets
  Total Assets

LIABILITIES
  Accounts Payable (from WC)
  Revolver & Current Portion of Debt (from Debt & Revolver schedule)
  Total Current Liabilities

  Long-Term Debt (from Debt & Revolver schedule)
  Total Liabilities

EQUITY
  Common Stock / APIC (Prior + Equity Issuance + SBC)
  Retained Earnings (Prior RE + NI - Dividends)
  Treasury Stock (Prior Treasury - Share Repurchases)
  Accumulated OCI (Prior AOCI + OCI)
  Total Equity

CHECK: Assets - Liabilities - Equity = 0
```

## Cash Flow Statement Structure

```text
CASH FROM OPERATIONS (CFO)
  Net Income (LINK: IS)
  + D&A (LINK: DA schedule)
  + Stock-Based Compensation (SBC) (LINK: IS memo or carved-out line)
  - ΔDTA (Deferred Tax Asset) (LINK: NOL schedule; = Deferred Tax Expense/(Benefit))
  - ΔAR (LINK: WC)
  - ΔInventory (LINK: WC)
  + ΔAP (LINK: WC)
  = CFO

CASH FROM INVESTING (CFI)
  - CapEx (LINK: DA schedule)
  = CFI

CASH FROM FINANCING (CFF)
  + Term Debt Issuance (LINK: Debt)
  - Term Debt Repayment (LINK: Debt)
  + / (-) Revolver Draw / (Paydown) (LINK: Debt & Revolver sweep)
  + Equity Issuance (LINK: BS Common Stock/APIC)
  - Share Repurchases (LINK: BS Treasury Stock)
  - Dividends Paid (LINK: RE schedule)
  = CFF

Net Change in Cash = CFO + CFI + CFF (+ FX/OCI cash effect if applicable)
Beginning Cash + Net Change in Cash = Ending Cash (LINK TO: BS Cash)
```

## Income Statement Structure (Avoiding SBC & D&A Double-Counting)

> **CRITICAL ACCOUNTING RULE:** Never double-count D&A or SBC! Choose **ONE** presentation consistently across historicals and forecasts and label it explicitly on the sheet:
> - **Option A — GAAP Functional Presentation (Preferred for 10-K filings):** `COGS`, `S&M`, `G&A`, and `R&D` INCLUDE embedded D&A and SBC. Compute `EBIT = Gross Profit - S&M - G&A - R&D`, and display `Total D&A` and `Total SBC` strictly as **Memo rows** below EBIT (so `EBITDA = EBIT + Memo D&A`; never subtract D&A or SBC twice!).
> - **Option B — Carved-Out Presentation:** Carve D&A and SBC out of historical `COGS`/`S&M`/`G&A`/`R&D` first (labeling them `S&M (ex-D&A/SBC)`, etc.), THEN subtract separate `D&A` and `SBC` lines to arrive at `EBIT`.

```text
Net Revenue
  Growth %
(-) Cost of Revenue (state whether incl. or excl. D&A/SBC)
  % of Net Revenue
────────────────
Gross Profit (= Net Revenue - Cost of Revenue)
  Gross Margin %

(-) S&M (state whether incl. or excl. D&A/SBC)
(-) G&A (state whether incl. or excl. D&A/SBC)
(-) R&D (state whether incl. or excl. D&A/SBC)
[Only if Option B carved-out: (-) D&A and (-) SBC]
────────────────
EBIT (= Gross Profit - Total OpEx)
  EBIT Margin %
Memo: Total D&A (linked from PP&E/D&A schedule)
Memo: Total SBC (linked to CF add-back & BS APIC)
Memo: EBITDA (= EBIT + Total D&A; label GAAP vs. Non-GAAP/SBC basis explicitly)
  EBITDA Margin %

(-) Interest Expense (linked to Beginning Debt & Revolver balance)
(+) Interest Income
────────────────
EBT (Pre-Tax Income)
(-) Current Tax Expense (MAX(0, Taxable Income × Tax Rate), from NOL schedule)
(-) Deferred Tax Expense / (Benefit) (= -ΔDTA from NOL schedule; 0 if 100% valuation allowance)
────────────────
Total Income Tax Expense / (Benefit) (= Current Tax Expense - ΔDTA)
────────────────
Net Income (= EBT - Total Income Tax Expense)
  Net Income Margin %
```

## Check Formulas

```text
BS Balance Check:        = Assets - Liabilities - Equity                             (must = 0)
Cash Tie-Out:            = BS Cash - CF Ending Cash                                  (must = 0)
Min Cash Floor:          = BS Cash >= Minimum Cash Balance                           (must be TRUE)
RE Roll-Forward:         = Prior RE + NI - Dividends - BS Ending RE                  (must = 0)
Common Stock / APIC:     = ΔCommon Stock/APIC (BS) - Equity Issuance (CFF) - SBC     (must = 0)
Treasury Stock Tie-Out:  = ΔTreasury Stock (BS) + Share Repurchases (CFF)            (must = 0)
AOCI Tie-Out:            = Prior AOCI + OCI - BS Ending AOCI                         (must = 0)
DTA Tie-Out:             = NOL Schedule DTA - BS DTA                                 (must = 0)
NOL Utilization Cap:     = NOL Utilized <= EBT × 80%                                 (must be TRUE post-2017)
NOL Non-Negative:        = Ending NOL Balance >= 0                                   (must be TRUE)
NOL Accumulation:        = NOL increases only when EBT < 0                           (must be TRUE)
(Startup/Monthly only):  = Closing Cash (Monthly) - Closing Cash (Annual) = 0; Beg NOL (Y1) = 0
```
