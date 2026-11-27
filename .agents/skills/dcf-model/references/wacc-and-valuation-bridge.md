<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/dcf-model/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# WACC Methodology, Cash Flows, Terminal Value & Equity Bridge

Reference guide for the `dcf-model` skill. Every formula and convention here matches `.agents/skills/dcf-model/scripts/build_dcf_skeleton.py`.

---

## 1. Weighted Average Cost of Capital (WACC)

Build WACC on a dedicated `WACC` worksheet and link `DCF!C21 = WACC!C30` as the single source of truth (`WACC!C20 = DCF!C14` for Market Cap, `WACC!C21 = DCF!C15` for Total Gross Debt, and `WACC!C22 = DCF!C16` for Preferred Stock).

### 1.1 Cost of Equity (CAPM)

```text
Cost of Equity (Ke) = Risk-Free Rate (Rf) + Equity Beta (β) × Equity Risk Premium (ERP)
```

- **Risk-Free Rate (`Rf`)**: Current 10-Year US Treasury yield (matching currency and long-term cash flow horizon).
- **Equity Beta (`β`)**: 5-year monthly levered equity beta versus the benchmark index (e.g., S&P 500), or bottom-up unlevered peer median beta relevered to target capital structure:
  $$\beta_L = \beta_U \times \left[1 + (1 - t) \frac{D}{E}\right]$$
- **Equity Risk Premium (`ERP`)**: Institutional implied US ERP, typically `5.0%–6.0%` (cite Damodaran or institutional source + date).

### 1.2 After-Tax Cost of Debt

```text
After-Tax Cost of Debt = Pre-Tax Cost of Debt (Kd) × (1 - Marginal Tax Rate)
```

Determine Pre-Tax `Kd` in priority order:
1. Yield-to-maturity (YTM) on the company's traded senior unsecured bonds of comparable maturity.
2. Credit-rating spread over the 10-Year US Treasury (synthetic rating via EBIT / Interest Coverage if unrated).
3. Fallback: Weighted-average coupon / effective interest rate from the debt footnote (verify against current market rates; avoid stale zero-rate coupons).

Use the **marginal** corporate tax rate (typically `21%–28%` for US filers combining federal statutory and state taxes) rather than a distorted one-off GAAP effective tax rate.

### 1.3 Cost of Preferred Stock (`Kp`)

```text
Cost of Preferred Stock (Kp) = Annual Preferred Dividend / Market (or Liquidation) Value of Preferred Stock
```

- Preferred dividends are **not tax-deductible** (do NOT multiply `Kp` by `(1 - t)`).
- When the company has no preferred stock (`P = $0M`), set `Kp = 0.0%` (its weight is `0.00%`, so WACC collapses cleanly to Equity + Debt).

### 1.4 Capital Structure Weights — Gross Capitalization Basis (`E + D + P`, MANDATORY)

Always weight WACC using **Market Value of Equity (`E`)**, **Total Gross Debt (`D`)** (market value of debt preferred; book value of short-term + long-term gross debt + finance leases as fallback), and **Preferred Stock (`P`)** (market or liquidation value):

```text
Market Value of Equity (E) = Current Share Price × Fully Diluted Shares (TSM)
Total Gross Debt (D)       = Short-Term Debt + Current Portion of LT Debt + Long-Term Debt + Finance Leases  (D >= 0)
Preferred Stock (P)        = Liquidation / Market Value of Preferred Stock                                    (P >= 0)
Total Capitalization       = E + D + P

Equity Weight              = E / (E + D + P)    [bounded 0% to 100%]
Debt Weight                = D / (E + D + P)    [bounded 0% to 100%, NEVER negative]
Preferred Weight           = P / (E + D + P)    [bounded 0% to 100%, NEVER negative]
Weight Sum Check           = Equity Weight + Debt Weight + Preferred Weight = 100.00%

WACC                       = (Ke × Equity Weight) + (After-Tax Kd × Debt Weight) + (Kp × Preferred Weight)
```

**Why Gross Debt & Preferred (NOT Net Debt / EV), and Why NCI Is Excluded:**
- Cash is a non-operating asset added back in the Enterprise-to-Equity bridge — it is NOT "negative debt" in the capital structure.
- Using `Net Debt / EV` for a net-cash company creates a **negative debt weight** and an equity weight `> 100%`, mechanically distorting WACC above `Ke` (or below `Kd`).
- If the company has **zero gross debt and zero preferred stock** (`D = 0, P = 0`), then `Debt Weight = 0%`, `Preferred Weight = 0%`, `Equity Weight = 100%`, and `WACC = Ke`.
- **Non-Controlling Interest (NCI) convention:** Exclude NCI from parent WACC weights. Because GAAP consolidated financial statements include 100% of partially owned subsidiaries' operating cash flows, discount 100% consolidated UFCFs at the consolidated WACC (`E + D + P`), and deduct NCI at market/fair value in the Enterprise-to-Equity bridge (`DCF!C103`).

### 1.5 Typical WACC Sanity Ranges
- Large-cap, stable cash-generative business: `7.0%–9.0%`
- Mature growth / mid-cap: `9.0%–12.0%`
- High-growth, small-cap, or elevated business risk: `12.0%–15.0%`
- `validate_dcf.py` flags WACC `< 5.0%` or `> 20.0%` for review.

---

## 2. Unlevered Free Cash Flow (UFCF) & SBC Treatment

### 2.1 Standard UFCF Sequence

```text
  Revenue
- Cost of Revenue
= Gross Profit
- Operating Expenses (R&D, S&M, G&A — all modeled as % of Revenue, NEVER % of Gross Profit)
= EBIT (Operating Income)
- Cash Taxes on EBIT (EBIT × Marginal/Effective Tax Rate)
= NOPAT (Net Operating Profit After Tax)
+ Depreciation & Amortization (D&A)
- Capital Expenditures (CapEx)
- Change in Net Working Capital (ΔNWC; positive ΔNWC = cash outflow)
= Unlevered Free Cash Flow (UFCF)
```

**Working Capital Sign Convention:**
- `NWC = (Current Operating Assets ex-Cash) - (Current Operating Liabilities ex-Debt)`
- `ΔNWC_t = NWC_t - NWC_{t-1}` (or `(Rev_t - Rev_{t-1}) × NWC % of ΔRev`).
- Positive `ΔNWC` means working capital grew → **use of cash** (subtracted in UFCF).
- Negative `ΔNWC` (common in upfront-billed SaaS/deferred revenue businesses) → **source of cash** (adds to UFCF).

### 2.2 Stock-Based Compensation (SBC) Treatment (MANDATORY)

- **Default institutional standard:** Treat SBC as a **real economic expense** embedded in GAAP EBIT / NOPAT — **do NOT add back SBC to Unlevered FCF**. Adding back SBC inflates FCF without capturing the ongoing economic cost of equity dilution (or cash repurchases required to offset dilution).
- If an analyst explicitly requests a non-GAAP / SBC-adjusted FCF sensitivity:
  1. Present GAAP-EBIT-based UFCF as the primary valuation.
  2. Reconcile the SBC adjustment explicitly on a separate memo line, and reflect the full projected share-count dilution (or cash buyback offset) so value is not double-counted.

---

## 3. Discounting Conventions & Terminal Value

### 3.1 Explicit Forecast Discounting (Mid-Year Convention)

Assuming cash flows are generated evenly throughout each fiscal year, apply mid-year discounting to explicit projection years `t = 1, 2, ..., N`:

```text
Mid-Year Discount Period (t) = t - 0.5        (0.5, 1.5, 2.5, 3.5, 4.5 for a 5-year model)
Discount Factor (t)          = 1 / (1 + WACC)^(t - 0.5)
PV of UFCF (t)               = UFCF(t) × Discount Factor(t)
```

*(Note: If valuing mid-fiscal-year with a partial stub period of length `s`, discount Year 1 stub at `s / 2` and subsequent years at `s + 0.5`, `s + 1.5`, etc.)*

### 3.2 Terminal Value — Perpetuity Growth Method (Primary)

When explicit forecast cash flows use the mid-year convention (`t = 0.5, 1.5, ..., N - 0.5`), the first terminal cash flow (`UFCF_N × (1 + g)`) also occurs mid-year at `t = N + 0.5` — exactly 1.0 year after `t = N - 0.5`. Therefore, Gordon Growth values the perpetuity at `t = N - 0.5`, and discounting to `t = 0` uses **`(1 + WACC)^(N - 0.5)`** (mathematically equivalent to `TV / (1 + WACC)^N × (1 + WACC)^0.5`):

```text
Undiscounted TV (Perpetuity) = UFCF_N × (1 + g) / (WACC - g)
Terminal Discount Factor     = 1 / (1 + WACC)^(N - 0.5)   (consistent mid-year perpetuity timing)
PV of Terminal Value         = Undiscounted TV × Terminal Discount Factor
```

*(Note: Discounting a mid-year Perpetuity Growth TV at `N - 0.5` eliminates the half-year timing gap between Year `N` and Year `N + 1` and matches `build_dcf_skeleton.py` dynamically via `=1/(1+$C$21)^$G$87`. If an explicit model uses end-of-year cash flows `1..N`, then both explicit cash flows and Perpetuity TV use integer periods `1..N` and `N`.)*

**Terminal Growth Rate (`g`) Guardrails:**
- **Mathematical requirement:** Strictly `g < WACC` (`validate_dcf.py` fails otherwise).
- **Economic ceiling:** `g <= long-term nominal GDP growth / 10Y risk-free rate` — typically **`2.0%–3.5%`** across scenarios (e.g., Bear `2.0%`, Base `2.25%–2.5%`, Bull `2.5%–3.0%`).
- **Never** set perpetuity `g` to `3.5%–5.0%+`: a company growing faster than nominal GDP forever mathematically exceeds the global economy. If a company has a longer high-growth runway, extend the explicit projection horizon to 7–10 years rather than inflating `g`.

### 3.3 Terminal Value — Exit Multiple Method (Cross-Check)

```text
Undiscounted TV (Exit Multiple) = EBITDA_N × Exit EV/EBITDA Multiple
PV of Exit Multiple TV          = Undiscounted TV / (1 + WACC)^N
```

- **CRITICAL DISCOUNTING RULE:** An exit multiple (`EBITDA_N × Multiple`) values the business at the **end of Year `N`** (point-in-time exit transaction). It MUST ALWAYS be discounted by **`(1 + WACC)^N`** (`N = 5.0` in a 5-year model, e.g., `^($G$87+0.5)`) — **NEVER `(1 + WACC)^(N - 0.5)`**, even when explicit FCFs and Perpetuity Growth TV use the mid-year convention.
- Always display **Implied Exit EV / Terminal EBITDA** from the Perpetuity Growth TV (`= Undiscounted Perpetuity TV / EBITDA_N`) and cross-check against peer trading multiples from `comps-analysis`.

### 3.4 Terminal Value Sanity Checks
- **TV % of Enterprise Value:** Target **`50%–70%`** (`validate_dcf.py` warns if `< 50%` or `> 70%`).
- If `> 70%`, review terminal growth `g` or extend the explicit forecast horizon to 7–10 years; if `< 50%`, verify terminal assumptions are not overly conservative.

---

## 4. Enterprise-to-Equity Valuation Bridge & Diluted Share Count

### 4.1 Complete Valuation Bridge

```text
  Sum of PV of Explicit UFCFs (Years 1..N)
+ PV of Terminal Value
= Implied Enterprise Value (EV)
- Total Gross Debt (short-term + long-term + finance leases)
- Preferred Stock (liquidation preference / market value)
- Non-Controlling Interest (NCI / minority interest)
+ Cash & Marketable Securities (unrestricted cash + short-term investments)
+ Equity Investments (non-consolidated equity-method / marketable non-operating stakes)
= Implied Equity Value
÷ Fully Diluted Shares Outstanding (Treasury Stock Method + If-Converted)
= Implied Share Price
```

*(Include underfunded after-tax pension obligations or other debt-like claims if material and disclosed in the 10-K footnotes.)*

### 4.2 ASC 842 Operating Lease Double-Counting Warning

Under **ASC 842**, operating lease liabilities appear on the balance sheet, while operating lease expense remains embedded in GAAP EBIT (and thus already reduces Unlevered FCF):
- **Do NOT deduct ASC 842 operating lease liabilities from EV** when your Unlevered FCF already deducts operating lease expense (doing so double-counts the obligation).
- Only include operating lease liabilities in debt if you simultaneously remove operating lease rent from OpEx and treat lease amortization/interest consistently with IFRS 16 / finance leases.

### 4.3 Fully Diluted Shares Outstanding

Never use basic shares alone when dilutive securities exist:
1. **Separate RSUs & Treasury Stock Method (TSM) for options/warrants:**
   - **Unvested RSUs (`$0` strike price):** Always included 1:1 in diluted shares regardless of share price (NEVER lump RSUs into an option strike test, or they would be erroneously dropped whenever `P <= K`!).
   - **Stock options & warrants:**
     $$\text{Net Option Dilution} = \max\left(0,\; N_{\text{opt}} - \frac{N_{\text{opt}} \times K}{P}\right)$$
     where `K` is weighted-average option strike price and `P` is the reference/implied share price.
   - **Total Fully Diluted Shares:** `Basic Shares + Unvested RSUs + Net Option Dilution` (`=C9+C10+IF(C8>C12,C11-(C11*C12/C8),0)` in `build_dcf_skeleton.py`).
2. **If-Converted Method** for convertible debt or convertible preferred stock:
   - Test whether conversion price `<` implied/reference share price (in-the-money).
   - If in-the-money: add converted shares to diluted share count AND remove the convertible face value from Total Gross Debt / Preferred Stock in the equity bridge (avoid double-counting!).
   - If out-of-the-money: keep the instrument as debt/preferred in the equity bridge and exclude the shares.

---

## 5. Industry Variations

| Archetype | Forecast Horizon | Key Modeling Adjustments |
|---|---|---|
| **High-Growth Tech / SaaS** | 7–10 years | Fade revenue growth smoothly toward GDP; model unit economics (ARR, NDR, Rule of 40); keep SBC as real expense; track deferred revenue working-capital tailwind fading as growth slows. |
| **Mature / Stable** | 5 years | Tight Bear/Base/Bull spreads; CapEx ≈ D&A in terminal year; WACC typically `7%–9%`. |
| **Cyclical / Industrials** | 5–7 years (full cycle) | Normalize terminal-year margins and CapEx to mid-cycle levels — never extrapolate peak or trough margins into perpetuity. |
| **Multi-Segment / Conglomerate** | 5–7 years per unit | Value distinct segments via Sum-of-the-Parts (SOTP) with segment-specific WACC and exit multiples, then subtract capitalized corporate unallocated overhead. |
