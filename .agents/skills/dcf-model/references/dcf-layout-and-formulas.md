<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/dcf-model/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# Canonical DCF Workbook Layout, Formulas & Formatting

Reference specification for `.agents/skills/dcf-model/SKILL.md`. All row coordinates, sheet names, and formulas match `.agents/skills/dcf-model/scripts/build_dcf_skeleton.py`.

---

## 1. Two-Sheet Workbook Architecture

| Sheet | Role | Key Outputs |
|---|---|---|
| `DCF` | Main valuation sheet: named `Case` selector (`B4`) with `1..3` validation, market/bridge inputs (separate RSU and option rows), stacked Bear/Base/Bull assumption blocks, `INDEX`-driven Active Case consolidation block, 5Y UFCF build, Terminal Value & Equity Bridge, Exit-Multiple cross-check, and inline `SUMPRODUCT` 5×5 live-formula sensitivity table(s) at bottom | `DCF!C108` (Implied Share Price — Perpetuity), `DCF!C119` (Implied Share Price — Exit Multiple), `DCF!E127` (Sensitivity center cell = `DCF!C108`) |
| `WACC` | Single source of truth for CAPM Cost of Equity, After-Tax Cost of Debt, Cost of Preferred Stock, and Gross-Capitalization (`E + D + P`) Weights | `WACC!C30` (linked into `DCF!C21`) |

Generate a clean, verified starting workbook anytime via:

```bash
python3 .agents/skills/dcf-model/scripts/build_dcf_skeleton.py \
  --output research/companies/<TICKER>/models/<TICKER>_DCF_YYYY-MM-DD.xlsx \
  --case 2
```

---

## 2. `WACC` Sheet Row Map

| Rows | Section | Key Cells & Live Formulas |
|---|---|---|
| `1–2` | Title & units note | `$ millions` basis |
| `4–9` | **1. Cost of Equity (CAPM)** | `C6` Risk-Free Rate, `C7` Levered Beta, `C8` ERP → `C9` Ke `=C6+C7*C8` |
| `11–16` | **2. After-Tax Cost of Debt & Preferred** | `C13` Pre-Tax Kd, `C14` Marginal Tax Rate → `C15` After-Tax Kd `=C13*(1-C14)`; `C16` Cost of Preferred Stock (`Kp`, non-tax-deductible; `0.0%` if Preferred `= $0M`) |
| `18–27` | **3. Capital Structure (Gross Cap `E + D + P`)** | `C20` Market Cap `=DCF!C14` (green link), `C21` Total Gross Debt `=DCF!C15` (green link), `C22` Preferred Stock `=DCF!C16` (green link), `C23` Total Capital `=SUM(C20:C22)`, `C24` Equity Weight `=C20/C23`, `C25` Debt Weight `=C21/C23`, `C26` Preferred Weight `=C22/C23`, `C27` Check `=SUM(C24:C26)` (`100.00%`) |
| `29–31` | **4. WACC Output & NCI Note** | `C30` `=C9*C24+C15*C25+C16*C26` (bold, `#BDD7EE` highlight); `A31:D31` NCI convention note (NCI excluded from parent WACC, deducted in equity bridge) |

---

## 3. `DCF` Sheet Row Map

### Header, Named Ranges & Case Selector (Rows 1–4)
- `DCF!B4`: Named cell `Case` (`wb.defined_names["Case"] = DefinedName("Case", attr_text="DCF!$B$4")`) containing `1` (Bear), `2` (Base), or `3` (Bull), enforced via `DataValidation(type="whole", operator="between", formula1="1", formula2="3", allow_blank=False)`. Blue bold input font with `#BDD7EE` fill.
- `DCF!C4`: `=CHOOSE(INT(Case),"Bear Case","Base Case","Bull Case")`.
- **Standard Workbook-Level Named Ranges (defined in `build_dcf_skeleton.py`, resolved first by `validate_dcf.py`):**
  - `Case` → `DCF!$B$4` (Active scenario selector: `1` = Bear, `2` = Base, `3` = Bull)
  - `DCF_WACC` → `WACC!$C$30` (Weighted Average Cost of Capital output)
  - `TerminalGrowth` → `DCF!$C$94` (Active Terminal Growth Rate `g` feeding Perpetuity TV; linked via `=C69`)
  - `TerminalGrowth_Bear` → `DCF!$C$33` (Bear Case Terminal Growth Rate)
  - `TerminalGrowth_Base` → `DCF!$C$45` (Base Case Terminal Growth Rate)
  - `TerminalGrowth_Bull` → `DCF!$C$57` (Bull Case Terminal Growth Rate)
  - `PV_TerminalValue` → `DCF!$C$99` (PV of Perpetuity Terminal Value)
  - `EnterpriseValue` → `DCF!$C$100` (Implied Enterprise Value)
  - *(Optional for multi-WACC models: `WACC_Bear`, `WACC_Base`, `WACC_Bull` — automatically resolved and paired per scenario by `validate_dcf.py` when defined.)*

### Section 1 — Market Data & Valuation Inputs (Rows 6–22)
- `C8` Current Share Price, `C9` Basic Shares Outstanding
- `C10` Unvested RSUs (`$0` strike price — always included in full regardless of share price)
- `C11` Options Outstanding, `C12` Weighted Average Option Strike Price
- `C13` Fully Diluted Shares (TSM): `=C9+C10+IF(C8>C12,C11-(C11*C12/C8),0)`
- `C14` Market Capitalization: `=C8*C13`
- `C15` Total Gross Debt, `C16` Preferred Stock, `C17` Non-Controlling Interest, `C18` Cash & Marketable Securities, `C19` Equity Investments
- `C20` Historical Anchor / Base-Year Revenue (e.g., LTM or latest FY actual)
- `C21` Discount Rate (WACC): `=WACC!C30` (green cross-sheet link)
- `C22` Effective Tax Rate: `=WACC!C14` (green cross-sheet link)

### Section 2 — Stacked Bear / Base / Bull Assumption Blocks (Rows 24–58)
Stack the three scenario blocks vertically with a **fixed 12-row stride** (`BLOCK_STRIDE = 12`) so every metric sits at the exact same relative offset (`1..9`) inside each block:
- **Bear Case (`Case = 1`)**: Banner row `24`, Year Headers row `25` (`C25:G25` = FY1E–FY5E), Metrics rows `26–34`
- **Base Case (`Case = 2`)**: Banner row `36`, Year Headers row `37`, Metrics rows `38–46`
- **Bull Case (`Case = 3`)**: Banner row `48`, Year Headers row `49`, Metrics rows `50–58`

Each block MUST include: (1) section banner, (2) column header row showing the explicit projection years (`FY1E`–`FY5E`), and (3) all 9 assumption rows in identical order:
1. Revenue Growth (`%`)
2. Gross Margin (`%`)
3. R&D `% of Revenue`
4. SG&A `% of Revenue` *(or split S&M / G&A as separate rows — keep uniform stride across all 3 blocks)*
5. D&A `% of Revenue`
6. CapEx `% of Revenue`
7. NWC `% of Revenue Change`
8. Terminal Growth Rate (`%`, scalar in col `C`, strictly `<= 3.5%` and `< WACC`)
9. Exit EV/EBITDA Multiple (`x`, scalar in col `C`)

### Section 3 — Active Case Consolidation Block (Rows 60–70)
Pull the active scenario dynamically using `INDEX` over the vertical stacked blocks with `INT(Case)` guarding against fractional values — **zero nested `IF` statements** in projection rows:

```excel
# Annual driver (offset k in 1..7, column col in C..G):
=INDEX(C$26:C$58,(INT(Case)-1)*12+k)

# Scalar terminal driver (offset k in 8..9, column C):
=INDEX(C$26:C$58,(INT(Case)-1)*12+k)
```

Row assignments in `build_dcf_skeleton.py`:
- `Row 62`: Active Revenue Growth (`C62:G62`)
- `Row 63`: Active Gross Margin (`C63:G63`)
- `Row 64`: Active R&D `% of Revenue` (`C64:G64`)
- `Row 65`: Active SG&A `% of Revenue` (`C65:G65`)
- `Row 66`: Active D&A `% of Revenue` (`C66:G66`)
- `Row 67`: Active CapEx `% of Revenue` (`C67:G67`)
- `Row 68`: Active NWC `% of Revenue Change` (`C68:G68`)
- `Row 69`: Active Terminal Growth Rate (`C69`)
- `Row 70`: Active Exit EV/EBITDA Multiple (`C70`)

### Section 4 — 5-Year Unlevered FCF Projections (Rows 72–89)
Column `B` holds the historical anchor year (`B74 = C20`); columns `C:G` hold projection years `1..5` (aligned identically with the Active Case columns `C:G`):

| Row | Line Item | Formula Pattern (Column `C` = Year 1; copy across `C:G`) |
|---|---|---|
| `74` | Revenue | `=B74*(1+C62)` |
| `75` | Gross Profit | `=C74*C63` |
| `76` | Less: R&D Expense | `=C74*C64` *(based on Revenue, NEVER Gross Profit, NEVER hardcoded %)* |
| `77` | Less: SG&A Expense | `=C74*C65` |
| `78` | Total Operating Expenses | `=C76+C77` |
| `79` | EBIT (Operating Income) | `=C75-C78` |
| `80` | Less: Cash Taxes on EBIT | `=C79*$C$22` |
| `81` | NOPAT | `=C79-C80` |
| `82` | Plus: D&A | `=C74*C66` |
| `83` | Memo: EBITDA | `=C79+C82` |
| `84` | Less: CapEx | `=C74*C67` |
| `85` | Less: Change in NWC | `=(C74-B74)*C68` |
| `86` | **Unlevered Free Cash Flow (UFCF)** | `=C81+C82-C84-C85` (`#BDD7EE` highlight) |
| `87` | Discount Period (Mid-Year) | `=1-0.5`, `=2-0.5`, ..., `=5-0.5` (`G87 = 4.5`) |
| `88` | Mid-Year Discount Factor | `=1/(1+$C$21)^C87` |
| `89` | **PV of Unlevered FCF** | `=C86*C88` |

*(If historical years are shown prior to `FY1E`, add historical actual columns to the left of the anchor column and shift the projection column letters consistently across Sections 2–7.)*

### Section 5 — Terminal Value & Enterprise-to-Equity Bridge (Rows 91–109)
- `C93` Terminal Year UFCF: `=G86`
- `C94` Active Terminal Growth (`g`): `=C69`
- `C95` Undiscounted Perpetuity TV: `=C93*(1+C94)/($C$21-C94)`
- `C96` Terminal Discount Factor (Mid-Year Perpetuity, `N - 0.5`): `=1/(1+$C$21)^$G$87`
- `C98` Sum of PV of Explicit FCFs: `=SUM(C89:G89)`
- `C99` PV of Terminal Value: `=C95*C96`
- `C100` **Enterprise Value**: `=C98+C99`
- `C101` Less: Total Gross Debt: `=-C15`
- `C102` Less: Preferred Stock: `=-C16`
- `C103` Less: Non-Controlling Interest: `=-C17`
- `C104` Plus: Cash & Marketable Securities: `=C18`
- `C105` Plus: Equity Investments: `=C19`
- `C106` **Implied Equity Value**: `=SUM(C100:C105)`
- `C107` Fully Diluted Shares (TSM): `=C13`
- `C108` **Implied Share Price (Perpetuity Growth)**: `=C106/C107` (`#BDD7EE` highlight, bold)
- `C109` Implied Upside / (Downside) `%`: `=C108/C8-1`

### Section 6 — Cross-Checks & Exit Multiple Method (Rows 111–119)
- `C113` TV % of Enterprise Value: `=C99/C100` (target `50%–70%`)
- `C114` Implied EV / Year 1 EBITDA: `=C100/C83`
- `C115` Implied Exit EV / Year 5 EBITDA (from Perpetuity): `=C95*(1+$C$21)^0.5/G83` *(compounds mid-year Perpetuity TV at `t = N - 0.5` by `(1 + WACC)^0.5` to end-of-year `t = N` so it is apples-to-apples with an end-of-year exit multiple)*
- `C116` Undiscounted Exit Multiple TV: `=G83*C70`
- `C117` PV of Exit Multiple TV: `=C116/(1+$C$21)^($G$87+0.5)` *(strictly end-of-year `(1+WACC)^N` = `(1+WACC)^5`, NOT `4.5`)*
- `C118` Implied Enterprise Value (Exit Multiple): `=C98+C117`
- `C119` **Implied Share Price (Exit Multiple)**: `=(C118+SUM(C101:C105))/C107`

---

## 4. Live-Formula 5×5 Sensitivity Analysis (Rows 122–129+)

Never use Excel's manual Data Table feature (LibreOffice re-export converts `<f t="dataTable">` into incompatible `TABLE(...)` formulas), and **never** use linear approximations or placeholder text.

### Primary 5×5 Grid: WACC vs. Terminal Growth (`DCF!B124:G129`)
- **Odd dimensions (5×5)** guarantee an exact center cell (`DCF!E127`).
- **Formula-driven axes** anchored on the model's live `WACC` (`C21`) and live `Terminal Growth` (`C69`):
  - Column header row `124` (`C124:G124`): `=E124-0.01`, `=E124-0.005`, **`=C69`**, `=E124+0.005`, `=E124+0.01`
  - Row header col `B` (`B125:B129`): `=B127-0.01`, `=B127-0.005`, **`=C21`**, `=B127+0.005`, `=B127+0.01`
- **Live horizon-independent `SUMPRODUCT` full-DCF formula in every cell (`C125:G129`)** — wrapped with `IF($B{row}<={col}${hdr},"NM",...)` so low-WACC / high-`g` corners where `WACC <= g` evaluate cleanly to `"NM"` instead of `#DIV/0!` or negative terminal values:

```excel
=IF($B125<=C$124,"NM",(SUMPRODUCT($C$86:$G$86/((1+$B125)^$C$87:$G$87))+(($G$86*(1+C$124)/($B125-C$124))/((1+$B125)^$G$87))+(-$C$15-$C$16-$C$17+$C$18+$C$19))/$C$13)
```

- **Center-cell verification:** `DCF!E127` MUST equal `DCF!C108` to floating-point precision. Highlight `E127` with `#BDD7EE` fill + bold black font.

### Optional Additional 5×5 Sensitivity Grids (Rows 131+)
When deeper scenario stress-testing is requested (e.g., for an initiation report or investment committee pack), stack additional live-formula 5×5 grids below Row 131 with 2 blank rows between tables:
1. **WACC vs. Exit EV/EBITDA Multiple** (recalculates PV of explicit FCFs via `SUMPRODUCT` at row WACC + `($G$83 * col_multiple) / (1 + row_WACC)^($G$87+0.5)` + equity bridge ÷ diluted shares).
2. **Terminal Revenue Growth Delta vs. Terminal EBIT Margin Delta** (recalculates full 5Y NOPAT/UFCF path via live formulas — never linear shortcuts).

---

## 5. Styling, Number Formats & Cell Comments

Aligned with `.agents/rules/financial-modeling.md` and `xlsx-author`:

| Element | Font | Fill | Border |
|---|---|---|---|
| Section banner | White bold (`#FFFFFF`) | Dark navy (`#1F4E79`) | — |
| Column header | Black bold (`#000000`) | Light blue (`#D9E1F2`) | Thin cell border |
| Hardcoded input | Blue (`#0000FF`) | Light grey (`#F2F2F2`) | — |
| Same-sheet formula | Black (`#000000`) | White | Thin top/bottom for subtotals |
| Cross-sheet link | Green (`#008000`) | White | — |
| Key output / Active selector / Sensitivity center | Black bold (or Blue bold for `B4`) | Medium blue (`#BDD7EE`) | Top thin, bottom double (`BORDER_TOTAL`) |

- **Number formats:**
  - Currency (`$M`): `$#,##0.0;($#,##0.0);"-"` (or `$#,##0;($#,##0);"-"` for large-cap `$M`)
  - Per-share price: `$#,##0.00;($#,##0.00);"-"`
  - Percentages: `0.0%` (margins/growth) and `0.00%` (WACC/Ke/Kd/g)
  - Multiples: `0.0"x"`
  - Years: Text strings (`"FY2026E"`)
- **Cell comments (MANDATORY):** Attach a cell comment (`Source: [Filing/System], [Date], [Reference]`) on EVERY blue hardcoded input as it is written — never defer to the end.

---

## 6. Common DCF Mistakes to Avoid

| Mistake | Why It Fails | Correct Pattern |
|---|---|---|
| Using `Net Debt / EV` in WACC | Produces negative debt weight for net-cash firms | Gross capitalization weights: `E/(E+D+P)`, `D/(E+D+P)`, `P/(E+D+P)` |
| Grouping RSUs with Options under a single strike check | Drops RSUs when Share Price `<` Option Strike | Separate `Unvested RSUs` (`$0` strike, always dilutive) from `Options Outstanding` |
| Hardcoded selector mismatch (`$B$6` vs `$B$4`) or float `Case` | Scenario switcher breaks or `CHOOSE`/`INDEX` diverge | Lock named cell `Case` at `DCF!$B$4`, add `1..3` `DataValidation`, and wrap in `INT()` |
| Discounting Mid-Year Perpetuity TV at `N` instead of `N - 0.5` | Introduces a half-year gap between Year `N` mid-year cash flow and Year `N+1` perpetuity stream | Discount Perpetuity TV at `(1 + WACC)^(N - 0.5)` (`^$G$87`) when explicit FCFs are mid-year |
| Discounting Exit-Multiple TV at `N - 0.5` | Overstates PV of point-in-time exit transaction by `(1 + WACC)^0.5` | Always discount Exit-Multiple TV strictly by `(1 + WACC)^N` (`^($G$87+0.5)`) |
| Terminal growth `> 3.5%` (e.g., Bull `4–5%`) | Exceeds long-run nominal GDP | Keep `g` in `2.0%–3.5%` and extend forecast years if needed |
| Deducting ASC 842 operating lease liabilities when FCF deducts rent | Double-counts lease obligations | Exclude ASC 842 operating lease liabilities from EV bridge unless rent is removed from OpEx |
| Adding back SBC to Unlevered FCF | Overstates economic FCF | Treat SBC as real operating expense (or reconcile explicitly) |
| Writing OpEx as `% of Gross Profit` or hardcoding `%` inside formulas | Breaks operating leverage and auditability | Model every OpEx line as `=Rev * Active_OpEx_%_Cell` |
| Hardcoded `0.5..4.5` exponents, unguarded `WACC - g` division, or Excel What-If Data Tables in sensitivities | Breaks when extending horizon, produces `#DIV/0!` / negative TV when `WACC <= g`, or corrupts on LibreOffice re-export | Use dynamic `=IF($B125<=C$124,"NM", ... SUMPRODUCT($C$86:$G$86/((1+$B125)^$C$87:$G$87)) ...)` formulas in all 25 grid cells |
| Inserting rows after writing formulas | Shifts references → `#REF!` or silent wrong-row links | Lock all section row constants before writing any formulas |
