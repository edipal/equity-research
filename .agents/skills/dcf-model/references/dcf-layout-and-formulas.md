<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/financial-analysis/skills/dcf-model/SKILL.md), Apache-2.0; modified for Google Antigravity. -->

# Canonical DCF Workbook Layout, Formulas & Formatting

Reference specification for `.agents/skills/dcf-model/SKILL.md`. All row coordinates, sheet names, and formulas match `.agents/skills/dcf-model/scripts/build_dcf_skeleton.py`.

---

## 1. Two-Sheet Workbook Architecture

| Sheet | Role | Key Outputs |
|---|---|---|
| `DCF` | Main valuation sheet: named `Case` selector (`B4`), market/bridge inputs, stacked Bear/Base/Bull assumption blocks, `INDEX`-driven Active Case consolidation block, 5Y UFCF build, Terminal Value & Equity Bridge, Exit-Multiple cross-check, and inline 5×5 live-formula sensitivity table(s) at bottom | `DCF!C107` (Implied Share Price — Perpetuity), `DCF!C118` (Implied Share Price — Exit Multiple), `DCF!E126` (Sensitivity center cell = `DCF!C107`) |
| `WACC` | Single source of truth for CAPM Cost of Equity, After-Tax Cost of Debt, and Gross-Debt Capital Structure Weights | `WACC!C27` (linked into `DCF!C20`) |

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
| `11–15` | **2. After-Tax Cost of Debt** | `C13` Pre-Tax Kd, `C14` Marginal Tax Rate → `C15` `=C13*(1-C14)` |
| `17–24` | **3. Capital Structure (Gross Debt)** | `C19` Market Cap `=DCF!C13` (green link), `C20` Total Gross Debt `=DCF!C14` (green link), `C21` Total Capital `=C19+C20`, `C22` Equity Weight `=C19/C21`, `C23` Debt Weight `=C20/C21`, `C24` Check `=C22+C23` (`100.00%`) |
| `26–27` | **4. WACC Output** | `C27` `=C9*C22+C15*C23` (bold, `#BDD7EE` highlight) |

---

## 3. `DCF` Sheet Row Map

### Header & Named Case Selector (Rows 1–4)
- `DCF!B4`: Named cell `Case` (`wb.defined_names["Case"] = DefinedName("Case", attr_text="DCF!$B$4")`) containing `1` (Bear), `2` (Base), or `3` (Bull). Blue bold input font with `#BDD7EE` fill.
- `DCF!C4`: `=CHOOSE(Case,"Bear Case","Base Case","Bull Case")`.

### Section 1 — Market Data & Valuation Inputs (Rows 6–21)
- `C8` Current Share Price, `C9` Basic Shares, `C10` Dilutive Options & RSUs, `C11` Weighted-Avg Strike Price
- `C12` Fully Diluted Shares (TSM): `=C9+IF(C8>C11,C10-(C10*C11/C8),0)`
- `C13` Market Capitalization: `=C8*C12`
- `C14` Total Gross Debt, `C15` Preferred Stock, `C16` Non-Controlling Interest, `C17` Cash & Marketable Securities, `C18` Equity Investments
- `C19` Historical Anchor / Base-Year Revenue (e.g., LTM or latest FY actual)
- `C20` Discount Rate (WACC): `=WACC!C27` (green cross-sheet link)
- `C21` Effective Tax Rate: `=WACC!C14` (green cross-sheet link)

### Section 2 — Stacked Bear / Base / Bull Assumption Blocks (Rows 23–57)
Stack the three scenario blocks vertically with a **fixed 12-row stride** (`BLOCK_STRIDE = 12`) so every metric sits at the exact same relative offset (`1..9`) inside each block:
- **Bear Case (`Case = 1`)**: Banner row `23`, Year Headers row `24` (`C24:G24` = FY1E–FY5E), Metrics rows `25–33`
- **Base Case (`Case = 2`)**: Banner row `35`, Year Headers row `36`, Metrics rows `37–45`
- **Bull Case (`Case = 3`)**: Banner row `47`, Year Headers row `48`, Metrics rows `49–57`

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

### Section 3 — Active Case Consolidation Block (Rows 59–69)
Pull the active scenario dynamically using `INDEX` over the vertical stacked blocks — **zero nested `IF` statements** in projection rows:

```excel
# Annual driver (offset k in 1..7, column col in C..G):
=INDEX(C$25:C$57,($B$4-1)*12+k)

# Scalar terminal driver (offset k in 8..9, column C):
=INDEX(C$25:C$57,($B$4-1)*12+k)
```

Row assignments in `build_dcf_skeleton.py`:
- `Row 61`: Active Revenue Growth (`C61:G61`)
- `Row 62`: Active Gross Margin (`C62:G62`)
- `Row 63`: Active R&D `% of Revenue` (`C63:G63`)
- `Row 64`: Active SG&A `% of Revenue` (`C64:G64`)
- `Row 65`: Active D&A `% of Revenue` (`C65:G65`)
- `Row 66`: Active CapEx `% of Revenue` (`C66:G66`)
- `Row 67`: Active NWC `% of Revenue Change` (`C67:G67`)
- `Row 68`: Active Terminal Growth Rate (`C68`)
- `Row 69`: Active Exit EV/EBITDA Multiple (`C69`)

### Section 4 — 5-Year Unlevered FCF Projections (Rows 71–88)
Column `B` holds the historical anchor year (`B73 = C19`); columns `C:G` hold projection years `1..5` (aligned identically with the Active Case columns `C:G`):

| Row | Line Item | Formula Pattern (Column `C` = Year 1; copy across `C:G`) |
|---|---|---|
| `73` | Revenue | `=B73*(1+C61)` |
| `74` | Gross Profit | `=C73*C62` |
| `75` | Less: R&D Expense | `=C73*C63` *(based on Revenue, NEVER Gross Profit, NEVER hardcoded %)* |
| `76` | Less: SG&A Expense | `=C73*C64` |
| `77` | Total Operating Expenses | `=C75+C76` |
| `78` | EBIT (Operating Income) | `=C74-C77` |
| `79` | Less: Cash Taxes on EBIT | `=C78*$C$21` |
| `80` | NOPAT | `=C78-C79` |
| `81` | Plus: D&A | `=C73*C65` |
| `82` | Memo: EBITDA | `=C78+C81` |
| `83` | Less: CapEx | `=C73*C66` |
| `84` | Less: Change in NWC | `=(C73-B73)*C67` |
| `85` | **Unlevered Free Cash Flow (UFCF)** | `=C80+C81-C83-C84` (`#BDD7EE` highlight) |
| `86` | Discount Period (Mid-Year) | `=1-0.5`, `=2-0.5`, ..., `=5-0.5` |
| `87` | Mid-Year Discount Factor | `=1/(1+$C$20)^C86` |
| `88` | **PV of Unlevered FCF** | `=C85*C87` |

*(If historical years are shown prior to `FY1E`, add historical actual columns to the left of the anchor column and shift the projection column letters consistently across Sections 2–7.)*

### Section 5 — Terminal Value & Enterprise-to-Equity Bridge (Rows 90–108)
- `C92` Terminal Year UFCF: `=G85`
- `C93` Active Terminal Growth (`g`): `=C68`
- `C94` Undiscounted Perpetuity TV: `=C92*(1+C93)/($C$20-C93)`
- `C95` Terminal Discount Factor (End of Year 5): `=1/(1+$C$20)^5`
- `C97` Sum of PV of Explicit FCFs: `=SUM(C88:G88)`
- `C98` PV of Terminal Value: `=C94*C95`
- `C99` **Enterprise Value**: `=C97+C98`
- `C100` Less: Total Gross Debt: `=-C14`
- `C101` Less: Preferred Stock: `=-C15`
- `C102` Less: Non-Controlling Interest: `=-C16`
- `C103` Plus: Cash & Marketable Securities: `=C17`
- `C104` Plus: Equity Investments: `=C18`
- `C105` **Implied Equity Value**: `=SUM(C99:C104)`
- `C106` Fully Diluted Shares (TSM): `=C12`
- `C107` **Implied Share Price (Perpetuity Growth)**: `=C105/C106` (`#BDD7EE` highlight, bold)
- `C108` Implied Upside / (Downside) `%`: `=C107/C8-1`

### Section 6 — Cross-Checks & Exit Multiple Method (Rows 110–118)
- `C112` TV % of Enterprise Value: `=C98/C99` (target `50%–70%`)
- `C113` Implied EV / Year 1 EBITDA: `=C99/C82`
- `C114` Implied Exit EV / Year 5 EBITDA (from Perpetuity): `=C94/G82`
- `C115` Undiscounted Exit Multiple TV: `=G82*C69`
- `C116` PV of Exit Multiple TV: `=C115/(1+$C$20)^5` *(strictly `(1+WACC)^5`, NOT `4.5`)*
- `C117` Implied Enterprise Value (Exit Multiple): `=C97+C116`
- `C118` **Implied Share Price (Exit Multiple)**: `=(C117+SUM(C100:C104))/C106`

---

## 4. Live-Formula 5×5 Sensitivity Analysis (Rows 121–128+)

Never use Excel's manual Data Table feature (cannot be recalculated cleanly by openpyxl/LibreOffice without manual GUI steps), and **never** use linear approximations or placeholder text.

### Primary 5×5 Grid: WACC vs. Terminal Growth (`DCF!B123:G128`)
- **Odd dimensions (5×5)** guarantee an exact center cell (`DCF!E126`).
- **Formula-driven axes** anchored on the model's live `WACC` (`C20`) and live `Terminal Growth` (`C68`):
  - Column header row `123` (`C123:G123`): `=E123-0.01`, `=E123-0.005`, **`=C68`**, `=E123+0.005`, `=E123+0.01`
  - Row header col `B` (`B124:B128`): `=B126-0.01`, `=B126-0.005`, **`=C20`**, `=B126+0.005`, `=B126+0.01`
- **Live full-DCF formula in every cell (`C124:G128`)**:

```excel
=(($C$85/(1+$B124)^0.5+$D$85/(1+$B124)^1.5+$E$85/(1+$B124)^2.5+$F$85/(1+$B124)^3.5+$G$85/(1+$B124)^4.5)+(($G$85*(1+C$123)/($B124-C$123))/(1+$B124)^5)+(-$C$14-$C$15-$C$16+$C$17+$C$18))/$C$12
```

- **Center-cell verification:** `DCF!E126` MUST equal `DCF!C107` to the cent. Highlight `E126` with `#BDD7EE` fill + bold black font.

### Optional Additional 5×5 Sensitivity Grids (Rows 130+)
When deeper scenario stress-testing is requested (e.g., for an initiation report or investment committee pack), stack additional live-formula 5×5 grids below Row 130 with 2 blank rows between tables:
1. **WACC vs. Exit EV/EBITDA Multiple** (recalculates PV of explicit FCFs at row WACC + `($G$82 * col_multiple) / (1 + row_WACC)^5` + equity bridge ÷ diluted shares).
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
| Using `Net Debt / EV` in WACC | Produces negative debt weight for net-cash firms | Gross debt weights: `E/(E+D)` and `D/(E+D)` |
| Hardcoded selector mismatch (`$B$6` vs `$B$4`) | Scenario switcher silently breaks | Lock named cell `Case` at `DCF!$B$4` and verify `C4` + Active Case block |
| Discounting Exit-Multiple TV at `N - 0.5` | Overstates PV of TV by `(1 + WACC)^0.5` | Always discount Exit-Multiple TV by `(1 + WACC)^N` |
| Terminal growth `> 3.5%` (e.g., Bull `4–5%`) | Exceeds long-run nominal GDP | Keep `g` in `2.0%–3.5%` and extend forecast years if needed |
| Deducting ASC 842 operating lease liabilities when FCF deducts rent | Double-counts lease obligations | Exclude ASC 842 operating lease liabilities from EV bridge unless rent is removed from OpEx |
| Adding back SBC to Unlevered FCF | Overstates economic FCF | Treat SBC as real operating expense (or reconcile explicitly) |
| Writing OpEx as `% of Gross Profit` or hardcoding `%` inside formulas | Breaks operating leverage and auditability | Model every OpEx line as `=Rev * Active_OpEx_%_Cell` |
| Linear approximation or blank sensitivity cells | Misstates convexity; fails QA | Write full live DCF formula into all 25 grid cells via `openpyxl` loop |
| Inserting rows after writing formulas | Shifts references → `#REF!` or silent wrong-row links | Lock all section row constants before writing any formulas |
