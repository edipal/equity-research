#!/usr/bin/env python3
# Independent implementation for Google Antigravity, licensed under the Apache License, Version 2.0.
"""Build an institutional, 2-sheet DCF valuation model skeleton (DCF + WACC) in Excel.

Usage:
    python3 .agents/skills/dcf-model/scripts/build_dcf_skeleton.py --output model.xlsx [--case 2]
"""

from __future__ import annotations

import argparse
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

# Institutional color palette (aligned with .agents/rules/financial-modeling.md)
COLOR_NAVY = "1F4E79"
COLOR_SUBHEADER = "D9E1F2"
COLOR_INPUT_BG = "F2F2F2"
COLOR_HIGHLIGHT = "BDD7EE"
COLOR_WHITE = "FFFFFF"
COLOR_BORDER = "BFBFBF"

FONT_FAMILY = "Calibri"
FONT_TITLE = Font(name=FONT_FAMILY, size=14, bold=True, color=COLOR_WHITE)
FONT_SECTION = Font(name=FONT_FAMILY, size=11, bold=True, color=COLOR_WHITE)
FONT_COL_HEADER = Font(name=FONT_FAMILY, size=10, bold=True, color="000000")
FONT_INPUT = Font(name=FONT_FAMILY, size=10, bold=False, color="0000FF")
FONT_INPUT_BOLD = Font(name=FONT_FAMILY, size=10, bold=True, color="0000FF")
FONT_FORMULA = Font(name=FONT_FAMILY, size=10, bold=False, color="000000")
FONT_FORMULA_BOLD = Font(name=FONT_FAMILY, size=10, bold=True, color="000000")
FONT_LINK = Font(name=FONT_FAMILY, size=10, bold=False, color="008000")
FONT_LINK_BOLD = Font(name=FONT_FAMILY, size=10, bold=True, color="008000")
FONT_NOTE = Font(name=FONT_FAMILY, size=9, italic=True, color="595959")

FILL_NAVY = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
FILL_SUBHEADER = PatternFill(start_color=COLOR_SUBHEADER, end_color=COLOR_SUBHEADER, fill_type="solid")
FILL_INPUT = PatternFill(start_color=COLOR_INPUT_BG, end_color=COLOR_INPUT_BG, fill_type="solid")
FILL_HIGHLIGHT = PatternFill(start_color=COLOR_HIGHLIGHT, end_color=COLOR_HIGHLIGHT, fill_type="solid")

THIN_SIDE = Side(style="thin", color=COLOR_BORDER)
DOUBLE_SIDE = Side(style="double", color="000000")
TOP_THIN = Side(style="thin", color="000000")

BORDER_CELL = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
BORDER_TOTAL = Border(top=TOP_THIN, bottom=DOUBLE_SIDE)
BORDER_SUBTOTAL = Border(top=TOP_THIN, bottom=TOP_THIN)

FMT_CURRENCY_M = "$#,##0.0;($#,##0.0);\"-\""
FMT_PRICE = "$#,##0.00;($#,##0.00);\"-\""
FMT_PCT = "0.0%"
FMT_PCT_2 = "0.00%"
FMT_DEC_1 = "#,##0.0"
FMT_DEC_2 = "0.00"
FMT_DEC_4 = "0.0000"
FMT_MULT = "0.0\"x\""
FMT_INT = "0"


def _comment(text: str) -> Comment:
    return Comment(text, "Equity Research Harness")


def _style_banner(ws, row: int, max_col: int, title: str, font: Font = FONT_SECTION) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = font
    cell.alignment = Alignment(horizontal="left", vertical="center")
    for col in range(1, max_col + 1):
        ws.cell(row=row, column=col).fill = FILL_NAVY


def _style_col_headers(ws, row: int, headers: list[str]) -> None:
    for idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=idx, value=text)
        cell.font = FONT_COL_HEADER
        cell.fill = FILL_SUBHEADER
        cell.border = BORDER_CELL
        cell.alignment = Alignment(
            horizontal="left" if idx == 1 else "right",
            vertical="center",
        )


def _write_input_cell(ws, row: int, col: int, value, num_format: str, note: str, bold: bool = False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = FONT_INPUT_BOLD if bold else FONT_INPUT
    cell.fill = FILL_INPUT
    cell.number_format = num_format
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.comment = _comment(note)
    return cell


def _write_formula_cell(
    ws,
    row: int,
    col: int,
    formula: str,
    num_format: str,
    *,
    bold: bool = False,
    cross_sheet: bool = False,
    highlight: bool = False,
    border: Border | None = None,
):
    cell = ws.cell(row=row, column=col, value=formula)
    if cross_sheet:
        cell.font = FONT_LINK_BOLD if bold else FONT_LINK
    else:
        cell.font = FONT_FORMULA_BOLD if bold else FONT_FORMULA
    if highlight:
        cell.fill = FILL_HIGHLIGHT
    if border is not None:
        cell.border = border
    cell.number_format = num_format
    cell.alignment = Alignment(horizontal="right", vertical="center")
    return cell


def _build_wacc_sheet(ws) -> None:
    """Populate the WACC worksheet using gross-debt capital weights and CAPM."""
    ws.views.sheetView[0].showGridLines = True

    _style_banner(ws, 1, 4, "WEIGHTED AVERAGE COST OF CAPITAL (WACC) BUILD", font=FONT_TITLE)
    ws.cell(row=2, column=1, value="Template / Illustrative Company — All dollar figures in $ millions").font = FONT_NOTE

    # Section 1: Cost of Equity (CAPM)
    _style_banner(ws, 4, 4, "1. COST OF EQUITY (CAPM)")
    _style_col_headers(ws, 5, ["Component", "Units", "Value", "Source / Basis"])

    ws.cell(row=6, column=1, value="Risk-Free Rate (10Y US Treasury)").font = FONT_FORMULA
    ws.cell(row=6, column=2, value="%").font = FONT_NOTE
    _write_input_cell(ws, 6, 3, 0.0425, FMT_PCT_2, "Source: US Treasury 10Y yield placeholder [ASSUMPTION]")
    ws.cell(row=6, column=4, value="10-Year US Treasury benchmark yield").font = FONT_NOTE

    ws.cell(row=7, column=1, value="Equity Beta (5Y Monthly)").font = FONT_FORMULA
    ws.cell(row=7, column=2, value="x").font = FONT_NOTE
    _write_input_cell(ws, 7, 3, 1.15, FMT_DEC_2, "Source: 5-year monthly regression vs S&P 500 [ASSUMPTION]")
    ws.cell(row=7, column=4, value="5-year monthly levered beta").font = FONT_NOTE

    ws.cell(row=8, column=1, value="Equity Risk Premium (ERP)").font = FONT_FORMULA
    ws.cell(row=8, column=2, value="%").font = FONT_NOTE
    _write_input_cell(ws, 8, 3, 0.0575, FMT_PCT_2, "Source: Damodaran / institutional ERP [ASSUMPTION]")
    ws.cell(row=8, column=4, value="Long-run implied US equity risk premium").font = FONT_NOTE

    ws.cell(row=9, column=1, value="Cost of Equity (Ke)").font = FONT_FORMULA_BOLD
    ws.cell(row=9, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(ws, 9, 3, "=C6+C7*C8", FMT_PCT_2, bold=True, border=BORDER_SUBTOTAL)
    ws.cell(row=9, column=4, value="CAPM: Rf + Beta * ERP").font = FONT_NOTE

    # Section 2: Cost of Debt
    _style_banner(ws, 11, 4, "2. AFTER-TAX COST OF DEBT")
    _style_col_headers(ws, 12, ["Component", "Units", "Value", "Source / Basis"])

    ws.cell(row=13, column=1, value="Pre-Tax Cost of Debt (Kd)").font = FONT_FORMULA
    ws.cell(row=13, column=2, value="%").font = FONT_NOTE
    _write_input_cell(ws, 13, 3, 0.055, FMT_PCT_2, "Source: Corporate bond YTM / weighted avg borrowing rate [ASSUMPTION]")
    ws.cell(row=13, column=4, value="Weighted-average yield on gross debt").font = FONT_NOTE

    ws.cell(row=14, column=1, value="Effective Marginal Tax Rate").font = FONT_FORMULA
    ws.cell(row=14, column=2, value="%").font = FONT_NOTE
    _write_input_cell(ws, 14, 3, 0.21, FMT_PCT, "Source: US federal statutory tax rate [ASSUMPTION]")
    ws.cell(row=14, column=4, value="Statutory + state marginal tax shield").font = FONT_NOTE

    ws.cell(row=15, column=1, value="After-Tax Cost of Debt").font = FONT_FORMULA_BOLD
    ws.cell(row=15, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(ws, 15, 3, "=C13*(1-C14)", FMT_PCT_2, bold=True, border=BORDER_SUBTOTAL)
    ws.cell(row=15, column=4, value="Pre-tax Kd * (1 - Tax Rate)").font = FONT_NOTE

    # Section 3: Capital Structure Weights (Gross Debt basis — never negative debt weight)
    _style_banner(ws, 17, 4, "3. CAPITAL STRUCTURE WEIGHTS (GROSS DEBT BASIS)")
    _style_col_headers(ws, 18, ["Component", "Units", "Value", "Source / Basis"])

    ws.cell(row=19, column=1, value="Market Value of Equity (E)").font = FONT_FORMULA
    ws.cell(row=19, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 19, 3, "=DCF!C13", FMT_CURRENCY_M, cross_sheet=True)
    ws.cell(row=19, column=4, value="Linked from DCF Market Capitalization").font = FONT_NOTE

    ws.cell(row=20, column=1, value="Total Gross Debt (D)").font = FONT_FORMULA
    ws.cell(row=20, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 20, 3, "=DCF!C14", FMT_CURRENCY_M, cross_sheet=True)
    ws.cell(row=20, column=4, value="Linked from DCF Total Gross Debt (gross basis, >= 0)").font = FONT_NOTE

    ws.cell(row=21, column=1, value="Total Capitalization (E + D)").font = FONT_FORMULA_BOLD
    ws.cell(row=21, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 21, 3, "=C19+C20", FMT_CURRENCY_M, bold=True, border=BORDER_SUBTOTAL)
    ws.cell(row=21, column=4, value="Market Cap + Total Gross Debt").font = FONT_NOTE

    ws.cell(row=22, column=1, value="Equity Weight in WACC (E / (E + D))").font = FONT_FORMULA
    ws.cell(row=22, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(ws, 22, 3, "=C19/C21", FMT_PCT_2)
    ws.cell(row=22, column=4, value="Gross-capitalization equity weight").font = FONT_NOTE

    ws.cell(row=23, column=1, value="Debt Weight in WACC (D / (E + D))").font = FONT_FORMULA
    ws.cell(row=23, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(ws, 23, 3, "=C20/C21", FMT_PCT_2)
    ws.cell(row=23, column=4, value="Gross-capitalization debt weight (always >= 0%)").font = FONT_NOTE

    ws.cell(row=24, column=1, value="Capital Weight Sum Check").font = FONT_NOTE
    ws.cell(row=24, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(ws, 24, 3, "=C22+C23", FMT_PCT_2)
    ws.cell(row=24, column=4, value="Must equal 100.00%").font = FONT_NOTE

    # Section 4: WACC Output
    _style_banner(ws, 26, 4, "4. WACC OUTPUT")
    ws.cell(row=27, column=1, value="WEIGHTED AVERAGE COST OF CAPITAL").font = FONT_FORMULA_BOLD
    ws.cell(row=27, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(
        ws,
        27,
        3,
        "=C9*C22+C15*C23",
        FMT_PCT_2,
        bold=True,
        highlight=True,
        border=BORDER_TOTAL,
    )
    ws.cell(row=27, column=4, value="(Ke * Equity Weight) + (After-Tax Kd * Debt Weight)").font = FONT_NOTE

    col_widths = {1: 40, 2: 10, 3: 18, 4: 50}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def _build_dcf_sheet(ws, wb: openpyxl.Workbook, default_case: int = 2) -> None:
    """Populate the DCF worksheet with scenario blocks, 5Y projections, TV, EV bridge, and 5x5 sensitivity."""
    ws.views.sheetView[0].showGridLines = True
    years = ["FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]

    # Header & Scenario Selector (Named cell `Case` at DCF!B4)
    _style_banner(ws, 1, 8, "DISCOUNTED CASH FLOW (DCF) VALUATION MODEL", font=FONT_TITLE)
    ws.cell(row=2, column=1, value="Illustrative Institutional Template — All dollar figures in $ millions except per-share data").font = FONT_NOTE

    ws.cell(row=4, column=1, value="Active Scenario Selector (1=Bear, 2=Base, 3=Bull)").font = FONT_FORMULA_BOLD
    _write_input_cell(
        ws,
        4,
        2,
        default_case,
        FMT_INT,
        "Named cell `Case`: 1 = Bear Case, 2 = Base Case, 3 = Bull Case",
        bold=True,
    )
    ws["B4"].fill = FILL_HIGHLIGHT
    wb.defined_names["Case"] = DefinedName("Case", attr_text="DCF!$B$4")

    case_label_cell = ws.cell(
        row=4,
        column=3,
        value='=CHOOSE(Case,"Bear Case","Base Case","Bull Case")',
    )
    case_label_cell.font = FONT_FORMULA_BOLD
    case_label_cell.fill = FILL_HIGHLIGHT
    case_label_cell.alignment = Alignment(horizontal="center", vertical="center")
    case_label_cell.border = BORDER_CELL

    # Section 1: Market Data & Valuation Inputs (Rows 6-22)
    _style_banner(ws, 6, 8, "1. MARKET DATA & VALUATION INPUTS")
    _style_col_headers(ws, 7, ["Item", "Units", "Value", "Notes / Source", "", "", "", ""])

    ws.cell(row=8, column=1, value="Current Share Price").font = FONT_FORMULA
    ws.cell(row=8, column=2, value="$").font = FONT_NOTE
    _write_input_cell(ws, 8, 3, 25.00, FMT_PRICE, "Source: Market closing price placeholder [ASSUMPTION]")
    ws.cell(row=8, column=4, value="Reference closing price").font = FONT_NOTE

    ws.cell(row=9, column=1, value="Basic Shares Outstanding").font = FONT_FORMULA
    ws.cell(row=9, column=2, value="M").font = FONT_NOTE
    _write_input_cell(ws, 9, 3, 96.0, FMT_DEC_1, "Source: Latest 10-Q/10-K cover page [ASSUMPTION]")
    ws.cell(row=9, column=4, value="Basic common shares").font = FONT_NOTE

    ws.cell(row=10, column=1, value="Dilutive Options & RSUs").font = FONT_FORMULA
    ws.cell(row=10, column=2, value="M").font = FONT_NOTE
    _write_input_cell(ws, 10, 3, 6.0, FMT_DEC_1, "Source: Stock-based compensation footnote [ASSUMPTION]")
    ws.cell(row=10, column=4, value="In-the-money options and unvested RSUs").font = FONT_NOTE

    ws.cell(row=11, column=1, value="Weighted Average Option Strike Price").font = FONT_FORMULA
    ws.cell(row=11, column=2, value="$").font = FONT_NOTE
    _write_input_cell(ws, 11, 3, 15.00, FMT_PRICE, "Source: Stock-based compensation footnote [ASSUMPTION]")
    ws.cell(row=11, column=4, value="Weighted-average exercise price for TSM").font = FONT_NOTE

    ws.cell(row=12, column=1, value="Fully Diluted Shares (TSM)").font = FONT_FORMULA_BOLD
    ws.cell(row=12, column=2, value="M").font = FONT_NOTE
    _write_formula_cell(
        ws,
        12,
        3,
        "=C9+IF(C8>C11,C10-(C10*C11/C8),0)",
        FMT_DEC_1,
        bold=True,
        border=BORDER_SUBTOTAL,
    )
    ws.cell(row=12, column=4, value="Basic Shares + Treasury Stock Method net dilution").font = FONT_NOTE

    ws.cell(row=13, column=1, value="Market Capitalization").font = FONT_FORMULA_BOLD
    ws.cell(row=13, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 13, 3, "=C8*C12", FMT_CURRENCY_M, bold=True)
    ws.cell(row=13, column=4, value="Current Share Price * Fully Diluted Shares").font = FONT_NOTE

    ws.cell(row=14, column=1, value="Total Gross Debt").font = FONT_FORMULA
    ws.cell(row=14, column=2, value="$M").font = FONT_NOTE
    _write_input_cell(ws, 14, 3, 500.0, FMT_CURRENCY_M, "Source: Balance Sheet ST + LT Debt & finance leases [ASSUMPTION]")
    ws.cell(row=14, column=4, value="Short-term + long-term gross debt").font = FONT_NOTE

    ws.cell(row=15, column=1, value="Preferred Stock").font = FONT_FORMULA
    ws.cell(row=15, column=2, value="$M").font = FONT_NOTE
    _write_input_cell(ws, 15, 3, 25.0, FMT_CURRENCY_M, "Source: Balance Sheet liquidation value of preferred [ASSUMPTION]")
    ws.cell(row=15, column=4, value="Preferred equity liquidation preference").font = FONT_NOTE

    ws.cell(row=16, column=1, value="Non-Controlling Interest").font = FONT_FORMULA
    ws.cell(row=16, column=2, value="$M").font = FONT_NOTE
    _write_input_cell(ws, 16, 3, 15.0, FMT_CURRENCY_M, "Source: Balance Sheet minority interest [ASSUMPTION]")
    ws.cell(row=16, column=4, value="Minority / non-controlling interest").font = FONT_NOTE

    ws.cell(row=17, column=1, value="Cash & Marketable Securities").font = FONT_FORMULA
    ws.cell(row=17, column=2, value="$M").font = FONT_NOTE
    _write_input_cell(ws, 17, 3, 300.0, FMT_CURRENCY_M, "Source: Balance Sheet cash + short-term investments [ASSUMPTION]")
    ws.cell(row=17, column=4, value="Unrestricted cash and marketable securities").font = FONT_NOTE

    ws.cell(row=18, column=1, value="Equity Investments").font = FONT_FORMULA
    ws.cell(row=18, column=2, value="$M").font = FONT_NOTE
    _write_input_cell(ws, 18, 3, 40.0, FMT_CURRENCY_M, "Source: Non-consolidated equity method investments [ASSUMPTION]")
    ws.cell(row=18, column=4, value="Non-operating unconsolidated equity stakes").font = FONT_NOTE

    ws.cell(row=19, column=1, value="LTM / Base Year Revenue (FY2025A)").font = FONT_FORMULA
    ws.cell(row=19, column=2, value="$M").font = FONT_NOTE
    _write_input_cell(ws, 19, 3, 1000.0, FMT_CURRENCY_M, "Source: FY2025 10-K Income Statement placeholder [ASSUMPTION]")
    ws.cell(row=19, column=4, value="Historical anchor revenue for Year 1 projection").font = FONT_NOTE

    ws.cell(row=20, column=1, value="Discount Rate (WACC)").font = FONT_FORMULA_BOLD
    ws.cell(row=20, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(ws, 20, 3, "=WACC!C27", FMT_PCT_2, bold=True, cross_sheet=True)
    ws.cell(row=20, column=4, value="Linked from WACC worksheet").font = FONT_NOTE

    ws.cell(row=21, column=1, value="Effective Tax Rate").font = FONT_FORMULA
    ws.cell(row=21, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(ws, 21, 3, "=WACC!C14", FMT_PCT, cross_sheet=True)
    ws.cell(row=21, column=4, value="Linked from WACC worksheet").font = FONT_NOTE

    # Section 2: Vertically Stacked Scenario Assumption Blocks (Bear / Base / Bull)
    # Fixed stride of 12 rows per block so Active Case uses a single clean INDEX(range, ($B$4-1)*12 + offset)
    # Bear header: row 23, col header: row 24, metrics: rows 25..33 (9 rows), blank: row 34
    # Base header: row 35, col header: row 36, metrics: rows 37..45 (9 rows), blank: row 46
    # Bull header: row 47, col header: row 48, metrics: rows 49..57 (9 rows), blank: row 58
    BLOCK_STRIDE = 12
    FIRST_METRIC_ROW = 25
    LAST_METRIC_ROW = 57  # row 49 + 8

    scenario_specs = [
        (
            23,
            "2A. BEAR CASE ASSUMPTIONS (Case = 1)",
            {
                "Revenue Growth": ([0.06, 0.05, 0.04, 0.035, 0.03], FMT_PCT),
                "Gross Margin": ([0.62, 0.62, 0.615, 0.615, 0.61], FMT_PCT),
                "R&D % of Revenue": ([0.14, 0.14, 0.135, 0.135, 0.135], FMT_PCT),
                "SG&A % of Revenue": ([0.25, 0.25, 0.245, 0.245, 0.245], FMT_PCT),
                "D&A % of Revenue": ([0.045, 0.045, 0.045, 0.045, 0.045], FMT_PCT),
                "CapEx % of Revenue": ([0.055, 0.055, 0.050, 0.050, 0.050], FMT_PCT),
                "NWC % of Revenue Change": ([0.02, 0.02, 0.02, 0.02, 0.02], FMT_PCT),
                "Terminal Growth Rate": ([0.020], FMT_PCT_2),
                "Exit EV/EBITDA Multiple": ([9.0], FMT_MULT),
            },
        ),
        (
            35,
            "2B. BASE CASE ASSUMPTIONS (Case = 2)",
            {
                "Revenue Growth": ([0.10, 0.09, 0.08, 0.07, 0.06], FMT_PCT),
                "Gross Margin": ([0.65, 0.655, 0.66, 0.665, 0.67], FMT_PCT),
                "R&D % of Revenue": ([0.14, 0.135, 0.13, 0.125, 0.12], FMT_PCT),
                "SG&A % of Revenue": ([0.24, 0.235, 0.23, 0.225, 0.22], FMT_PCT),
                "D&A % of Revenue": ([0.045, 0.045, 0.045, 0.045, 0.045], FMT_PCT),
                "CapEx % of Revenue": ([0.055, 0.050, 0.050, 0.045, 0.045], FMT_PCT),
                "NWC % of Revenue Change": ([0.015, 0.015, 0.015, 0.015, 0.015], FMT_PCT),
                "Terminal Growth Rate": ([0.025], FMT_PCT_2),
                "Exit EV/EBITDA Multiple": ([11.0], FMT_MULT),
            },
        ),
        (
            47,
            "2C. BULL CASE ASSUMPTIONS (Case = 3)",
            {
                "Revenue Growth": ([0.14, 0.13, 0.11, 0.09, 0.08], FMT_PCT),
                "Gross Margin": ([0.67, 0.68, 0.685, 0.69, 0.695], FMT_PCT),
                "R&D % of Revenue": ([0.135, 0.13, 0.125, 0.12, 0.115], FMT_PCT),
                "SG&A % of Revenue": ([0.23, 0.22, 0.215, 0.21, 0.205], FMT_PCT),
                "D&A % of Revenue": ([0.045, 0.045, 0.045, 0.045, 0.045], FMT_PCT),
                "CapEx % of Revenue": ([0.055, 0.050, 0.045, 0.045, 0.040], FMT_PCT),
                "NWC % of Revenue Change": ([0.010, 0.010, 0.010, 0.010, 0.010], FMT_PCT),
                "Terminal Growth Rate": ([0.030], FMT_PCT_2),
                "Exit EV/EBITDA Multiple": ([13.0], FMT_MULT),
            },
        ),
    ]

    metric_order = [
        ("Revenue Growth", "%", FMT_PCT, True),
        ("Gross Margin", "%", FMT_PCT, True),
        ("R&D % of Revenue", "%", FMT_PCT, True),
        ("SG&A % of Revenue", "%", FMT_PCT, True),
        ("D&A % of Revenue", "%", FMT_PCT, True),
        ("CapEx % of Revenue", "%", FMT_PCT, True),
        ("NWC % of Revenue Change", "%", FMT_PCT, True),
        ("Terminal Growth Rate", "%", FMT_PCT_2, False),
        ("Exit EV/EBITDA Multiple", "x", FMT_MULT, False),
    ]

    for start_row, banner_title, values_map in scenario_specs:
        _style_banner(ws, start_row, 8, banner_title)
        _style_col_headers(ws, start_row + 1, ["Assumption", "Units", *years, "Notes"])
        for offset, (label, unit, fmt, is_annual) in enumerate(metric_order):
            row = start_row + 2 + offset
            ws.cell(row=row, column=1, value=label).font = FONT_FORMULA
            ws.cell(row=row, column=2, value=unit).font = FONT_NOTE
            vals, _ = values_map[label]
            if is_annual:
                for yr_idx, val in enumerate(vals):
                    _write_input_cell(
                        ws,
                        row,
                        3 + yr_idx,
                        val,
                        fmt,
                        f"{banner_title}: {label} ({years[yr_idx]}) [ASSUMPTION]",
                    )
            else:
                _write_input_cell(
                    ws,
                    row,
                    3,
                    vals[0],
                    fmt,
                    f"{banner_title}: {label} [ASSUMPTION]",
                )
                ws.cell(row=row, column=8, value="Scalar terminal assumption").font = FONT_NOTE

    # Section 3: Active Case Consolidation Block driven by INDEX(..., $B$4) (Rows 59-69)
    ACTIVE_HEADER_ROW = 59
    ACTIVE_START_ROW = 61  # rows 61..69 map 1:1 to metric_order
    _style_banner(ws, ACTIVE_HEADER_ROW, 8, "3. ACTIVE CASE ASSUMPTIONS (DRIVEN BY INDEX / CASE SELECTOR)")
    _style_col_headers(ws, ACTIVE_HEADER_ROW + 1, ["Active Assumption", "Units", *years, "Formula Pattern"])

    for offset, (label, unit, fmt, is_annual) in enumerate(metric_order):
        row = ACTIVE_START_ROW + offset
        ws.cell(row=row, column=1, value=label).font = FONT_FORMULA_BOLD if not is_annual else FONT_FORMULA
        ws.cell(row=row, column=2, value=unit).font = FONT_NOTE
        if is_annual:
            for yr_idx in range(5):
                col = 3 + yr_idx
                col_ltr = get_column_letter(col)
                formula = (
                    f"=INDEX({col_ltr}${FIRST_METRIC_ROW}:{col_ltr}${LAST_METRIC_ROW},"
                    f"($B$4-1)*{BLOCK_STRIDE}+{offset + 1})"
                )
                _write_formula_cell(ws, row, col, formula, fmt)
            ws.cell(row=row, column=8, value="INDEX over stacked Bear/Base/Bull blocks").font = FONT_NOTE
        else:
            formula = (
                f"=INDEX(C${FIRST_METRIC_ROW}:C${LAST_METRIC_ROW},"
                f"($B$4-1)*{BLOCK_STRIDE}+{offset + 1})"
            )
            _write_formula_cell(ws, row, 3, formula, fmt, bold=True)
            ws.cell(row=row, column=8, value="INDEX scalar terminal driver").font = FONT_NOTE

    # Active assumption row constants:
    ROW_ACT_REV_G = ACTIVE_START_ROW + 0   # 61
    ROW_ACT_GM = ACTIVE_START_ROW + 1      # 62
    ROW_ACT_RD = ACTIVE_START_ROW + 2      # 63
    ROW_ACT_SGA = ACTIVE_START_ROW + 3     # 64
    ROW_ACT_DA = ACTIVE_START_ROW + 4      # 65
    ROW_ACT_CAPEX = ACTIVE_START_ROW + 5   # 66
    ROW_ACT_NWC = ACTIVE_START_ROW + 6     # 67
    ROW_ACT_TG = ACTIVE_START_ROW + 7      # 68
    ROW_ACT_EXIT_M = ACTIVE_START_ROW + 8  # 69

    # Section 4: 5-Year Unlevered FCF Projections (Rows 71-88)
    _style_banner(ws, 71, 8, "4. UNLEVERED FREE CASH FLOW PROJECTIONS ($M)")
    _style_col_headers(ws, 72, ["Line Item", "FY2025A", *years, "Notes"])

    # Row 73: Revenue
    ROW_REV = 73
    ws.cell(row=ROW_REV, column=1, value="Revenue").font = FONT_FORMULA_BOLD
    _write_formula_cell(ws, ROW_REV, 2, "=C19", FMT_CURRENCY_M, bold=True)
    for yr_idx in range(5):
        col = 3 + yr_idx
        prev_ltr = get_column_letter(col - 1)
        col_ltr = get_column_letter(col)
        _write_formula_cell(
            ws,
            ROW_REV,
            col,
            f"={prev_ltr}{ROW_REV}*(1+{col_ltr}{ROW_ACT_REV_G})",
            FMT_CURRENCY_M,
            bold=True,
        )
    ws.cell(row=ROW_REV, column=8, value="Prior Revenue * (1 + Active Revenue Growth)").font = FONT_NOTE

    # Row 74: Gross Profit
    ROW_GP = 74
    ws.cell(row=ROW_GP, column=1, value="Gross Profit").font = FONT_FORMULA
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(ws, ROW_GP, col, f"={col_ltr}{ROW_REV}*{col_ltr}{ROW_ACT_GM}", FMT_CURRENCY_M)

    # Row 75: R&D Expense
    ROW_RD = 75
    ws.cell(row=ROW_RD, column=1, value="Less: R&D Expense").font = FONT_FORMULA
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(ws, ROW_RD, col, f"={col_ltr}{ROW_REV}*{col_ltr}{ROW_ACT_RD}", FMT_CURRENCY_M)

    # Row 76: SG&A Expense
    ROW_SGA = 76
    ws.cell(row=ROW_SGA, column=1, value="Less: SG&A Expense").font = FONT_FORMULA
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(ws, ROW_SGA, col, f"={col_ltr}{ROW_REV}*{col_ltr}{ROW_ACT_SGA}", FMT_CURRENCY_M)

    # Row 77: Total Operating Expenses
    ROW_OPEX = 77
    ws.cell(row=ROW_OPEX, column=1, value="Total Operating Expenses").font = FONT_FORMULA_BOLD
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(
            ws,
            ROW_OPEX,
            col,
            f"={col_ltr}{ROW_RD}+{col_ltr}{ROW_SGA}",
            FMT_CURRENCY_M,
            bold=True,
            border=BORDER_SUBTOTAL,
        )

    # Row 78: EBIT (Operating Income)
    ROW_EBIT = 78
    ws.cell(row=ROW_EBIT, column=1, value="EBIT (Operating Income)").font = FONT_FORMULA_BOLD
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(
            ws,
            ROW_EBIT,
            col,
            f"={col_ltr}{ROW_GP}-{col_ltr}{ROW_OPEX}",
            FMT_CURRENCY_M,
            bold=True,
        )

    # Row 79: Taxes on EBIT
    ROW_TAX = 79
    ws.cell(row=ROW_TAX, column=1, value="Less: Cash Taxes on EBIT").font = FONT_FORMULA
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(ws, ROW_TAX, col, f"={col_ltr}{ROW_EBIT}*$C$21", FMT_CURRENCY_M)

    # Row 80: NOPAT
    ROW_NOPAT = 80
    ws.cell(row=ROW_NOPAT, column=1, value="NOPAT").font = FONT_FORMULA_BOLD
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(
            ws,
            ROW_NOPAT,
            col,
            f"={col_ltr}{ROW_EBIT}-{col_ltr}{ROW_TAX}",
            FMT_CURRENCY_M,
            bold=True,
            border=BORDER_SUBTOTAL,
        )

    # Row 81: D&A
    ROW_DA = 81
    ws.cell(row=ROW_DA, column=1, value="Plus: Depreciation & Amortization (D&A)").font = FONT_FORMULA
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(ws, ROW_DA, col, f"={col_ltr}{ROW_REV}*{col_ltr}{ROW_ACT_DA}", FMT_CURRENCY_M)

    # Row 82: EBITDA (memo for Exit Multiple & cross-checks)
    ROW_EBITDA = 82
    ws.cell(row=ROW_EBITDA, column=1, value="Memo: EBITDA (EBIT + D&A)").font = FONT_NOTE
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(ws, ROW_EBITDA, col, f"={col_ltr}{ROW_EBIT}+{col_ltr}{ROW_DA}", FMT_CURRENCY_M)

    # Row 83: CapEx
    ROW_CAPEX = 83
    ws.cell(row=ROW_CAPEX, column=1, value="Less: Capital Expenditures (CapEx)").font = FONT_FORMULA
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(ws, ROW_CAPEX, col, f"={col_ltr}{ROW_REV}*{col_ltr}{ROW_ACT_CAPEX}", FMT_CURRENCY_M)

    # Row 84: Change in NWC
    ROW_DNWC = 84
    ws.cell(row=ROW_DNWC, column=1, value="Less: Change in Net Working Capital").font = FONT_FORMULA
    for yr_idx in range(5):
        col = 3 + yr_idx
        prev_ltr = get_column_letter(col - 1)
        col_ltr = get_column_letter(col)
        _write_formula_cell(
            ws,
            ROW_DNWC,
            col,
            f"=({col_ltr}{ROW_REV}-{prev_ltr}{ROW_REV})*{col_ltr}{ROW_ACT_NWC}",
            FMT_CURRENCY_M,
        )

    # Row 85: Unlevered Free Cash Flow
    ROW_UFCF = 85
    ws.cell(row=ROW_UFCF, column=1, value="Unlevered Free Cash Flow (UFCF)").font = FONT_FORMULA_BOLD
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(
            ws,
            ROW_UFCF,
            col,
            f"={col_ltr}{ROW_NOPAT}+{col_ltr}{ROW_DA}-{col_ltr}{ROW_CAPEX}-{col_ltr}{ROW_DNWC}",
            FMT_CURRENCY_M,
            bold=True,
            highlight=True,
            border=BORDER_TOTAL,
        )

    # Row 86: Discount Period (Mid-Year Convention)
    ROW_PERIOD = 86
    ws.cell(row=ROW_PERIOD, column=1, value="Discount Period (Mid-Year Convention)").font = FONT_FORMULA
    for yr_idx in range(5):
        col = 3 + yr_idx
        # Live formula referencing year index (1..5) minus 0.5
        _write_formula_cell(ws, ROW_PERIOD, col, f"={yr_idx + 1}-0.5", FMT_DEC_1)
    ws.cell(row=ROW_PERIOD, column=8, value="t - 0.5 (mid-year cash flow timing)").font = FONT_NOTE

    # Row 87: Discount Factor
    ROW_DF = 87
    ws.cell(row=ROW_DF, column=1, value="Mid-Year Discount Factor").font = FONT_FORMULA
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(ws, ROW_DF, col, f"=1/(1+$C$20)^{col_ltr}{ROW_PERIOD}", FMT_DEC_4)

    # Row 88: PV of Unlevered FCF
    ROW_PV_FCF = 88
    ws.cell(row=ROW_PV_FCF, column=1, value="Present Value of Unlevered FCF").font = FONT_FORMULA_BOLD
    for yr_idx in range(5):
        col = 3 + yr_idx
        col_ltr = get_column_letter(col)
        _write_formula_cell(
            ws,
            ROW_PV_FCF,
            col,
            f"={col_ltr}{ROW_UFCF}*{col_ltr}{ROW_DF}",
            FMT_CURRENCY_M,
            bold=True,
        )

    # Section 5: Terminal Value & Enterprise-to-Equity Valuation Bridge (Rows 90-118)
    _style_banner(ws, 90, 8, "5. TERMINAL VALUE & ENTERPRISE-TO-EQUITY VALUATION BRIDGE")
    _style_col_headers(ws, 91, ["Valuation Component", "Units", "Value", "Methodology / Formula Basis", "", "", "", ""])

    # 5A. Perpetuity Growth Terminal Value (Primary)
    ws.cell(row=92, column=1, value="Terminal Year (FY2030E) Unlevered FCF").font = FONT_FORMULA
    ws.cell(row=92, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 92, 3, f"=G{ROW_UFCF}", FMT_CURRENCY_M)

    ws.cell(row=93, column=1, value="Active Terminal Growth Rate (g)").font = FONT_FORMULA
    ws.cell(row=93, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(ws, 93, 3, f"=C{ROW_ACT_TG}", FMT_PCT_2)

    ws.cell(row=94, column=1, value="Undiscounted Terminal Value (Perpetuity Growth)").font = FONT_FORMULA
    ws.cell(row=94, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 94, 3, "=C92*(1+C93)/($C$20-C93)", FMT_CURRENCY_M)
    ws.cell(row=94, column=4, value="FCF_5 * (1 + g) / (WACC - g)").font = FONT_NOTE

    ws.cell(row=95, column=1, value="Terminal Discount Factor (End of Year 5)").font = FONT_FORMULA
    ws.cell(row=95, column=2, value="x").font = FONT_NOTE
    _write_formula_cell(ws, 95, 3, "=1/(1+$C$20)^5", FMT_DEC_4)
    ws.cell(row=95, column=4, value="End-of-year Year 5 discounting: 1 / (1 + WACC)^5").font = FONT_NOTE

    # Bridge (Primary Perpetuity Growth)
    ws.cell(row=97, column=1, value="Sum of PV of Explicit FCFs (FY2026E-FY2030E)").font = FONT_FORMULA
    ws.cell(row=97, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 97, 3, f"=SUM(C{ROW_PV_FCF}:G{ROW_PV_FCF})", FMT_CURRENCY_M)

    ws.cell(row=98, column=1, value="PV of Terminal Value").font = FONT_FORMULA
    ws.cell(row=98, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 98, 3, "=C94*C95", FMT_CURRENCY_M)
    ws.cell(row=98, column=4, value="Perpetuity Growth Terminal Value discounted at (1 + WACC)^5").font = FONT_NOTE

    ws.cell(row=99, column=1, value="Enterprise Value").font = FONT_FORMULA_BOLD
    ws.cell(row=99, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 99, 3, "=C97+C98", FMT_CURRENCY_M, bold=True, border=BORDER_SUBTOTAL)

    ws.cell(row=100, column=1, value="Less: Total Gross Debt").font = FONT_FORMULA
    ws.cell(row=100, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 100, 3, "=-C14", FMT_CURRENCY_M)

    ws.cell(row=101, column=1, value="Less: Preferred Stock").font = FONT_FORMULA
    ws.cell(row=101, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 101, 3, "=-C15", FMT_CURRENCY_M)

    ws.cell(row=102, column=1, value="Less: Non-Controlling Interest").font = FONT_FORMULA
    ws.cell(row=102, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 102, 3, "=-C16", FMT_CURRENCY_M)

    ws.cell(row=103, column=1, value="Plus: Cash & Marketable Securities").font = FONT_FORMULA
    ws.cell(row=103, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 103, 3, "=C17", FMT_CURRENCY_M)

    ws.cell(row=104, column=1, value="Plus: Equity Investments").font = FONT_FORMULA
    ws.cell(row=104, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 104, 3, "=C18", FMT_CURRENCY_M)

    ws.cell(row=105, column=1, value="Implied Equity Value").font = FONT_FORMULA_BOLD
    ws.cell(row=105, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 105, 3, "=SUM(C99:C104)", FMT_CURRENCY_M, bold=True, border=BORDER_SUBTOTAL)

    ws.cell(row=106, column=1, value="Fully Diluted Shares (TSM)").font = FONT_FORMULA
    ws.cell(row=106, column=2, value="M").font = FONT_NOTE
    _write_formula_cell(ws, 106, 3, "=C12", FMT_DEC_1)

    ROW_IMPLIED_PRICE = 107
    ws.cell(row=ROW_IMPLIED_PRICE, column=1, value="Implied Share Price (Perpetuity Growth)").font = FONT_FORMULA_BOLD
    ws.cell(row=ROW_IMPLIED_PRICE, column=2, value="$").font = FONT_NOTE
    _write_formula_cell(
        ws,
        ROW_IMPLIED_PRICE,
        3,
        "=C105/C106",
        FMT_PRICE,
        bold=True,
        highlight=True,
        border=BORDER_TOTAL,
    )

    ws.cell(row=108, column=1, value="Implied Upside / (Downside) %").font = FONT_FORMULA_BOLD
    ws.cell(row=108, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(ws, 108, 3, f"=C{ROW_IMPLIED_PRICE}/C8-1", FMT_PCT, bold=True)

    # 5B. Valuation Cross-Checks & Exit Multiple Method (Rows 110-119)
    _style_banner(ws, 110, 8, "6. VALUATION CROSS-CHECKS & EXIT MULTIPLE METHOD")
    _style_col_headers(ws, 111, ["Cross-Check Metric", "Units", "Value", "Notes", "", "", "", ""])

    ws.cell(row=112, column=1, value="TV % of Enterprise Value").font = FONT_FORMULA
    ws.cell(row=112, column=2, value="%").font = FONT_NOTE
    _write_formula_cell(ws, 112, 3, "=C98/C99", FMT_PCT)
    ws.cell(row=112, column=4, value="Target institutional range: 50%-70%").font = FONT_NOTE

    ws.cell(row=113, column=1, value="Implied EV / FY2026E EBITDA").font = FONT_FORMULA
    ws.cell(row=113, column=2, value="x").font = FONT_NOTE
    _write_formula_cell(ws, 113, 3, f"=C99/C{ROW_EBITDA}", FMT_MULT)
    ws.cell(row=113, column=4, value="Cross-check vs peer NTM EV/EBITDA").font = FONT_NOTE

    ws.cell(row=114, column=1, value="Implied Exit EV / FY2030E EBITDA (from Perpetuity)").font = FONT_FORMULA
    ws.cell(row=114, column=2, value="x").font = FONT_NOTE
    _write_formula_cell(ws, 114, 3, f"=C94/G{ROW_EBITDA}", FMT_MULT)
    ws.cell(row=114, column=4, value="Undiscounted Perpetuity TV / Year 5 EBITDA").font = FONT_NOTE

    ws.cell(row=115, column=1, value="Undiscounted Exit Multiple TV (FY2030E EBITDA * Multiple)").font = FONT_FORMULA
    ws.cell(row=115, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 115, 3, f"=G{ROW_EBITDA}*C{ROW_ACT_EXIT_M}", FMT_CURRENCY_M)

    ws.cell(row=116, column=1, value="PV of Exit Multiple TV (Discounted at (1 + WACC)^5)").font = FONT_FORMULA
    ws.cell(row=116, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 116, 3, "=C115/(1+$C$20)^5", FMT_CURRENCY_M)
    ws.cell(row=116, column=4, value="Strictly end-of-year Year 5 discounting (NOT N-0.5)").font = FONT_NOTE

    ws.cell(row=117, column=1, value="Implied Enterprise Value (Exit Multiple Method)").font = FONT_FORMULA
    ws.cell(row=117, column=2, value="$M").font = FONT_NOTE
    _write_formula_cell(ws, 117, 3, "=C97+C116", FMT_CURRENCY_M)

    ws.cell(row=118, column=1, value="Implied Share Price (Exit Multiple Method)").font = FONT_FORMULA_BOLD
    ws.cell(row=118, column=2, value="$").font = FONT_NOTE
    _write_formula_cell(ws, 118, 3, "=(C117+SUM(C100:C104))/C106", FMT_PRICE, bold=True)

    # Section 7: 5x5 Live-Formula Sensitivity Analysis (Rows 121-128)
    SENS_BANNER_ROW = 121
    SENS_HEADER_ROW = 123
    SENS_FIRST_DATA_ROW = 124  # rows 124..128 (center row = 126)
    SENS_CENTER_ROW = 126
    SENS_CENTER_COL = 5        # cols C(3), D(4), E(5), F(6), G(7) -> center = E(5)

    _style_banner(ws, SENS_BANNER_ROW, 8, "7. SENSITIVITY ANALYSIS: WACC vs TERMINAL GROWTH (IMPLIED SHARE PRICE $)")
    ws.cell(
        row=122,
        column=1,
        value="Live full-DCF recalculation in every cell; center cell (highlighted) matches Active Case Implied Share Price.",
    ).font = FONT_NOTE

    corner = ws.cell(row=SENS_HEADER_ROW, column=2, value="WACC \\ Terminal g")
    corner.font = FONT_COL_HEADER
    corner.fill = FILL_SUBHEADER
    corner.border = BORDER_CELL
    corner.alignment = Alignment(horizontal="center", vertical="center")

    # Formula-driven Terminal Growth axis (centered on live Active Case Terminal Growth C68)
    tg_axis_formulas = {
        3: f"={get_column_letter(SENS_CENTER_COL)}{SENS_HEADER_ROW}-0.01",
        4: f"={get_column_letter(SENS_CENTER_COL)}{SENS_HEADER_ROW}-0.005",
        5: f"=C{ROW_ACT_TG}",
        6: f"={get_column_letter(SENS_CENTER_COL)}{SENS_HEADER_ROW}+0.005",
        7: f"={get_column_letter(SENS_CENTER_COL)}{SENS_HEADER_ROW}+0.01",
    }
    for col, formula in tg_axis_formulas.items():
        cell = _write_formula_cell(ws, SENS_HEADER_ROW, col, formula, FMT_PCT_2, bold=True, border=BORDER_CELL)
        cell.fill = FILL_SUBHEADER

    # Formula-driven WACC axis (centered on live WACC C20)
    wacc_axis_formulas = {
        124: f"=B{SENS_CENTER_ROW}-0.01",
        125: f"=B{SENS_CENTER_ROW}-0.005",
        126: "=C20",
        127: f"=B{SENS_CENTER_ROW}+0.005",
        128: f"=B{SENS_CENTER_ROW}+0.01",
    }
    for row, formula in wacc_axis_formulas.items():
        cell = _write_formula_cell(ws, row, 2, formula, FMT_PCT_2, bold=True, border=BORDER_CELL)
        cell.fill = FILL_SUBHEADER

    # Populate all 25 sensitivity cells with live full-DCF implied share price formulas
    for row in range(SENS_FIRST_DATA_ROW, SENS_FIRST_DATA_ROW + 5):
        for col in range(3, 8):
            col_ltr = get_column_letter(col)
            w_ref = f"$B{row}"
            g_ref = f"{col_ltr}${SENS_HEADER_ROW}"
            pv_fcf_expr = (
                f"($C${ROW_UFCF}/(1+{w_ref})^0.5"
                f"+$D${ROW_UFCF}/(1+{w_ref})^1.5"
                f"+$E${ROW_UFCF}/(1+{w_ref})^2.5"
                f"+$F${ROW_UFCF}/(1+{w_ref})^3.5"
                f"+$G${ROW_UFCF}/(1+{w_ref})^4.5)"
            )
            pv_tv_expr = f"(($G${ROW_UFCF}*(1+{g_ref})/({w_ref}-{g_ref}))/(1+{w_ref})^5)"
            eq_adj_expr = "(-$C$14-$C$15-$C$16+$C$17+$C$18)"
            formula = f"=({pv_fcf_expr}+{pv_tv_expr}+{eq_adj_expr})/$C$12"
            is_center = row == SENS_CENTER_ROW and col == SENS_CENTER_COL
            _write_formula_cell(
                ws,
                row,
                col,
                formula,
                FMT_PRICE,
                bold=is_center,
                highlight=is_center,
                border=BORDER_CELL,
            )

    col_widths = {1: 48, 2: 18, 3: 15, 4: 15, 5: 15, 6: 15, 7: 15, 8: 44}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def build_dcf_skeleton(output_path: Path, default_case: int = 2) -> Path:
    """Create a spec-compliant 2-sheet DCF workbook on disk."""
    if default_case not in (1, 2, 3):
        raise ValueError(f"default_case must be 1 (Bear), 2 (Base), or 3 (Bull); got {default_case}")

    wb = openpyxl.Workbook()
    ws_dcf = wb.active
    ws_dcf.title = "DCF"
    ws_wacc = wb.create_sheet(title="WACC")

    _build_wacc_sheet(ws_wacc)
    _build_dcf_sheet(ws_dcf, wb, default_case=default_case)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()
    return output_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Build an institutional 2-sheet DCF valuation model skeleton (DCF + WACC)."
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Destination path for the generated .xlsx workbook.",
    )
    parser.add_argument(
        "--case",
        type=int,
        choices=(1, 2, 3),
        default=2,
        help="Default active scenario (1=Bear, 2=Base, 3=Bull; default: 2).",
    )
    args = parser.parse_args(argv)
    out = build_dcf_skeleton(Path(args.output), default_case=args.case)
    print(f"Created DCF skeleton workbook: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
