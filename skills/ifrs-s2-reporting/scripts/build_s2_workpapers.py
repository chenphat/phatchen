#!/usr/bin/env python3
"""Build the IFRS S2 working-paper workbook.

Sheets: CR-00 index/reliefs, CR-01 climate risk register, CR-02 strategy & financial effects,
SA scenario analysis (SA1–SA9), CR-03 cross-industry metrics, CR-04 targets & carbon credits,
CR-05 industry metrics, GHG-00 to GHG-09 (incl. 03a, 03b, 07b), FE-01 financed emissions (--bank),
S2-CL disclosure checklist.

Usage:
    python build_s2_workpapers.py --entity "Mekong Garment Co." --year-end 2026-12-31 \
        --currency "KHR million" --out S2_Working_Papers.xlsx [--bank] [--no-examples]

Calculations pull factors from the GHG-02 register by Factor ID (INDEX/MATCH) so a factor is changed
in one controlled place. Example data = illustrative Mekong Garment case (UNFCCC calculator / DEFRA 2022
factors, AR6 refrigerant GWP); replace with current, sourced values. Recalculate before sharing.
"""
import argparse
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
F_EX = Font(name=FONT, size=10, italic=True, color="7F7F7F")
F_NOTE = Font(name=FONT, size=9, italic=True, color="404040")
F_LINK = Font(name=FONT, size=10, color="008000")
FILL_HEAD = PatternFill("solid", fgColor="1F3864")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_TOT = PatternFill("solid", fgColor="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
NUM = "#,##0.00;(#,##0.00);-"
NUM0 = "#,##0;(#,##0);-"
EXAMPLES = True


def header(wb, ref, title, anchor, args, purpose):
    ws = wb.create_sheet(ref)
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"{ref}  {title}"
    ws["A1"].font = F_TITLE
    for i, (k, v) in enumerate([("Entity", args.entity), ("Period ended", args.year_end),
                                ("Presentation unit", args.currency), ("IFRS anchor", anchor),
                                ("Purpose", purpose)], start=2):
        ws.cell(i, 1, k).font = F_BOLD
        ws.cell(i, 2, v).font = F_BASE
    for r, lab in zip((2, 3, 4), ("Prepared by / date", "Reviewed by / date", "Approved by / date")):
        ws.cell(r, 6, lab).font = F_BOLD
        for col in (7, 8):
            ws.cell(r, col).fill = FILL_INPUT
            ws.cell(r, col).border = BORDER
    ws["A7"] = ("Legend: yellow = input; black formulas = calculated; green = link to another WP; "
                "grey italic = illustrative example — overwrite or delete.")
    ws["A7"].font = F_NOTE
    return ws


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def table(ws, start_row, headers, rows, widths=None, input_cols=None, example_rows=1, height=None, fmt=None):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(start_row, j, h)
        c.font, c.fill, c.border = F_HEAD, FILL_HEAD, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[start_row].height = 32
    input_cols = set(input_cols or [])
    fmt = fmt or {}
    for i, row in enumerate(rows):
        r = start_row + 1 + i
        for j, v in enumerate(row, start=1):
            if isinstance(v, str) and "{r}" in v:
                v = v.replace("{r}", str(r))
            c = ws.cell(r, j, v)
            c.border, c.alignment = BORDER, WRAP
            ex = EXAMPLES and i < example_rows
            if is_formula(v):
                c.font = F_LINK if "!" in v else F_BASE
            else:
                c.font = F_EX if ex else F_BASE
                if j in input_cols:
                    c.fill = FILL_INPUT
            if j in fmt:
                c.number_format = fmt[j]
        if height:
            ws.row_dimensions[r].height = height
    if widths:
        for j, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w
    return start_row + len(rows)


def ex(rows, n_blank, ncols, filler=None):
    """Example rows (if enabled) followed by blank rows; filler maps col->formula template for blanks."""
    out = list(rows) if EXAMPLES else []
    filler = filler or {}
    for _ in range(n_blank):
        out.append([filler.get(j) for j in range(1, ncols + 1)])
    return out


def dv_list(ws, rng, options):
    d = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    ws.add_data_validation(d)
    d.add(rng)


def total_row(ws, r, label_col, label, sum_cols, first, last, fmt=NUM):
    ws.cell(r, label_col, label).font = F_BOLD
    for col in sum_cols:
        L = get_column_letter(col)
        c = ws.cell(r, col, f"=SUM({L}{first}:{L}{last})")
        c.font, c.fill, c.number_format, c.border = F_BOLD, FILL_TOT, fmt, BORDER


# ----------------------------------------------------------------------------- factor register
FACTORS = [
    ["EF-ELEC-KH", "Cambodia grid electricity (location-based)", 0.5878, "kg CO2e/kWh", "S2", "IFI 2021 harmonised grid factor (UNFCCC calculator v1.3)", "2021", "CO2e (as published)", 0.50, "Confirm official MoE/EAC factor and year"],
    ["EF-DIESEL", "Diesel, 100% mineral (stationary/mobile)", 2.6988, "kg CO2e/L", "S1", "DEFRA 2022", "2022", "CO2e (AR5 inside factor)", None, "Update to DESNZ 2025"],
    ["EF-DIESEL-WTT", "Diesel well-to-tank", 0.6287, "kg CO2e/L", "S3 Cat 3", "DEFRA 2022", "2022", "CO2e", None, "Update to DESNZ 2025"],
    ["EF-PETROL", "Petrol, average blend", 2.1619, "kg CO2e/L", "S1", "DEFRA 2022", "2022", "CO2e", None, "Update"],
    ["EF-LPG", "LPG", 1.5571, "kg CO2e/L", "S1", "DEFRA 2022", "2022", "CO2e", None, "Update"],
    ["EF-VAN", "Van up to 3.5 t, diesel", 0.2316, "kg CO2e/km", "S1", "DEFRA 2022", "2022", "CO2e", None, "Update"],
    ["GWP-R410A", "Refrigerant R-410A", 2256, "kg CO2e/kg", "S1", "IPCC AR6 (blend) — calculator uses AR4 2,088", "AR6", "AR6 100-yr", 2088, "S2.B21–B22: use AR6"],
    ["GWP-R32", "Refrigerant HFC-32 (R-32)", 771, "kg CO2e/kg", "S1", "IPCC AR6", "AR6", "AR6 100-yr", 675, "S2.B21–B22"],
    ["EF-WATER-S", "Water supply", 0.149, "kg CO2e/m3", "S3 Cat 1", "DEFRA 2022 (UK proxy)", "2022", "CO2e", None, "Find local proxy"],
    ["EF-WATER-T", "Wastewater treatment", 0.272, "kg CO2e/m3", "S3 Cat 5", "DEFRA 2022 (UK proxy)", "2022", "CO2e", None, "Find local proxy"],
    ["EF-WASTE-LF", "Commercial & industrial waste to landfill", 467.0, "kg CO2e/t", "S3 Cat 5", "DEFRA 2022 (UK proxy)", "2022", "CO2e", None, "Find local proxy"],
    ["EF-BUS", "Local bus (proxy for worker trucks/buses)", 0.1078, "kg CO2e/pkm", "S3 Cat 7", "DEFRA 2022", "2022", "CO2e", None, "Update"],
    ["EF-FABRIC", "Purchased fabric (placeholder)", 9.0, "kg CO2e/kg", "S3 Cat 1", "PLACEHOLDER — obtain supplier EPDs", "—", "CO2e", None, "Replace before use"],
]
FR_FIRST, FR_LAST = 10, 50  # register rows


def lookup(col_letter_id, value_col="C"):
    return (f"=IFERROR(INDEX('GHG-02'!${value_col}${FR_FIRST}:${value_col}${FR_LAST},"
            f"MATCH({col_letter_id}{{r}},'GHG-02'!$A${FR_FIRST}:$A${FR_LAST},0)),\"\")")


def ghg_02(wb, a):
    ws = header(wb, "GHG-02", "Emission factor register", "S2.29(a)(iii); S2.B21–B22", a,
                "Single controlled source for every factor: value, unit, source, year, GWP basis, approval, change vs prior year.")
    chg = '=IF(AND(ISNUMBER(C{r}),ISNUMBER(I{r}),I{r}<>0),C{r}/I{r}-1,"")'
    rows = [f[:9] + [chg, f[9], None] for f in FACTORS] if True else []
    rows += [[None] * 9 + [chg, None, None] for _ in range(FR_LAST - FR_FIRST + 1 - len(rows))]
    end = table(ws, 9, ["Factor ID", "Activity", "Value", "Unit", "Scope / category", "Source", "Year", "GWP basis",
                        "Prior-year value", "Change %", "Action / note", "Approved by"], rows,
                [14, 38, 10, 14, 12, 40, 8, 18, 11, 9, 28, 12], input_cols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12],
                example_rows=len(FACTORS), fmt={10: "0.0%"})
    ws.conditional_formatting.add(f"J10:J{end}", CellIsRule(operator="notBetween", formula=["-0.05", "0.05"],
                                                            fill=PatternFill("solid", fgColor="F8CBAD")))
    ws.cell(end + 2, 1, "Rule: changes >5% (Scope 1/2) or >10% (Scope 3) vs prior year are highlighted for review; "
                        "never apply GWP to a factor already in CO2e; refrigerants at AR6 unless a jurisdictional GWP relief applies (Dec 2025 amendments).").font = F_NOTE


def calc_sheet(wb, a, ref, title, anchor, purpose, rows_ex, extra_col=None, n_blank=15):
    """Generic AD × EF calc sheet. Columns: A Source, B Site/category, C Qty, D Unit, E Data ref, F Factor ID,
    G Factor (lookup), H Factor unit (lookup), I tCO2e, J Notes."""
    ws = header(wb, ref, title, anchor, a, purpose)
    g = lookup("F", "C")
    h = lookup("F", "D")
    i = '=IF(AND(ISNUMBER(C{r}),ISNUMBER(G{r})),C{r}*G{r}/1000,"")'
    rows = [r[:6] + [g, h, i] + r[6:] for r in rows_ex] if EXAMPLES else []
    rows += [[None] * 6 + [g, h, i, None] for _ in range(n_blank)]
    hdr = ["Source", "Site / category", "Activity quantity", "Unit", "Data ref (GHG-03)", "Factor ID (GHG-02)",
           "Factor", "Factor unit", "tCO2e", "Notes / adjustments"]
    end = table(ws, 9, hdr, rows, [30, 18, 14, 8, 14, 14, 10, 14, 11, 40], input_cols=[1, 2, 3, 4, 5, 6, 10],
                example_rows=len(rows_ex), fmt={3: NUM0, 7: "0.0000", 9: NUM})
    total_row(ws, end + 1, 8, "Total tCO2e", [9], 10, end)
    return ws, end


def ghg_04(wb, a):
    rows = [["Diesel – backup generators", "Phnom Penh SEZ", 50000, "L", "GHG-03b", "EF-DIESEL", "Consumption from tank stock movement"],
            ["Company vans", "Fleet", 60000, "km", "DA05", "EF-VAN", ""],
            ["R-410A recharged (screening)", "Chillers", 20, "kg", "DA08", "GWP-R410A", "AR6 2,256; calculator AR4 = 41.76 t (+3.36 t adjustment)"],
            ["LPG – canteen", "Canteen", 3000, "L", "DA07", "EF-LPG", ""]]
    ws, end = calc_sheet(wb, a, "GHG-04", "Scope 1 calculations", "S2.29(a)(i); S2.B19–B22", "Direct emissions: fuels, vehicles, refrigerants (AR6), process/land use.", rows)
    return end + 1


def ghg_05(wb, a):
    ws = header(wb, "GHG-05", "Scope 2 calculations", "S2.29(a)(i), (v)", a,
                "Location-based Scope 2 (required) and contractual-instrument information (market-based memo).")
    g, h = lookup("F", "C"), lookup("F", "D")
    i = '=IF(AND(ISNUMBER(C{r}),ISNUMBER(G{r})),C{r}*G{r}/1000,"")'
    rows = ex([["EDC grid electricity", "Phnom Penh SEZ", 1200000, "kWh", "GHG-03a", "EF-ELEC-KH", g, h, i, "Prorated to calendar year"]],
              8, 10, {7: g, 8: h, 9: i})
    end = table(ws, 9, ["Source", "Site", "kWh", "Unit", "Data ref", "Factor ID", "Factor", "Factor unit", "tCO2e (location)", "Notes"],
                rows, [28, 18, 14, 8, 12, 14, 10, 14, 13, 30], input_cols=[1, 2, 3, 4, 5, 6, 10], fmt={3: NUM0, 7: "0.0000", 9: NUM})
    total_row(ws, end + 1, 8, "Scope 2 location-based", [9], 10, end)
    loc_total = end + 1
    r = end + 3
    ws.cell(r, 1, "Contractual instruments (S2.29(a)(v)) — market-based memo; Cambodia has no residual mix").font = F_BOLD
    mb = '=IF(AND(ISNUMBER(B{r}),ISNUMBER(C{r})),B{r}*C{r}/1000,"")'
    rows = ex([["Rooftop solar PPA (from Q3 2027)", 0, 0, mb, "PPA contract; I-REC redemption", "Not yet operational"]], 4, 6, {4: mb})
    e2 = table(ws, r + 1, ["Instrument", "kWh covered", "Contract factor (kg CO2e/kWh)", "tCO2e", "Evidence", "Notes"], rows,
               None, input_cols=[1, 2, 3, 5, 6], fmt={2: NUM0, 4: NUM})
    return loc_total


SCOPE3_CATS = ["1 Purchased goods and services", "2 Capital goods", "3 Fuel- and energy-related", "4 Upstream transport",
               "5 Waste generated in operations", "6 Business travel", "7 Employee commuting", "8 Upstream leased assets",
               "9 Downstream transport", "10 Processing of sold products", "11 Use of sold products",
               "12 End-of-life of sold products", "13 Downstream leased assets", "14 Franchises", "15 Investments (financed)"]


def ghg_06(wb, a):
    ws = header(wb, "GHG-06", "Scope 3 screening and calculations", "S2.29(a)(vi); S2.B32–B57; S2.C3(b)", a,
                "Screen all 15 categories, then calculate relevant ones; record method and data quality.")
    scr_ex = {
        "1 Purchased goods and services": ["Y", "~75% of inventory; buyer demand", "Average-data (kg) → supplier EPDs", "3 Generic", "Included"],
        "3 Fuel- and energy-related": ["Y", "WTT of diesel; T&D", "Average-data", "2", "Included (WTT)"],
        "5 Waste generated in operations": ["Y", "Offcuts, wastewater", "Waste-type", "2", "Included"],
        "7 Employee commuting": ["Y", "4,000 workers by truck/bus", "Survey distance", "1–2", "Included"],
        "11 Use of sold products": ["N", "Garments: no direct use-phase emissions", "—", "—", "Excluded"],
        "15 Investments (financed)": ["N", "No investments", "—", "—", "Excluded"],
    }
    rows = [[c] + (scr_ex.get(c, [None] * 5) if EXAMPLES else [None] * 5) for c in SCOPE3_CATS]
    end = table(ws, 9, ["Category", "Relevant? (Y/N)", "Rationale (size, influence, risk, buyer demand)", "Method",
                        "Data quality", "Included / excluded"], rows, [34, 10, 40, 30, 12, 16], input_cols=[2, 3, 4, 5, 6],
                example_rows=0)
    dv_list(ws, f"B10:B{end}", ["Y", "N"])
    r = end + 2
    ws.cell(r, 1, "Calculations (Category column must use the exact category text above)").font = F_BOLD
    g, h = lookup("F", "C"), lookup("F", "D")
    i = '=IF(AND(ISNUMBER(C{r}),ISNUMBER(G{r})),C{r}*G{r}/1000,"")'
    exs = [["Purchased fabric", "1 Purchased goods and services", 300000, "kg", "DA06", "EF-FABRIC", "Placeholder factor — obtain EPDs"],
           ["Water supply", "1 Purchased goods and services", 20000, "m3", "DA14", "EF-WATER-S", ""],
           ["Diesel well-to-tank", "3 Fuel- and energy-related", 50000, "L", "GHG-03b", "EF-DIESEL-WTT", ""],
           ["Wastewater treatment", "5 Waste generated in operations", 18000, "m3", "DA14", "EF-WATER-T", ""],
           ["Fabric offcuts & general waste to landfill", "5 Waste generated in operations", 100, "t", "DA15", "EF-WASTE-LF", ""],
           ["Worker transport by bus/truck", "7 Employee commuting", 400000, "pkm", "DA11", "EF-BUS", "Survey response 72%"]]
    rows = [e[:6] + [g, h, i, e[6]] for e in exs] if EXAMPLES else []
    rows += [[None] * 6 + [g, h, i, None] for _ in range(20)]
    first = r + 2
    e2 = table(ws, r + 1, ["Source", "Category", "Activity quantity", "Unit", "Data ref", "Factor ID", "Factor", "Factor unit", "tCO2e", "Notes"],
               rows, None, input_cols=[1, 2, 3, 4, 5, 6, 10], example_rows=len(exs), fmt={3: NUM0, 7: "0.0000", 9: NUM})
    total_row(ws, e2 + 1, 8, "Scope 3 total", [9], first, e2)
    return first, e2


def ghg_00(wb, a, s1_tot_row, s2_tot_row, s3_first, s3_last):
    ws = header(wb, "GHG-00", "GHG lead schedule", "S2.29(a)(i)–(vi)", a,
                "Totals by scope and category linked to calculation WPs; prior year; change; intensity; checks. Feeds the disclosure note.")
    rows = [["Scope 1 (gross)", f"='GHG-04'!I{s1_tot_row}", None, '=IF(AND(ISNUMBER(B{r}),ISNUMBER(C{r}),C{r}<>0),B{r}/C{r}-1,"")', "GHG-04"],
            ["Scope 2 location-based", f"='GHG-05'!I{s2_tot_row}", None, '=IF(AND(ISNUMBER(B{r}),ISNUMBER(C{r}),C{r}<>0),B{r}/C{r}-1,"")', "GHG-05"]]
    for c in SCOPE3_CATS:
        rows.append([f"Scope 3 – Cat {c}",
                     f"=SUMIF('GHG-06'!$B${s3_first}:$B${s3_last},\"{c}\",'GHG-06'!$I${s3_first}:$I${s3_last})",
                     None, '=IF(AND(ISNUMBER(B{r}),ISNUMBER(C{r}),C{r}<>0),B{r}/C{r}-1,"")', "GHG-06"])
    end = table(ws, 9, ["Line", "Current year tCO2e", "Prior year tCO2e", "Change %", "WP ref"], rows,
                [40, 18, 18, 12, 12], input_cols=[3], example_rows=0, fmt={2: NUM, 3: NUM, 4: "0.0%"})
    s3a, s3b = 12, end
    r = end + 1
    labels = [("Scope 3 total", f"=SUM(B{s3a}:B{s3b})", f"=SUM(C{s3a}:C{s3b})"),
              ("Total Scope 1 + 2 (location)", "=B10+B11", "=C10+C11"),
              ("Total Scope 1 + 2 + 3 (disclosed scopes)", f"=B10+B11+B{r}", f"=C10+C11+C{r}")]
    for k, (lab, f1, f2) in enumerate(labels):
        rr = r + k
        ws.cell(rr, 1, lab).font = F_BOLD
        for col, f in ((2, f1), (3, f2)):
            c = ws.cell(rr, col, f)
            c.number_format, c.font, c.fill, c.border = NUM, F_BOLD, FILL_TOT, BORDER
        c = ws.cell(rr, 4, f'=IF(AND(ISNUMBER(B{rr}),C{rr}<>0),B{rr}/C{rr}-1,"")')
        c.number_format = "0.0%"
    tot = r + 2
    r = tot + 2
    ws.cell(r, 1, "Intensity").font = F_BOLD
    inten = [["Revenue (presentation currency, from FS)", 400000, "FS income statement"],
             ["Employees (average FTE)", 4000, "Payroll"]]
    e2 = table(ws, r + 1, ["Denominator", "Value", "Source"], inten, None, input_cols=[2, 3], example_rows=2, fmt={2: NUM0})
    ws.cell(e2 + 1, 1, "tCO2e (S1+2+3) per unit of revenue").font = F_BOLD
    ws.cell(e2 + 1, 2, f'=IF(ISNUMBER(B{r+2}),B{tot}/B{r+2},"")').number_format = "0.0000"
    ws.cell(e2 + 2, 1, "tCO2e (S1+2+3) per employee").font = F_BOLD
    ws.cell(e2 + 2, 2, f'=IF(ISNUMBER(B{r+3}),B{tot}/B{r+3},"")').number_format = "0.000"
    r = e2 + 4
    ws.cell(r, 1, "Checks").font = F_BOLD
    checks = [("Scope 3 lead = GHG-06 total", f"=ROUND(B{tot-2}-'GHG-06'!I{s3_last+1},6)=0"),
              ("No calculation line without a factor (GHG-04)", f"=COUNTIFS('GHG-04'!F10:F200,\"<>\",'GHG-04'!G10:G200,\"\")=0"),
              ("No calculation line without a factor (GHG-06)", f"=COUNTIFS('GHG-06'!F{s3_first}:F{s3_last},\"<>\",'GHG-06'!G{s3_first}:G{s3_last},\"\")=0"),
              ("Market-based Scope 2 not used in place of location-based", "Manual – confirm"),
              ("Carbon credits not netted (S2.36(e))", "Manual – confirm")]
    for k, (lab, f) in enumerate(checks):
        ws.cell(r + 1 + k, 1, lab).font = F_BASE
        ws.cell(r + 1 + k, 2, f).font = F_BASE
    ws.cell(r + 7, 1, "Disaggregate Scope 1 and 2 between consolidated group and other investees (S2.29(a)(iv)); "
                      "state consolidation approach, GHG Protocol standards, GWP source, factor sources/years, exclusions, estimates and restatements.").font = F_NOTE


def ghg_01(wb, a):
    ws = header(wb, "GHG-01", "Boundary and approach memo", "S2.29(a)(ii)–(iii); S1.20", a,
                "Document the organisational and operational boundary and measurement approach.")
    q = [["Consolidation approach (equity share / financial control / operational control) and reason", "Operational control — aligns with management of energy use", "GHG Protocol ch.3"],
         ["Sites included (agree to FS consolidation list and site register)", None, None],
         ["Exclusions and rationale (de minimis must be justified)", None, None],
         ["Investees for S2.29(a)(iv) disaggregation (associates, JVs, unconsolidated subs)", None, None],
         ["Scope 2 method: location-based (required) + contractual instruments", None, None],
         ["GWP basis (AR6 100-yr unless jurisdictional relief)", None, None],
         ["Measurement standard (GHG Protocol 2004 or jurisdictional method — Dec 2025 relief)", None, None],
         ["Biogenic CO2 treatment (reported separately)", None, None],
         ["Base year and recalculation policy (structural change, method change, errors)", None, None],
         ["Estimates and limitations; reliefs used (S2.C3)", None, None]]
    table(ws, 9, ["Item", "Conclusion", "Evidence / ref"], q, [62, 56, 22], input_cols=[2, 3], height=30)


def ghg_03(wb, a):
    ws = header(wb, "GHG-03", "Source-data register", "S2.B39–B54; S1.77–82", a,
                "One line per activity datum: document, period, owner, data-quality tier, evidence reference.")
    rows = ex([["DA01", "S1", "Diesel generators", "Phnom Penh SEZ", "FY2026", 50000, "L", "Tank dips + invoices (GHG-03b)", "Facilities", "1 Measured", 14, "ENV-DSL-2026-001"],
               ["DA04", "S2", "EDC electricity", "Phnom Penh SEZ", "FY2026", 1200000, "kWh", "EDC bills (GHG-03a)", "Finance", "1 Measured", 15, "ENV-ELEC-2026-001"]],
              30, 12)
    end = table(ws, 9, ["Data ID", "Scope", "Activity", "Site", "Period", "Quantity", "Unit", "Source document", "Owner",
                        "Quality tier", "EDQA score /15", "Evidence ref"], rows, [8, 8, 24, 16, 9, 13, 7, 30, 12, 12, 10, 18],
                input_cols=list(range(1, 13)), example_rows=2, fmt={6: NUM0})
    dv_list(ws, f"J10:J{end}", ["1 Measured", "2 Specific proxy", "3 Generic"])


def ghg_03a(wb, a):
    ws = header(wb, "GHG-03a", "Invoice and expense register (cut-off proration)", "S1.77–82 (cut-off control)", a,
                "One line per utility/fuel/service invoice; prorate billing periods to the reporting period; tie to GL in GHG-07b.")
    try:
        ye = dt.date.fromisoformat(a.year_end)
        ys = dt.date(ye.year - 1, ye.month, ye.day) + dt.timedelta(days=1)
    except ValueError:
        ye, ys = None, None
    ws["J2"], ws["J3"] = "Period start", "Period end"
    ws["K2"], ws["K3"] = ys, ye
    for c in ("K2", "K3"):
        ws[c].fill, ws[c].number_format, ws[c].border = FILL_INPUT, "yyyy-mm-dd", BORDER
    ws["J2"].font = ws["J3"].font = F_BOLD
    days = '=IF(AND(ISNUMBER(D{r}),ISNUMBER(E{r})),E{r}-D{r}+1,"")'
    inp = '=IF(AND(ISNUMBER(D{r}),ISNUMBER(E{r})),MAX(0,MIN(E{r},$K$3)-MAX(D{r},$K$2)+1),"")'
    pro = '=IF(AND(ISNUMBER(G{r}),ISNUMBER(H{r}),F{r}>0),H{r}*G{r}/F{r},"")'
    exrows = [["EDC-0112", "EDC", "Electricity", dt.date(2025, 12, 16), dt.date(2026, 1, 15), days, inp, 100000, "kWh", pro, 70.0, "6061 Electricity"],
              ["EDC-1212", "EDC", "Electricity", dt.date(2026, 12, 16), dt.date(2027, 1, 15), days, inp, 104000, "kWh", pro, 72.8, "6061 Electricity"]]
    rows = ex(exrows, 40, 12, {6: days, 7: inp, 10: pro})
    end = table(ws, 9, ["Invoice no.", "Supplier", "Type", "Bill start", "Bill end", "Days in bill", "Days in period",
                        "Quantity", "Unit", "Quantity in period", f"Amount ({a.currency})", "GL account"], rows,
                [12, 12, 12, 11, 11, 9, 10, 12, 7, 13, 13, 18], input_cols=[1, 2, 3, 4, 5, 8, 9, 11, 12], example_rows=2,
                fmt={4: "yyyy-mm-dd", 5: "yyyy-mm-dd", 8: NUM0, 10: NUM0, 11: NUM})
    total_row(ws, end + 1, 9, "Totals", [10, 11], 10, end)
    ws.cell(end + 3, 1, "Use meter-read dates, not calendar months. Accrue the last bill straddling year end. VAT excluded from Amount.").font = F_NOTE


def ghg_03b(wb, a):
    ws = header(wb, "GHG-03b", "Stock and logs (consumption and reasonableness)", "S1.77–82 (completeness)", a,
                "Derive fuel consumption from stock movement and test against run-hours / odometer.")
    cons = '=IF(COUNT(B{r}:D{r})=3,B{r}+C{r}-D{r},"")'
    exp = '=IF(AND(ISNUMBER(F{r}),ISNUMBER(G{r})),F{r}*G{r},"")'
    diff = '=IF(AND(ISNUMBER(E{r}),ISNUMBER(H{r}),H{r}<>0),E{r}/H{r}-1,"")'
    rows = ex([["Diesel – generators (L)", 5000, 52000, 7000, cons, 2000, 25, exp, diff, "Within ±5%"]], 8, 10,
              {5: cons, 8: exp, 9: diff})
    end = table(ws, 9, ["Fuel / log", "Opening stock", "Purchases", "Closing stock", "Consumption", "Run-hours / km",
                        "Burn rate (L/h or L/km)", "Expected consumption", "Difference %", "Conclusion"], rows,
                [24, 12, 12, 12, 12, 12, 12, 13, 11, 24], input_cols=[1, 2, 3, 4, 6, 7, 10], fmt={2: NUM0, 3: NUM0, 4: NUM0, 5: NUM0, 6: NUM0, 8: NUM0, 9: "0.0%"})
    ws.conditional_formatting.add(f"I10:I{end}", CellIsRule(operator="notBetween", formula=["-0.05", "0.05"], fill=PatternFill("solid", fgColor="F8CBAD")))


def ghg_07(wb, a):
    ws = header(wb, "GHG-07", "Three-way reconciliation and recalculation (Case F)", "S1.77–86; ISSA 5000", a,
                "Reconcile invoices, register and operational records; recalculate with the approved factor; evaluate misstatement.")
    lab = [("EDC invoices (kWh)", 120000), ("Monthly register (kWh)", 115000), ("Manager spreadsheet (kWh)", 123000),
           ("Factor used in report (kg CO2e/kWh)", 0.50), ("Reported Scope 2 (tCO2e)", 60), ("Review threshold", 0.05)]
    r0 = 10
    for k, (l, v) in enumerate(lab):
        ws.cell(r0 + k, 1, l).font = F_BOLD
        c = ws.cell(r0 + k, 2, v if EXAMPLES else None)
        c.fill, c.border = FILL_INPUT, BORDER
        c.number_format = "0.0000" if k == 3 else ("0%" if k == 5 else NUM0)
    ws.cell(r0 + 6, 1, "Approved factor (GHG-02 EF-ELEC-KH)").font = F_BOLD
    ws.cell(r0 + 6, 2, f"=INDEX('GHG-02'!$C${FR_FIRST}:$C${FR_LAST},MATCH(\"EF-ELEC-KH\",'GHG-02'!$A${FR_FIRST}:$A${FR_LAST},0))").font = F_LINK
    calc = [("Invoices vs register (kWh)", "=B10-B11", NUM0),
            ("Manager sheet vs invoices (kWh)", "=B12-B10", NUM0),
            ("Total range (kWh)", "=MAX(B10:B12)-MIN(B10:B12)", NUM0),
            ("Range as % of invoices", '=IF(B10>0,B20/B10,"")', "0.0%"),
            ("Escalate? (range > threshold)", '=IF(B21="","",IF(B21>B15,"Yes – escalate","No"))', "General"),
            ("Recalculated Scope 2 (invoices × approved factor)", "=B10*B16/1000", NUM),
            ("Misstatement (recalculated − reported)", "=B23-B14", NUM),
            ("% of reported figure", '=IF(B14<>0,B24/B14,"")', "0.0%"),
            ("% of corrected figure", '=IF(B23<>0,B24/B23,"")', "0.0%"),
            ("Effect of kWh range at approved factor (tCO2e)", "=B20*B16/1000", NUM)]
    for k, (l, f, nf) in enumerate(calc):
        rr = 18 + k
        ws.cell(rr, 1, l).font = F_BASE
        c = ws.cell(rr, 2, f)
        c.number_format = nf
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 18
    ws.cell(29, 1, "Conclusion and actions").font = F_BOLD
    steps = ["1 Do not pick a figure — reconcile all three records to bills, meter-read dates and GL (GHG-07b).",
             "2 Confirm boundary and meter coverage; exclude out-of-boundary sub-meters; move generator energy to Scope 1.",
             "3 Replace the prior-year factor with the approved current factor; record source and year (GHG-02).",
             "4 Recalculate; evaluate misstatement of the metric (factor effect vs kWh effect).",
             "5 Independent review and approval; exception log (S1-14); remediation plan to CFO and audit committee."]
    for k, s in enumerate(steps):
        ws.cell(30 + k, 1, s).font = F_BASE


def ghg_07b(wb, a):
    ws = header(wb, "GHG-07b", "Invoice register vs general ledger", "S1.23; S1.77–82", a,
                "Agree invoice register amounts by GL account to the trial balance; explain reconciling items.")
    reg = "=SUMIF('GHG-03a'!$L$10:$L$60,A{r},'GHG-03a'!$K$10:$K$60)"
    diff = '=IF(ISNUMBER(B{r}),B{r}-C{r},"")'
    unexp = '=IF(ISNUMBER(D{r}),D{r}-IF(ISNUMBER(E{r}),E{r},0),"")'
    rows = ex([["6061 Electricity", 150, reg, diff, 7.2, unexp, "Accrual for Dec bill and late-posted invoice"]], 8, 7,
              {3: reg, 4: diff, 6: unexp})
    table(ws, 9, ["GL account (exact text as GHG-03a)", f"GL balance ({a.currency})", "Invoice register", "Difference",
                  "Explained reconciling items", "Unexplained", "Explanation"], rows, [30, 16, 16, 12, 16, 12, 40],
          input_cols=[1, 2, 5, 7], fmt={2: NUM, 3: NUM, 4: NUM, 5: NUM, 6: NUM})


def ghg_08(wb, a):
    ws = header(wb, "GHG-08", "Review checklist, EDQA scoring and exceptions", "S1.77–82; ISSA 5000", a,
                "Reviewer procedures (Lab 3) and data-quality scoring before consolidation.")
    chk = [["Agree every input to the source-data register and evidence (bill, log, invoice)"],
           ["Recalculate at least three lines by hand: AD × EF ÷ 1,000"],
           ["Check factor, year and source; flag DEFRA 2022 factors changed >5% (S1/S2) or >10% (S3)"],
           ["Refrigerants converted at AR6; calculator AR4 output adjusted"],
           ["Scope 2 location-based reported; contractual info separate"],
           ["Scope 3 screening complete for all 15 categories"],
           ["Lead schedule totals agree to calculation WPs and disclosure"],
           ["Invoice register reconciled to GL (GHG-07b); cut-off tested"],
           ["Prior-year comparatives consistent; restatements explained (S1.83–86, B52)"],
           ["Exceptions logged with owner and closed or disclosed"]]
    rows = [c + [None, None, None] for c in chk]
    end = table(ws, 9, ["Procedure", "Done (Y/N)", "Result / exception", "WP ref"], rows, [70, 10, 40, 12],
                input_cols=[2, 3, 4], example_rows=0)
    r = end + 2
    ws.cell(r, 1, "EDQA scoring (each 1–3; total < 10 → improvement action and estimate note)").font = F_BOLD
    tot = '=IF(COUNT(B{r}:F{r})=5,SUM(B{r}:F{r}),"")'
    act = '=IF(G{r}="","",IF(G{r}<10,"Action + estimate note","OK"))'
    rows = ex([["EDC electricity", 3, 3, 3, 3, 3, tot, act, ""],
               ["Generator diesel", 3, 3, 3, 3, 2, tot, act, ""],
               ["Worker commuting", 2, 1, 1, 2, 2, tot, act, "Repeat survey, 70% response target"],
               ["Purchased fabric", 1, 1, 1, 2, 1, tot, act, "Obtain supplier EPDs"]], 10, 9, {7: tot, 8: act})
    table(ws, r + 1, ["Data point", "Completeness", "Accuracy", "Timeliness", "Consistency", "Source reliability", "Total /15", "Flag", "Improvement action"],
          rows, None, input_cols=[1, 2, 3, 4, 5, 6, 9], example_rows=4)


def ghg_09(wb, a):
    ws = header(wb, "GHG-09", "Assurance evidence index (GHG)", "ISSA 5000; S2.29(a)", a,
                "Cross-reference every disclosed GHG figure to GHG-00–08 and source evidence.")
    rows = ex([["Climate note table 1", "Scope 2 location-based", "='GHG-00'!B11", "tCO2e", "GHG-05", "EDC invoices; ENV-ELEC-2026-001..004", "Y", "Limited"]], 25, 8)
    end = table(ws, 9, ["Disclosure ref", "Figure", "Value", "Unit", "WP", "Evidence", "Reconciled (Y/N)", "Assurance scope"], rows,
                [22, 30, 14, 9, 10, 40, 12, 14], input_cols=[1, 2, 4, 5, 6, 7, 8], fmt={3: NUM})
    dv_list(ws, f"H10:H{end}", ["Limited", "Reasonable", "No"])


# ----------------------------------------------------------------------------- climate (CR / SA)
def cr_00(wb, a):
    ws = header(wb, "CR-00", "Index, reliefs elected and amendments applied", "S2.C1–C5; S1.E3–E6; Dec 2025 amendments", a,
                "Record which reliefs and amendments apply this period — each must be disclosed.")
    rel = [["First annual reporting period under ISSB?", "S1.E1–E6", None, None],
           ["No comparatives (year 1)", "S1.E3", None, None],
           ["Report after FS (timing relief)", "S1.E4", None, None],
           ["Climate-first (only IFRS S2 in year 1)", "S1.E5", None, None],
           ["Continue non-GHG-Protocol method (year 1)", "S2.C3(a)", None, None],
           ["Omit Scope 3 incl. Cat 15 (year 1)", "S2.C3(b)", None, None],
           ["Dec 2025 amendments applied early (mandatory from 1 Jan 2027)?", "S2 amendments", None, None],
           ["  – Cat 15 limitation (derivatives, facilitated, insurance-associated)", "S2 amendments", None, None],
           ["  – Industry classification other than GICS", "S2 amendments", None, None],
           ["  – Jurisdictional measurement method relief", "S2 amendments", None, None],
           ["  – Jurisdictional GWP relief", "S2 amendments", None, None]]
    end = table(ws, 9, ["Relief / amendment", "Para", "Applied? (Y/N)", "Disclosure wording / ref"], rel, [62, 16, 12, 50],
                input_cols=[3, 4], example_rows=0)
    dv_list(ws, f"C10:C{end}", ["Y", "N", "NA"])
    r = end + 2
    sheets = ["CR-01", "CR-02", "SA", "CR-03", "CR-04", "CR-05", "CR-06", "CR-07", "GHG-00", "GHG-01", "GHG-02", "GHG-03", "GHG-03a", "GHG-03b",
              "GHG-04", "GHG-05", "GHG-06", "GHG-07", "GHG-07b", "GHG-08", "GHG-09", "FE-01", "S2-CL"]
    ws.cell(r, 1, "Working papers").font = F_BOLD
    for k, s in enumerate(sheets):
        c = ws.cell(r + 1 + k, 1, s)
        c.font = Font(name=FONT, size=10, color="0563C1", underline="single")
        c.hyperlink = f"#'{s}'!A1"


def cr_01(wb, a):
    ws = header(wb, "CR-01", "Climate-related risk and opportunity register", "S2.10–12; S2.24–26", a,
                "Classify CRROs (physical acute/chronic; transition policy/technology/market/reputation; opportunities), assess and link to responses and metrics.")
    sc = '=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r})),H{r}*I{r},"")'
    rows = ex([["R01", "Buyer decarbonisation requirements", "Transition – market", "Downstream", "Medium (1–5y)", "Orders from EU/US brands", "Revenue; access to green loans", 4, 4, sc, "Supplier data programme; solar PPA", "Scope 3 Cat 1; % revenue from buyers with SBTs", "SA6"],
               ["R02", "Flooding at Phnom Penh SEZ site", "Physical – acute", "Own operations", "Short (0–1y)", "Factory, inventory", "PP&E, inventory, downtime", 3, 4, sc, "Raised plinths; insurance", "% assets in flood zone (S2.29(c))", "SA5"],
               ["R04", "Heat stress reduces productivity", "Physical – chronic", "Own operations", "Long (>5y)", "4,000 workers", "Labour cost, output", 4, 3, sc, "Ventilation, work-rest schedule", "Hours lost to heat", "SA5"],
               ["R05", "Domestic carbon pricing / ETS (sub-decree pending) and pass-through in power tariffs", "Transition – policy/legal", "Own operations; upstream power", "Medium (1–5y)", "Scope 1 and purchased electricity", "Energy cost; compliance cost", 3, 2, sc, "Solar PPA; efficiency; monitor MoE rules", "CR-06 cost by scenario; internal carbon price (S2.29(f))", "SA6; CR-06"],
               ["O01", "Rooftop solar lowers energy cost and Scope 2", "Opportunity – energy source", "Own operations", "Medium", "Energy cost", "Cost of sales", 4, 2, sc, "2 MW PPA by Q3 2027", "% renewable; capex (S2.29(e))", "SA6"]], 20, 13, {10: sc})
    end = table(ws, 9, ["ID", "Description", "Type", "Value-chain location", "Horizon", "Exposure (assets/activities)", "Financial channel",
                        "Likelihood 1–5", "Magnitude 1–5", "Score", "Response (mitigation/adaptation)", "Metric / target", "Scenario WP"],
                rows, [7, 32, 20, 16, 12, 22, 22, 9, 9, 8, 30, 28, 9], input_cols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13], example_rows=5)
    dv_list(ws, f"C10:C{end}", ["Physical – acute", "Physical – chronic", "Transition – policy/legal", "Transition – technology",
                                 "Transition – market", "Transition – reputation", "Opportunity – resource efficiency",
                                 "Opportunity – energy source", "Opportunity – products/services", "Opportunity – markets", "Opportunity – resilience"])
    ws.conditional_formatting.add(f"J10:J{end}", CellIsRule(operator="greaterThanOrEqual", formula=["12"], fill=PatternFill("solid", fgColor="F8CBAD")))


def cr_02(wb, a):
    ws = header(wb, "CR-02", "Strategy, transition plan and financial effects", "S2.13–21", a,
                "Business model effects, strategy and transition plan, resourcing, progress, and financial effects.")
    q = [["Effects on business model and value chain; concentration", "13(a)–(b)"],
         ["Changes to business model and resource allocation", "14(a)(i)"],
         ["Direct mitigation and adaptation", "14(a)(ii)"],
         ["Indirect mitigation and adaptation (suppliers, customers)", "14(a)(iii)"],
         ["Transition plan: key assumptions and dependencies (label as plan, not result)", "14(a)(iv)"],
         ["How targets (incl. GHG) will be achieved", "14(a)(v)"],
         ["Resourcing (capex, opex, financing — e.g. green loan)", "14(b)"],
         ["Progress on previously disclosed plans", "14(c)"]]
    rows = [x + [None, None] for x in q]
    end = table(ws, 9, ["Item", "S2 para", "Entity disclosure points", "Evidence / WP"], rows, [52, 12, 60, 18], input_cols=[3, 4], example_rows=0, height=30)
    r = end + 2
    ws.cell(r, 1, f"Financial effects ({a.currency})").font = F_BOLD
    rows = ex([["R02", "PP&E – SEZ factory (Note 10)", 0, "N", 50, 120, 300, "Range", "Long: high uncertainty (S2.19)", "SA5; S1-09"]], 12, 10)
    e2 = table(ws, r + 1, ["CRRO", "FS line item / note", "Current-period effect", "Significant risk of material adjustment next year? (Y/N)",
                           "Anticipated – short", "Anticipated – medium", "Anticipated – long", "Quant / qual / range",
                           "Reason if not quantitative (S2.19–21)", "Link"], rows, None, input_cols=list(range(1, 11)),
               fmt={3: NUM0, 5: NUM0, 6: NUM0, 7: NUM0})


SA_SECTIONS = [
    ("SA1 Purpose, scope and proportionality (S2.22; B1–B8)", ["Exposure assessment (high/medium/low)", "Skills, capabilities, resources", "Approach chosen (qualitative / quantitative / mixed) and why", "Moving up: plan to quantitative (B17)"]),
    ("SA2 Scenario selection (S2.22(b)(i)(1)–(5))", None),
    ("SA3 Time horizons and operational scope (S2.22(b)(i)(6)–(7))", ["Horizons (consistent with CR-01 / S1-02)", "Own operations covered", "Upstream suppliers covered", "Downstream customers/clients covered"]),
    ("SA4 Key assumptions (S2.22(b)(ii))", ["Climate policies in jurisdictions (NDC 3.0, carbon pricing, buyer/EU rules)", "Macroeconomic trends", "National/regional variables (Mekong floods, heat, land use, infrastructure)", "Energy usage and mix (hydro, coal, solar; grid factor path)", "Technology developments"]),
    ("SA5 Physical risk analysis (hazard × exposure × vulnerability)", ["Hazards and data sources", "Exposed assets/sites/portfolios", "Vulnerability and adaptation", "Results by scenario and horizon"]),
    ("SA6 Transition risk analysis", ["Carbon price path and pass-through", "Demand / buyer requirements", "Technology and energy costs", "Results by scenario and horizon"]),
    ("SA7 Financial effects (S2.22(a)(i); S2.15–21)", None),
    ("SA8 Resilience and capacity to adjust (S2.22(a)(iii))", ["Availability and flexibility of financial resources", "Ability to redeploy, repurpose, upgrade or decommission assets", "Effect of current and planned investments", "Resilience conclusion"]),
    ("SA9 Uncertainties, quality rating and update (S2.22(a)(ii), (b)(iii); B18)", ["Significant areas of uncertainty", "Date of analysis / reporting period", "Next full update (strategic cycle) — resilience reassessed annually", "Two actions to move up one quality level"]),
]


def sa(wb, a):
    ws = header(wb, "SA", "Climate scenario analysis SA1–SA9", "S2.22; S2.B1–B18", a,
                "Proportionate scenario analysis and resilience assessment; quality checklist and ladder rating.")
    ws.column_dimensions["A"].width = 52
    for col, w in zip("BCDEFG", (22, 22, 22, 22, 22, 22)):
        ws.column_dimensions[col].width = w
    r = 9
    for title, items in SA_SECTIONS:
        ws.cell(r, 1, title).font = Font(name=FONT, size=11, bold=True, color="1F3864")
        r += 1
        if title.startswith("SA2"):
            rows = ex([["NGFS Net Zero 2050", "NGFS Phase V", "~1.5°C", "Transition + physical", "Paris-aligned (S2.22(b)(i)(4))", "Buyer and policy pressure fastest"],
                       ["NGFS Current Policies", "NGFS Phase V", "~3°C", "Physical + transition", "Hot-house", "Floods and heat most severe"]], 3, 6)
            r = table(ws, r, ["Scenario", "Source", "Temperature outcome", "Risk type(s)", "Role in diverse range", "Why relevant to entity"], rows,
                      input_cols=list(range(1, 7)), example_rows=2) + 2
        elif title.startswith("SA7"):
            rows = ex([["Revenue (lost orders)", -5, -12, -2, -6, "NZ2050 hits transition; CP hits physical"],
                       ["PP&E impairment (flood)", 0, -30, -10, -80, "Linked to S1-09 IAS 36"]], 5, 6)
            e = table(ws, r, [f"Line item ({a.currency})", "NZ2050 – medium", "NZ2050 – long", "Current Policies – medium", "Current Policies – long", "Comment"],
                      rows, input_cols=list(range(1, 7)), example_rows=2, fmt={2: NUM0, 3: NUM0, 4: NUM0, 5: NUM0})
            total_row(ws, e + 1, 1, "Total", [2, 3, 4, 5], r + 1, e, NUM0)
            r = e + 3
        else:
            rows = [[it, None] for it in items]
            r = table(ws, r, ["Item", "Response / reference"], rows, input_cols=[2], example_rows=0, height=30) + 2
    ws.cell(r, 1, "Quality checklist (Y/N)").font = Font(name=FONT, size=11, bold=True, color="1F3864")
    chk = ["Physical and transition risks covered", "≥1 scenario <2°C and ≥1 >2°C per risk type", "Named sources and relevance stated",
           "Scope: own ops, upstream, downstream", "Horizons stated and consistent with register", "Key assumptions disclosed",
           "Results translated into financial effects and strategy", "Capacity to adjust assessed", "Significant uncertainties named",
           "Date of analysis and next update stated"]
    first = r + 2
    e = table(ws, r + 1, ["Checklist item", "Met? (Y/N)"], [[c, None] for c in chk], input_cols=[2], example_rows=0)
    dv_list(ws, f"B{first}:B{e}", ["Y", "N"])
    ws.cell(e + 1, 1, "Items met").font = F_BOLD
    ws.cell(e + 1, 2, f'=COUNTIF(B{first}:B{e},"Y")')
    ws.cell(e + 2, 1, "Quality-ladder level (0–4) — enter after review").font = F_BOLD
    ws.cell(e + 2, 2).fill = FILL_INPUT
    ws.cell(e + 3, 1, "Ladder: 0 none; 1 assessment, no scenarios; 2 single-dimension; 3 limited (≥1 physical + ≥1 transition); "
                      "4 comprehensive (each risk type with one scenario >2°C and one <2°C).").font = F_NOTE


def cr_03(wb, a):
    ws = header(wb, "CR-03", "Cross-industry metrics S2.29(b)–(g)", "S2.29(b)–(g)", a,
                "Amounts and percentages vulnerable/aligned, capital deployment, internal carbon price, remuneration.")
    pct = '=IF(AND(ISNUMBER(D{r}),ISNUMBER(E{r}),E{r}<>0),D{r}/E{r},"")'
    rows = [["29(b)", "Assets or business activities vulnerable to transition risks", "Revenue from buyers with supplier-decarbonisation requirements", 260000 if EXAMPLES else None, 400000 if EXAMPLES else None, pct, "Sales ledger by customer"],
            ["29(c)", "Assets or business activities vulnerable to physical risks", "Carrying amount of PP&E in 1-in-100-year flood zone", None, None, pct, None],
            ["29(d)", "Assets or business activities aligned with climate-related opportunities", None, None, None, pct, None],
            ["29(e)", "Capital expenditure, financing or investment deployed towards CRROs", None, None, None, pct, None],
            ["29(f)(i)", "Internal carbon price — whether and how applied in decision-making", None, None, None, None, None],
            ["29(f)(ii)", "Internal carbon price — price per tCO2e", None, None, None, None, None],
            ["29(g)(i)", "Whether and how climate considerations factored into executive remuneration", None, None, None, None, None],
            ["29(g)(ii)", "% of executive management remuneration recognised in period linked to climate", None, None, None, pct, None]]
    table(ws, 9, ["Para", "Metric", "Definition / basis used", f"Amount ({a.currency})", "Denominator (total)", "%", "Source / WP"], rows,
          [10, 52, 44, 16, 16, 9, 24], input_cols=[3, 4, 5, 7], example_rows=1, fmt={4: NUM0, 5: NUM0, 6: "0.0%"})


def cr_04(wb, a):
    ws = header(wb, "CR-04", "Climate targets and carbon credits", "S2.33–36", a,
                "Record every climate target with all S2.33–36 elements; separate gross and net; credits disclosed, never netted in the inventory.")
    prog = '=IF(AND(ISNUMBER(H{r}),ISNUMBER(J{r}),ISNUMBER(K{r}),J{r}<>H{r}),(K{r}-H{r})/(J{r}-H{r}),"")'
    rows = ex([["T01", "Scope 1+2 (location) tCO2e", "Mitigation", "All sites", "2026–2030", "Absolute", "FY2026", 904, "−15% by 2028", 768, 904, prog,
                "Gross", "CO2, CH4, N2O, HFCs", "S1, S2", "Informed by NDC 3.0 and Paris", "Not validated", "None planned", "None"]], 10, 19, {12: prog})
    table(ws, 9, ["ID", "Metric", "Objective", "Part of entity", "Period", "Absolute / intensity", "Base period", "Base value",
                  "Milestones / interim", "Target value", "Current value", "Progress", "Gross / net", "Gases", "Scopes",
                  "Latest international agreement / NDC link", "Third-party validation (e.g. SBTi)", "Planned carbon credits (extent, scheme, type, integrity)", "Revisions"],
          rows, [6, 22, 12, 12, 10, 11, 9, 9, 14, 9, 9, 9, 8, 14, 9, 22, 16, 28, 12], input_cols=[i for i in range(1, 20) if i != 12],
          fmt={8: NUM0, 10: NUM0, 11: NUM0, 12: "0%"})


def cr_05(wb, a):
    ws = header(wb, "CR-05", "Industry-based metrics", "S2.32; Industry-based Guidance on Implementing IFRS S2", a,
                "Refer to and consider the Industry-based Guidance for the entity's industry; record metrics used or reasons for omission.")
    rows = ex([["Apparel, Accessories & Footwear", "Raw materials sourcing", "Description of environmental and social risks associated with sourcing priority raw materials (illustrative)", "Applicable", "Qualitative", "Procurement"]], 12, 6)
    table(ws, 9, ["Industry", "Disclosure topic", "Metric", "Applicable? / reason if not", "Value", "Source"], rows, [26, 24, 60, 24, 14, 16],
          input_cols=list(range(1, 7)))
    ws.cell(30, 1, "Check the current Industry-based Guidance code and wording; the ISSB is enhancing SASB Standards (exposure drafts 2025–2026).").font = F_NOTE


YEARS = list(range(2027, 2036))
PRICE_PATHS = [  # illustrative USD/tCO2e assumptions, NOT forecasts
    ("Current Policies (~3°C) — no domestic ETS before 2030", [0, 0, 0, 0, 5, 5, 6, 7, 8]),
    ("NDC 3.0 / Delayed transition — ETS from 2029", [0, 0, 5, 7, 9, 11, 13, 14, 15]),
    ("Net Zero 2050 (~1.5°C) — ETS from 2027, regional convergence", [10, 14, 18, 23, 28, 33, 39, 44, 50]),
]


def cr_06(wb, a, kwh_first, kwh_last):
    """Cambodia carbon pricing / ETS exposure model (scenario-based)."""
    ws = header(wb, "CR-06", "Cambodia carbon pricing and ETS exposure model", "S2.10, 15–22, 29(f), 36(e); S1.23", a,
                "Scenario model of a possible Cambodian ETS (or carbon tax): direct compliance cost if above the threshold, "
                "indirect cost via electricity pass-through, credit use. Parameters are assumptions — no Cambodian ETS is in force (Nov 2026).")
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 16
    for j in range(3, 3 + len(YEARS)):
        ws.column_dimensions[get_column_letter(j)].width = 11
    ws.column_dimensions[get_column_letter(3 + len(YEARS))].width = 14

    def sec(r, t):
        ws.cell(r, 1, t).font = Font(name=FONT, size=11, bold=True, color="1F3864")

    # A. Regulatory status
    sec(9, "A. Regulatory status — confirm each item before use (☐)")
    status = [["Domestic ETS or carbon tax in force?", "No (as understood at Nov 2026)", "Slaughter and May ESG in APAC 2025 (Cambodia)", "☐"],
              ["Carbon market sub-decree", "NDC 3.0 milestone: sub-decree approved by 2026 — status unconfirmed", "Cambodia NDC 3.0 (UNFCCC, Aug 2025)", "☐"],
              ["Article 6 framework", "MoE Article 6 Operations Manual (Dec 2023/2024); eligible mechanisms: JCM, Gold Standard, VCS, Article 6.4 (PACM)", "DFDL legal update; MoE", "☐"],
              ["Institution", "Carbon Crediting Secretariat (CCS) under MoE leads implementation", "Khmer Times, May 2026", "☐"],
              ["Bilateral", "JCM (Japan); Korea Article 6.2 partnership; Cambodia–Singapore MoU (2023)", "Press; MoE", "☐"],
              ["Registry", "UN PACM registry and independent-mechanism registries pending a national registry", "MoE Operations Manual", "☐"],
              ["External carbon prices faced by exporters", "EU CBAM (definitive from 2026: cement, iron/steel, aluminium, fertilisers, hydrogen, electricity); buyer internal carbon prices", "EU Regulation 2023/956", "☐"]]
    lastc = 3 + len(YEARS)
    for j, h in enumerate(["Item", "Current understanding  [source]  — tick ☐ when confirmed"], start=1):
        c = ws.cell(10, j, h)
        c.font, c.fill, c.border = F_HEAD, FILL_HEAD, BORDER
    ws.merge_cells(start_row=10, start_column=2, end_row=10, end_column=lastc)
    for k, (item, und, src, tick) in enumerate(status):
        rr = 11 + k
        ws.cell(rr, 1, item).font = F_BOLD
        ws.cell(rr, 1).alignment = WRAP
        c = ws.cell(rr, 2, f"{und}  [{src}]  {tick}")
        c.font, c.alignment, c.fill = F_BASE, WRAP, FILL_INPUT
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=lastc)
        ws.row_dimensions[rr].height = 30
    r = 10 + len(status)

    # B. Scheme design assumptions
    r += 2
    sec(r, "B. Scheme design assumptions (scenario inputs — edit; label as assumptions in any disclosure)")
    assum = [("Scheme modelled", "Hypothetical Cambodia ETS", "text"),
             ("Earliest possible compliance year (scenario start = later of this and first year its price > 0)", 2027, "0"),
             ("Coverage threshold — direct (Scope 1) emissions per entity/facility (tCO2e/yr)", 25000, NUM0),
             ("Power generation covered (cost passed through to tariffs)? 1 = yes, 0 = no", 1, "0"),
             ("Electricity cost pass-through to customers (%)", 0.5, "0%"),
             ("Free allocation in first compliance year (%)", 0.9, "0%"),
             ("Annual reduction in free allocation (percentage points)", 0.1, "0%"),
             ("Offset limit — share of obligation met with eligible Cambodian credits (%)", 0.1, "0%"),
             ("Eligible credit price (USD/tCO2e)", 8, NUM),
             ("Annual change in covered emissions after base year (%) — negative = reduction", -0.02, "0.0%"),
             ("Base year for emissions (from GHG-00)", 2026, "0"),
             ("Base-year Scope 1 (tCO2e) — link", "='GHG-00'!B10", NUM),
             ("Base-year purchased electricity (kWh) — link", f"=SUM('GHG-05'!C{kwh_first}:C{kwh_last})", NUM0),
             ("Grid emission factor (kg CO2e/kWh) — link GHG-02", f"=INDEX('GHG-02'!$C${FR_FIRST}:$C${FR_LAST},MATCH(\"EF-ELEC-KH\",'GHG-02'!$A${FR_FIRST}:$A${FR_LAST},0))", "0.0000"),
             ("Presentation units per USD (e.g. KHR million per USD = 0.004 at KHR 4,000)", 0.004, "0.000000"),
             ("Discount rate for NPV (%)", 0.1, "0.0%"),
             ("EBITDA, presentation currency (for materiality context)", 60000, NUM0)]
    A = {}
    for k, (lab, val, nf) in enumerate(assum):
        rr = r + 1 + k
        ws.cell(rr, 1, lab).font = F_BASE
        c = ws.cell(rr, 2, val)
        c.border = BORDER
        c.number_format = "General" if nf == "text" else nf
        if is_formula(val):
            c.font = F_LINK
        else:
            c.fill = FILL_INPUT
            c.font = F_EX if EXAMPLES else F_BASE
        A[k] = f"$B${rr}"
    (sch, start, thr, pwr, pt, free0, freed, olim, cprice, chg, base, s1, kwh, gef, fx, disc, ebitda) = [A[k] for k in range(17)]
    r = r + len(assum) + 2

    # C. Price paths
    sec(r, "C. Carbon price paths (USD/tCO2e) — illustrative assumptions, not forecasts; align with SA2 scenarios")
    hdr_r = r + 1
    ws.cell(hdr_r, 1, "Scenario").font = F_HEAD
    ws.cell(hdr_r, 1).fill = FILL_HEAD
    ws.cell(hdr_r, 2, "Unit").font = F_HEAD
    ws.cell(hdr_r, 2).fill = FILL_HEAD
    for j, y in enumerate(YEARS):
        c = ws.cell(hdr_r, 3 + j, str(y))
        c.font, c.fill, c.alignment = F_HEAD, FILL_HEAD, Alignment(horizontal="center")
    price_rows = []
    for k, (name, path) in enumerate(PRICE_PATHS):
        rr = hdr_r + 1 + k
        ws.cell(rr, 1, name).font = F_BASE
        ws.cell(rr, 2, "USD/tCO2e").font = F_BASE
        for j, p in enumerate(path):
            c = ws.cell(rr, 3 + j, p)
            c.fill, c.border, c.number_format = FILL_INPUT, BORDER, NUM0
            c.font = F_EX if EXAMPLES else F_BASE
        price_rows.append(rr)
    r = hdr_r + len(PRICE_PATHS) + 2

    # D. Calculation per scenario
    sec(r, "D. Exposure by scenario and year")
    r += 1
    summary = []
    lastcol = get_column_letter(2 + len(YEARS))
    totcol = 3 + len(YEARS)
    for k, (name, _) in enumerate(PRICE_PATHS):
        pr = price_rows[k]
        for j, y in enumerate(YEARS):
            c = ws.cell(r, 3 + j, y)
            c.font, c.fill = F_BOLD, FILL_TOT
        ws.cell(r, totcol, "Total / NPV").font = F_BOLD
        yr = r
        c = ws.cell(yr, 2, f'=MAX({start},IF(_xlfn.MINIFS(C{yr}:{lastcol}{yr},C{pr}:{lastcol}{pr},">0")=0,9999,_xlfn.MINIFS(C{yr}:{lastcol}{yr},C{pr}:{lastcol}{pr},">0")))')
        c.font, c.fill, c.number_format = F_BOLD, FILL_TOT, "0"
        ws.cell(yr, 1, f"{name}  — scheme start year →").font = F_BOLD
        lines = [
            ("Covered direct emissions (tCO2e)", "={s1}*(1+{chg})^({Y}{yr}-{base})", NUM),
            ("In scheme? (year ≥ start and emissions ≥ threshold)", "=IF(AND({Y}{yr}>=$B${yr},{Y}{r1}>={thr}),1,0)", "0"),
            ("Free allocation (%)", "=IF({Y}{yr}<$B${yr},0,MAX(0,{free0}-{freed}*({Y}{yr}-$B${yr})))", "0%"),
            ("Net obligation after free allocation (tCO2e)", "={Y}{r1}*(1-{Y}{r3})*{Y}{r2}", NUM),
            ("Met with eligible credits (tCO2e)", "={Y}{r4}*{olim}", NUM),
            ("Allowances purchased (tCO2e)", "={Y}{r4}-{Y}{r5}", NUM),
            ("Direct compliance cost (USD)", "={Y}{r6}*{Y}{pr}+{Y}{r5}*{cprice}*IF({Y}{pr}>0,1,0)", NUM0),
            ("Indirect cost: electricity pass-through (USD)", "=IF({Y}{yr}>=$B${yr},{kwh}*{gef}/1000*{Y}{pr}*{pt}*{pwr},0)", NUM0),
            ("Total cost (USD)", "={Y}{r7}+{Y}{r8}", NUM0),
            ("Total cost (presentation currency)", "={Y}{r9}*{fx}", NUM),
            ("Total cost as % of EBITDA", "=IF({ebitda}<>0,{Y}{r10}/{ebitda},\"\")", "0.00%"),
        ]
        rows_idx = {}
        for i, (lab, tmpl, nf) in enumerate(lines, start=1):
            rr = yr + i
            rows_idx[f"r{i}"] = rr
            ws.cell(rr, 1, lab).font = F_BASE
        for i, (lab, tmpl, nf) in enumerate(lines, start=1):
            rr = yr + i
            for j in range(len(YEARS)):
                Y = get_column_letter(3 + j)
                f = tmpl.format(Y=Y, yr=yr, pr=pr, s1=s1, chg=chg, base=base, start=start, thr=thr, free0=free0,
                                freed=freed, olim=olim, cprice=cprice, kwh=kwh, gef=gef, pt=pt, pwr=pwr, fx=fx,
                                ebitda=ebitda, **rows_idx)
                c = ws.cell(rr, 3 + j, f)
                c.number_format, c.border, c.font = nf, BORDER, F_BASE
            if i in (6, 7, 8, 9):
                c = ws.cell(rr, totcol, f"=SUM(C{rr}:{lastcol}{rr})")
                c.number_format, c.font, c.fill = nf, F_BOLD, FILL_TOT
            if i == 10:
                c = ws.cell(rr, totcol, f"=NPV({disc},C{rr}:{lastcol}{rr})")
                c.number_format, c.font, c.fill = NUM, F_BOLD, FILL_TOT
        summary.append((name, yr + 9, yr + 10))
        r = yr + len(lines) + 2

    # E. Summary for disclosure
    sec(r, "E. Summary for S2.15–22 disclosure (anticipated financial effects of carbon-pricing transition risk)")
    r += 1
    rows = [[n, f"='CR-06'!{get_column_letter(totcol)}{tu}", f"='CR-06'!{get_column_letter(totcol)}{tp}",
             f"=MAX(C{tp}:{lastcol}{tp})"] for n, tu, tp in summary]
    rows = [[x[0], x[1].replace("'CR-06'!", ""), x[2].replace("'CR-06'!", ""), x[3]] for x in rows]
    e = table(ws, r, ["Scenario", "Cumulative cost 2027–2035 (USD)", "NPV (presentation currency)", "Peak annual cost (presentation currency)"],
              rows, None, example_rows=0, fmt={2: NUM0, 3: NUM, 4: NUM})
    notes = [
        "Disclosure links: transition risk — policy/legal (S2.10(b)); anticipated financial effects (S2.15–21, ranges acceptable); SA6 transition analysis (S2.22);",
        "internal carbon price — if the entity uses one of these paths in decisions, disclose how and the price (S2.29(f)); planned use of credits towards targets (S2.36(e)) — CR-07.",
        "Teaching point: an SME garment exporter with ~200 tCO2e Scope 1 falls far below a 25,000 t threshold — its exposure is indirect (power tariffs, buyer pricing, CBAM-type mechanisms), not direct compliance.",
        "Accounting (if a scheme starts): no specific IFRS for allowances (IFRIC 3 withdrawn) — choose and disclose a policy (e.g. allowances as intangible assets/inventory at cost; obligation as an IAS 37 liability",
        "measured at carrying amount of allowances held plus market price for the shortfall). Use the same price and FX assumptions as the FS (S1.23); connect to S1-09.",
    ]
    for k, n in enumerate(notes):
        ws.cell(e + 2 + k, 1, n).font = F_NOTE


def cr_07(wb, a):
    ws = header(wb, "CR-07", "Carbon credit register (Article 6 / voluntary)", "S2.36(e); Cambodia Article 6 Operations Manual", a,
                "Register credits held or planned: mechanism, authorisation, corresponding adjustment, type, intended use. Credits are disclosed separately — never netted against the GHG inventory.")
    rows = ex([["CC-01", "Rooftop solar / cookstove programme (illustrative)", "Gold Standard", "2027", 500, "Technological", "Avoidance",
                "N", "N", "GS registry", "Voluntary claim (not for net target)", "N", 8, 4000, "Planned — not yet contracted"],
               ["CC-02", "Paddy AWD methane reduction (illustrative)", "JCM", "2028", 300, "Nature-based", "Avoidance",
                "Y", "Y", "JCM registry", "Towards net Scope 1+2 target T02", "N", 12, 3600, "LoA from MoE required"]], 15, 15)
    end = table(ws, 9, ["ID", "Project", "Mechanism (JCM / Gold Standard / VCS / A6.4 PACM / A6.2)", "Vintage", "Quantity (tCO2e)",
                        "Nature-based / technological", "Removal / avoidance", "Letter of authorisation (MoE)?", "Corresponding adjustment?",
                        "Registry", "Intended use (S2.36(e))", "Retired?", "Price (USD/t)", "Cost (USD)", "Integrity notes (permanence, additionality, leakage, forestry export limits)"],
                rows, [7, 32, 20, 8, 11, 14, 12, 12, 12, 14, 26, 8, 9, 11, 40], input_cols=list(range(1, 16)), example_rows=2,
                fmt={5: NUM0, 13: NUM, 14: NUM0})
    dv_list(ws, f"F10:F{end}", ["Nature-based", "Technological"])
    dv_list(ws, f"G10:G{end}", ["Removal", "Avoidance"])
    total_row(ws, end + 1, 4, "Total", [5, 14], 10, end, NUM0)
    ws.cell(end + 3, 1, "Cambodian rules: credits from eligible mechanisms only; forestry credits counted toward the NDC are not exportable (company transfers via long-term contracts); "
                        "authorisation and corresponding adjustments determine whether a credit can support an international claim. Confirm against the current MoE Operations Manual / sub-decree.").font = F_NOTE


def fe_01(wb, a):
    ws = header(wb, "FE-01", "Financed emissions (PCAF) — Scope 3 Cat 15", "S2.29(a)(vi)(2); S2.B58–B63; Dec 2025 amendments", a,
                "Attribute counterparty emissions to the loan book by asset class and sector; score data quality.")
    attr = '=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r}),F{r}>0),E{r}/F{r},"")'
    fin = '=IF(AND(ISNUMBER(G{r}),ISNUMBER(H{r})),G{r}*H{r},"")'
    wq = '=IF(AND(ISNUMBER(E{r}),ISNUMBER(J{r})),E{r}*J{r},"")'
    rows = ex([["L-0001", "Business loans", "Brick kiln (manufacturing – clay products)", "In scope", 2.0, 5.0, attr, 3200, fin, 4, wq, "Revenue × sector factor"],
               ["L-0002", "Business loans", "Rice mill", "In scope", 1.5, 6.0, attr, 900, fin, 4, wq, ""],
               ["M-0001", "Mortgages", "Residential real estate", "In scope", 0.08, 0.12, attr, 2.5, fin, 5, wq, "Floor area proxy"]], 40, 12,
             {7: attr, 9: fin, 11: wq})
    end = table(ws, 9, ["Exposure ID", "Asset class", "Sector (state classification)", "Scope (in / excluded – reason)",
                        f"Outstanding ({a.currency})", "Denominator: EVIC / equity+debt / property value", "Attribution factor",
                        "Counterparty S1+2 tCO2e", "Financed tCO2e", "PCAF score 1–5", "Exposure × score", "Data source"], rows,
                [11, 16, 30, 18, 14, 18, 11, 14, 13, 10, 13, 22], input_cols=[1, 2, 3, 4, 5, 6, 8, 10, 12], example_rows=3,
                fmt={5: NUM, 6: NUM, 7: "0.0000", 8: NUM0, 9: NUM})
    total_row(ws, end + 1, 4, "Totals (in scope rows)", [5, 9, 11], 10, end)
    ws.cell(end + 2, 4, "Weighted PCAF score").font = F_BOLD
    ws.cell(end + 2, 11, f'=IF(E{end+1}>0,K{end+1}/E{end+1},"")').number_format = "0.0"
    ws.cell(end + 4, 1, "Disclose by scope, industry and asset class; % of gross exposure included; method; excluded items and amounts "
                        "if the Dec 2025 Cat 15 limitation is used. Reconcile outstanding amounts to the FS loan note.").font = F_NOTE


S2_CHECK = [
    ("G1–G7", "Governance: oversight body, skills, information, strategy/transactions, targets & remuneration, management role, controls", "S2.5–7"),
    ("S1", "CRROs described; physical vs transition; horizons; horizon definitions", "S2.10"),
    ("S2", "Industry-based Guidance referred to and considered", "S2.12"),
    ("S3", "Business model and value-chain effects; concentration", "S2.13"),
    ("S4", "Strategy: business model changes, direct/indirect mitigation and adaptation", "S2.14(a)(i)–(iii)"),
    ("S5", "Transition plan incl. key assumptions and dependencies", "S2.14(a)(iv)"),
    ("S6", "How targets will be achieved", "S2.14(a)(v)"),
    ("S7", "Resourcing", "S2.14(b)"),
    ("S8", "Progress on prior plans", "S2.14(c)"),
    ("S9", "Current financial effects; significant risk of material adjustment", "S2.15–16"),
    ("S10", "Anticipated financial effects (position, performance, cash flows)", "S2.15–16"),
    ("S11", "Quantitative omissions justified; qualitative information given", "S2.17–21"),
    ("S12", "Resilience: implications, uncertainties, capacity to adjust", "S2.22(a)"),
    ("S13", "Scenario analysis inputs, assumptions and period", "S2.22(b)"),
    ("R1–R8", "Risk management processes, scenario use, prioritisation, monitoring, changes, opportunities, ERM integration", "S2.24–26"),
    ("M1", "Absolute gross Scope 1, 2, 3 in tCO2e", "S2.29(a)(i)"),
    ("M2", "GHG Protocol 2004 (or jurisdictional relief)", "S2.29(a)(ii)"),
    ("M3", "Measurement approach, inputs, assumptions, reasons, changes", "S2.29(a)(iii)"),
    ("M4", "Scope 1 and 2 disaggregated: consolidated group vs other investees", "S2.29(a)(iv)"),
    ("M5", "Location-based Scope 2 + contractual instruments info", "S2.29(a)(v)"),
    ("M6", "Scope 3 categories included", "S2.29(a)(vi)(1)"),
    ("M7", "Financed emissions (asset management, banking, insurance)", "S2.29(a)(vi)(2); B58–B63"),
    ("M8", "AR6 100-year GWPs (or jurisdictional relief)", "S2.B21–B22"),
    ("M9", "Transition-risk vulnerable amount and %", "S2.29(b)"),
    ("M10", "Physical-risk vulnerable amount and %", "S2.29(c)"),
    ("M11", "Opportunity-aligned amount and %", "S2.29(d)"),
    ("M12", "Capital deployment", "S2.29(e)"),
    ("M13", "Internal carbon price", "S2.29(f)"),
    ("M14", "Climate-linked remuneration", "S2.29(g)"),
    ("I1", "Industry-based metrics", "S2.32"),
    ("T1", "Target elements incl. international agreement link", "S2.33"),
    ("T2", "Validation, review, monitoring metrics, revisions", "S2.34"),
    ("T3", "Performance and trends", "S2.35"),
    ("T4", "GHG targets: gases, scopes, gross/net, sectoral approach", "S2.36(a)–(d)"),
    ("T5", "Planned carbon credits: reliance, scheme, type, integrity", "S2.36(e)"),
    ("TR", "Reliefs and amendments used disclosed", "S2.C3–C5; amendments"),
]


def s2_cl(wb, a):
    ws = header(wb, "S2-CL", "IFRS S2 disclosure checklist", "IFRS S2 (June 2023) + Dec 2025 amendments", a,
                "Status per requirement; use with S1-11 for the general requirements.")
    rows = [[x, y, z, None, None, None] for x, y, z in S2_CHECK]
    end = table(ws, 9, ["#", "Requirement", "Para", "Response / draft wording", "Evidence / WP", "Status"], rows,
                [8, 64, 20, 44, 16, 9], input_cols=[4, 5, 6], example_rows=0, height=30)
    dv_list(ws, f"F10:F{end}", ["Y", "P", "N", "NA", "R"])
    r = end + 2
    for k, s in enumerate(["Y", "P", "N", "NA", "R"]):
        ws.cell(r + k, 2, f"Count {s}").font = F_BOLD
        ws.cell(r + k, 3, f'=COUNTIF(F10:F{end},"{s}")')


def main():
    global EXAMPLES
    p = argparse.ArgumentParser()
    p.add_argument("--entity", default="[Entity name]")
    p.add_argument("--year-end", default="2026-12-31")
    p.add_argument("--currency", default="[presentation currency]")
    p.add_argument("--out", default="S2_Working_Papers.xlsx")
    p.add_argument("--bank", action="store_true", help="add FE-01 financed emissions (PCAF)")
    p.add_argument("--no-examples", action="store_true")
    a = p.parse_args()
    EXAMPLES = not a.no_examples

    wb = Workbook()
    wb.remove(wb.active)
    cr_00(wb, a)
    cr_01(wb, a)
    cr_02(wb, a)
    sa(wb, a)
    cr_03(wb, a)
    cr_04(wb, a)
    cr_05(wb, a)
    # GHG: create lead schedule placeholder position first so it sits before calc sheets
    lead_idx = len(wb.sheetnames)
    ghg_01(wb, a)
    ghg_02(wb, a)
    ghg_03(wb, a)
    ghg_03a(wb, a)
    ghg_03b(wb, a)
    s1_tot = ghg_04(wb, a)
    s2_tot = ghg_05(wb, a)
    s3_first, s3_last = ghg_06(wb, a)
    ghg_07(wb, a)
    ghg_07b(wb, a)
    ghg_08(wb, a)
    ghg_09(wb, a)
    ghg_00(wb, a, s1_tot, s2_tot, s3_first, s3_last)
    wb.move_sheet("GHG-00", offset=lead_idx - wb.sheetnames.index("GHG-00"))
    cr_06(wb, a, 10, s2_tot - 1)
    cr_07(wb, a)
    for k, name in enumerate(("CR-06", "CR-07")):
        wb.move_sheet(name, offset=(wb.sheetnames.index("CR-05") + 1 + k) - wb.sheetnames.index(name))
    if a.bank:
        fe_01(wb, a)
    s2_cl(wb, a)
    if not a.bank:
        idx = wb["CR-00"]
        for row in idx.iter_rows():
            for c in row:
                if c.value == "FE-01":
                    c.value, c.hyperlink = "FE-01 (run with --bank)", None
    for ws in wb.worksheets:
        ws.freeze_panes = "A10"
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    wb.save(a.out)
    print(f"Saved {a.out} with sheets: {', '.join(wb.sheetnames)}")


if __name__ == "__main__":
    main()
