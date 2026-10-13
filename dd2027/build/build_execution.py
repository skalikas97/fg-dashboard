"""AL-IDRISI DD2027 - execution / preparation workbook (AA style), built from spec.

Spec: pack/03, pack/03b and build/spec/AA_* (extract_spec.py), TS decisions in STATE.md.
Tabs: Daily SOW (master), Spares, Subcontractors-attendees, Planning.
Usage: python3 -I build/build_execution.py out/AI_DD2027_Execution_Rev01.xlsx
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

import frame
from common import (AUTHOR, CATEGORIES, DATE_FMT, REV0, REV_ROWS, finish)

L = frame.L
VESSEL_TITLE = "AL-IDRISI - Drydock 2027 <Place>"
MASTER = "Daily SOW"
BUFFER_ROWS = 5          # T4
BUFFER_SPARES = 15       # T4

GREY_HEADER = "A6A6A6"   # th0-0.35 (AA)
GREY_CAT = "BFBFBF"      # th0-0.25 (AA)
TITLE_RED = "A80000"
TITLE_FONT = "F2F2F2"    # th0-0.05 (AA)


def side(style):
    return Side(style=style) if style else Side()


def border(l=None, r=None, t=None, b=None):
    return Border(left=side(l), right=side(r), top=side(t), bottom=side(b))


def fill(rgb):
    return PatternFill("solid", fgColor="FF" + rgb) if rgb else PatternFill()


def arial(size, bold=False, color=None):
    return Font(name="Arial", size=size, bold=bold, color=("FF" + color) if color else None)


def style(cell, font=None, fil=None, brd=None, al=None, nf=None):
    if font: cell.font = font
    if fil: cell.fill = fil
    if brd: cell.border = brd
    if al: cell.alignment = al
    if nf: cell.number_format = nf


CENTRE = Alignment(horizontal="center", vertical="center")


def span_cells(ws, r, c1, c2, value=None, font=None, brd_first=None, brd_mid=None, brd_last=None, nf=None, al=CENTRE):
    """Write a value over columns c1..c2 of row r (merged when c2 > c1), with edge borders."""
    if c2 > c1:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    for c in range(c1, c2 + 1):
        b = brd_first if c == c1 else brd_last if c == c2 else brd_mid
        style(ws.cell(r, c), font, brd=b, nf=nf, al=al)
    if value is not None:
        ws.cell(r, c1, value)


def top_block(ws, subtitle, title_c2, rev_spans, master, legend=(), master_name=None, rev_src=6):
    """AA top block, rows 1-13: title B2:<title_c2>13, Rev block (10 rows) + Last change row.

    rev_spans: three (c1, c2) column spans for "Rev. #", "Name", "Date".
    legend: list of (c1, c2, row, text, rgb, font_size, bordered).
    Returns the Last change row number.
    """
    last = 2 + REV_ROWS + 1
    ws.row_dimensions[1].height = 10.5
    ws.row_dimensions[2].height = 15.0
    for r in range(3, last):
        ws.row_dimensions[r].height = 14.25
    ws.row_dimensions[last].height = 15.0

    ws.merge_cells(start_row=2, start_column=2, end_row=last, end_column=title_c2)
    t = ws.cell(2, 2, f"{VESSEL_TITLE}\n{subtitle}")
    if master:
        style(t, arial(28, color=TITLE_FONT), fill(TITLE_RED))
    else:
        style(t, arial(28), fill(GREY_HEADER))
    t.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(2, last + 1):
        for c in range(2, title_c2 + 1):
            ws.cell(r, c).border = border(l="medium" if c == 2 else None,
                                          t="medium" if r == 2 else None,
                                          b="medium" if r == last else None)

    (a1, a2), (b1, b2), (d1, d2) = rev_spans
    heads = (("Rev. #", a1, a2, "medium", "thin"), ("Name", b1, b2, "thin", "thin"), ("Date", d1, d2, "thin", "medium"))
    for txt, c1, c2, lft, rgt in heads:
        span_cells(ws, 2, c1, c2, txt, arial(10),
                   border(l=lft, t="medium", r=rgt if c1 == c2 else None),
                   border(t="medium"), border(r=rgt, t="medium"))
    for i in range(REV_ROWS):
        r = 3 + i
        for k, (txt, c1, c2, lft, rgt) in enumerate(heads):
            if master:
                val = REV0[k] if i == 0 else None
            else:
                ref = f"'{master_name or MASTER}'!{L(rev_src + k)}{r}"
                val = f'=IF({ref}="","",{ref})'
            span_cells(ws, r, c1, c2, val, arial(10),
                       border(l=lft, t="thin", b="thin", r=rgt if c1 == c2 else None),
                       border(t="thin", b="thin"), border(r=rgt, t="thin", b="thin"),
                       nf=DATE_FMT if k == 2 else None)

    # Last change row: label over the Rev. # + Name spans, date under Date
    span_cells(ws, last, a1, b2, "Last change", arial(10),
               border(l="medium", t="medium", b="medium"), border(t="medium", b="medium"),
               border(r="thin", t="medium", b="medium"))
    span_cells(ws, last, d1, d2, REV0[2] if master else "=LastChange", arial(10),
               border(l="thin", t="medium", b="medium", r="medium" if d1 == d2 else None),
               border(t="medium", b="medium"), border(r="medium", t="medium", b="medium"), nf=DATE_FMT)

    for c1, c2, r, txt, rgb, size, bordered in legend:
        b = border("thin", "thin", "thin", "thin") if bordered else None
        if c2 > c1:
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        for c in range(c1, c2 + 1):
            style(ws.cell(r, c), arial(size), fill(rgb), b,
                  CENTRE if bordered else Alignment(vertical="center"))
        ws.cell(r, c1, txt)
    return last


def page_aa(ws, orientation, paper, fit_w, fit_h, scale, margins, zoom=None):
    ps = ws.page_setup
    ps.orientation, ps.paperSize = orientation, paper
    if fit_w is not None or fit_h is not None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        if fit_w is not None: ps.fitToWidth = fit_w
        if fit_h is not None: ps.fitToHeight = fit_h
    if scale: ps.scale = scale
    m = ws.page_margins
    m.left, m.right, m.top, m.bottom, m.header, m.footer = margins
    if zoom: ws.sheet_view.zoomScale = zoom


# ---------------------------------------------------------------- Daily SOW
def build_daily_sow(wb):
    ws = wb.active
    ws.title = MASTER
    widths = {"A": 2.0, "B": 14.28515625, "C": 70.85546875, "D": 6.0, "E": 17.5703125,
              "F": 12.140625, "G": 12.42578125, "H": 16.42578125, "I": 62.42578125, "J": 28.0}
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    legend = [(10, 10, 2 + i, t, rgb, 11, False) for i, (t, rgb) in enumerate(
        [("Planned", None), ("Ongoing", "FFFF00"), ("Completed", "92D050"), ("TBC", "FFC000"),
         ("Cancelled/postponed", "FF0000"), ("All necessary parts on board", "00B0F0")])]
    last = top_block(ws, "Scope of works", 5, [(6, 6), (7, 7), (8, 8)], True, legend)
    wb.defined_names["LastChange"] = DefinedName("LastChange", attr_text=f"'{MASTER}'!$H${last}")

    hdr = last + 2
    ws.row_dimensions[hdr - 1].height = 15.0
    ws.row_dimensions[hdr].height = 18.75
    heads = {2: ("Specs", "left", True, ("medium", "medium")),
             3: ("Description", "center", False, (None, "medium")),
             4: (None, "center", False, (None, "medium")),
             5: ("WL item", "center", True, ("medium", "medium")),
             6: ("3rd party", "center", True, ("medium", "medium")),
             7: ("Crew", "center", True, ("medium", "medium")),
             8: ("JDN WS", "center", True, ("medium", "medium")),
             9: ("Remarks", "left", True, ("medium", "medium"))}
    for c, (txt, h, wrap, (lb, rb)) in heads.items():
        style(ws.cell(hdr, c, txt), arial(14), fill(GREY_HEADER),
              border(l=lb, r=rb, t="medium", b="medium"),
              Alignment(horizontal=h, vertical="center", wrap_text=wrap or None))

    cat_al = {c: Alignment(horizontal="left") for c in (2, 3, 4)}
    cat_al.update({c: CENTRE for c in (5, 6, 7, 8)})
    cat_al[9] = Alignment(horizontal="left", wrap_text=True)
    data = {2: (border(l="medium", r="thin", t="dashed", b="dashed"), Alignment(horizontal="left", vertical="center")),
            3: (border(r="medium", t="dashed", b="dashed"), Alignment(vertical="center")),
            4: (border(r="medium", t="dashed", b="dashed"), Alignment(vertical="center")),
            9: (border(l="medium", r="medium", t="dashed", b="dashed"), Alignment(horizontal="left", vertical="center", wrap_text=True))}
    for c in (5, 6, 7, 8):
        data[c] = (border(l="medium", r="medium", t="dashed", b="dashed"), CENTRE)

    def cat_row(r, label):
        ws.row_dimensions[r].height = 15.0
        for c in range(2, 10):
            style(ws.cell(r, c), arial(10, bold=True), fill(GREY_CAT),
                  border(l="medium" if c == 2 else None, t="medium", b="medium"), cat_al[c])
        if label:
            ws.cell(r, 2, label)

    r = hdr + 1
    for label in CATEGORIES:
        cat_row(r, label)
        r += 1
        for _ in range(BUFFER_ROWS):
            for c, (brd, al) in data.items():
                style(ws.cell(r, c), arial(10), brd=brd, al=al)
            r += 1
    cat_row(r, None)
    end = r

    ws.freeze_panes = ws.cell(hdr + 1, 1)
    ws.auto_filter.ref = f"A{hdr}:J{end}"
    ws.print_area = f"A1:I{end}"
    # P1 default: fit to one page wide (Remarks on page 1) instead of AA fit-to-height
    page_aa(ws, "portrait", 8, 1, 0, 80,
            (0.23622047244094488,) * 2 + (0.7480314960629921,) * 2 + (0.31496062992125984,) * 2, 85)
    return ws


# ---------------------------------------------------------------- generic AA list tab
def aa_list_tab(wb, ref_name, title, subtitle, title_c2, rev_spans, legend, last_col,
                header_ref_row, cat_ref_row, data_ref_row, buffer_rows, cat_merge, header_text=None):
    spec = frame.load(ref_name)
    ws = wb.create_sheet(title)
    frame.widths(spec, ws, 1, last_col)
    last = top_block(ws, subtitle, title_c2, rev_spans, False, legend)
    hdr = last + 2
    ws.row_dimensions[hdr - 1].height = 15.0
    frame.copy_row(spec, header_ref_row, ws, hdr, 2, last_col, values=True)
    for c, txt in (header_text or {}).items():
        ws.cell(hdr, c, txt)
    r = hdr + 1
    for label in CATEGORIES:
        frame.copy_row(spec, cat_ref_row, ws, r, 2, last_col)
        ws.cell(r, 2, label)
        if cat_merge:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        r += 1
        for _ in range(buffer_rows):
            frame.copy_row(spec, data_ref_row, ws, r, 2, last_col, fill=False)
            for c in range(2, last_col + 1):          # status colour is set by the TS, not by the template
                ws.cell(r, c).font = Font(name="Arial", size=10)
            r += 1
    ws.freeze_panes = ws.cell(hdr + 1, 1)
    return ws, spec, hdr, r - 1


def build_spares(wb):
    legend = [(12, 12, 2 + i, t, rgb, 11, False) for i, (t, rgb) in enumerate(
        [("too late", "FF0000"), ("close call", "FFC000"), ("delivered in <Place>", "FFFF00"),
         ("on board", "92D050"), ("on the way", None)])]
    ws, spec, hdr, end = aa_list_tab(
        wb, "AA_DAILY_SOW_Afloat_Repairs_final__Spares", "Spares", "Overview of spares",
        7, [(8, 8), (9, 9), (10, 10)], legend, 19, 10, 11, 12, BUFFER_SPARES, True,
        header_text={15: "ETA <Place>"})
    ws.auto_filter.ref = f"B{hdr}:S{end}"
    page_aa(ws, "portrait", 9, None, None, None, tuple(spec["margins"]), 81)


def build_subcontractors(wb):
    ws, spec, hdr, end = aa_list_tab(
        wb, "AA_DAILY_SOW_Afloat_Repairs_final__Subcontractors-attendees", "Subcontractors-attendees",
        "Overview of subcontractors", 4, [(5, 5), (6, 6), (7, 7)], [], 12, 10, 11, 12, BUFFER_ROWS, True)
    ws.auto_filter.ref = f"B{hdr}:P{hdr}"
    for word, colour in (("Canceled", "FF0000"), ("Requested", "FFFF00"), ("Granted", "92D050")):
        frame.text_rule(ws, f"B{hdr + 1}:L{end}", word, colour)
    page_aa(ws, "landscape", 8, None, 0, 73, tuple(spec["margins"]))


def build_planning(wb, days=35):
    spec = frame.load("AA_DAILY_SOW_Afloat_Repairs_final__Planning")
    ws = wb.create_sheet("Planning")
    first_day, remarks = 4, 4 + days               # D .. AL, AM
    frame.widths(spec, ws, 1, 3)
    for c in range(first_day, remarks):
        ws.column_dimensions[L(c)].width = 2.7109375
    ws.column_dimensions[L(remarks)].width = 51.5703125
    team = [("Engine room", "C6E0B4"), ("Electricians", "B4C6E7"), ("Deck", "FFFF00"),
            ("Workshop", "F8CBAD"), ("3rd party", "FF8181")]
    crew = [("ER crew change", "C5E0B4", 10), ("ELEC crew change", "B4C7E7", 10),
            ("DECK crew change", "FFFF00", 10), ("LOCAL crew change", "F8CBAD", 8)]
    legend = [(20, 24, 2 + i, t, rgb, 10, True) for i, (t, rgb) in enumerate(team)]
    legend += [(27, 34, 2 + i, t, rgb, s, True) for i, (t, rgb, s) in enumerate(crew)]
    last = top_block(ws, "General planning", 3, [(4, 8), (9, 13), (14, 18)], False, legend)

    # white spacer row, crew change rows, header (ref rows 9, 10-11, 12)
    white, cc1, cc2, hdr = last + 1, last + 2, last + 3, last + 4
    frame.copy_row(spec, 9, ws, white, 2, remarks - 1)
    frame.copy_row(spec, 10, ws, cc1, 2, remarks)
    frame.copy_row(spec, 11, ws, cc2, 2, remarks)
    for r_new, r_ref in ((cc1, 10), (cc2, 11)):
        for c in range(first_day, remarks):
            frame.apply(ws.cell(r_new, c), frame.cell(spec, f"{L(min(c, 38))}{r_ref}"))
    ws.merge_cells(start_row=cc1, start_column=3, end_row=cc2, end_column=3)
    ws.cell(cc1, 3, "Crew change")

    frame.copy_row(spec, 12, ws, hdr, 2, 3, values=True)
    frame.apply(ws.cell(hdr, remarks), frame.cell(spec, "AM12"), value=True)
    for i, c in enumerate(range(first_day, remarks)):
        frame.apply(ws.cell(hdr, c), frame.cell(spec, f"{L(4 + i % 7)}12"))   # Mon-Fri E7E6E6, Sat-Sun AFABAB
        ws.cell(hdr, c).number_format = "d-mmm"
        ws.cell(hdr, c).value = None if i == 0 else f'=IF({L(c - 1)}{hdr}="","",{L(c - 1)}{hdr}+1)'
    # T6: one start-date cell drives the grid; a Monday keeps weekend shading and Sunday lines right
    dv = DataValidation(type="custom", formula1=f"WEEKDAY({L(first_day)}{hdr},2)=1", allow_blank=True,
                        showInputMessage=True, showErrorMessage=True,
                        promptTitle="Start date", prompt="Enter the first day of the planning (a Monday).",
                        errorTitle="Start date", error="The first day must be a Monday.")
    dv.add(f"{L(first_day)}{hdr}")
    ws.add_data_validation(dv)

    r = hdr + 1
    for label in CATEGORIES:
        frame.copy_row(spec, 13, ws, r, 2, 3)
        for c in range(first_day, remarks + 1):
            frame.apply(ws.cell(r, c), frame.cell(spec, "D13" if c < remarks else "AM13"))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=remarks)
        ws.cell(r, 2, label)
        r += 1
        for _ in range(BUFFER_ROWS):
            frame.apply(ws.cell(r, 2), frame.cell(spec, "B15"), fill=False)
            frame.apply(ws.cell(r, 3), frame.cell(spec, "C15"), fill=False)
            for i, c in enumerate(range(first_day, remarks)):
                d = dict(frame.cell(spec, f"{L(4 + i % 7)}16"))
                d["font"] = ["Arial", 10.0, False, False, None, None]
                frame.apply(ws.cell(r, c), d, fill=False)
            frame.apply(ws.cell(r, remarks), frame.cell(spec, "AM15"), fill=False)
            r += 1
    end = r - 1

    body = f"{L(first_day)}{hdr + 1}:{L(remarks - 1)}{end}"
    # day codes coloured with font = fill (letter invisible), colours as the legend
    for code, colour in (('"E"', "C6E0B4"), ('"L"', "B4C6E7"), ('"D"', "FFFF00"), ('"W"', "F8CBAD"), ("3", "FF8181")):
        frame.equal_rule(ws, body, code, colour)
    ws.conditional_formatting.add(f"B{hdr + 1}:C{end}", FormulaRule(
        formula=[f"COUNTA(${L(first_day)}{hdr + 1}:${L(remarks - 1)}{hdr + 1})>0"], fill=fill("FFFFCC")))

    ws.freeze_panes = ws.cell(hdr + 1, 1)
    page_aa(ws, "landscape", 8, None, 0, 87, tuple(spec["margins"]))


def main(out):
    wb = Workbook()
    build_daily_sow(wb)
    build_spares(wb)
    build_subcontractors(wb)
    build_planning(wb)
    wb.active = 0
    finish(wb, out)


if __name__ == "__main__":
    main(sys.argv[1])
