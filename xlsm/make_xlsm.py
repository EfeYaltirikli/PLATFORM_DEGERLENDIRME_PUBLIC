# -*- coding: utf-8 -*-
"""python3 make_xlsm.py [out.xlsm] - builds PLATFORM_DEGERLENDIRME.xlsm:
the ANALIZ sheet (buttons, part info, the result table) and NOKTALAR,
with the macro of PlatformAnaliz.bas inside."""

import os
import sys
import tempfile

import xlsxwriter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import vba_bin  # noqa: E402

VERSION = "2.3.0"
HEAD_ROW = 6             # Excel row of the table header (same as in the macro)


def build(out, macro=True):
    with open(os.path.join(HERE, "PlatformAnaliz.bas"), encoding="ascii") as f:
        src = f.read()
    if out.lower().endswith(".xlsx"):
        macro = False
    binp = os.path.join(tempfile.mkdtemp(), "vbaProject.bin")
    vba_bin.build([("ThisWorkbook", "doc_wb", ""), ("shAnaliz", "doc_ws", ""),
                   ("shNokta", "doc_ws", ""), ("PlatformAnaliz", "std", src)], binp)

    wb = xlsxwriter.Workbook(out)
    ws = wb.add_worksheet("ANALYSIS")
    wn = wb.add_worksheet("POINTS")
    if macro:
        wb.set_vba_name("ThisWorkbook")
        wb.add_vba_project(binp)
        ws.set_vba_name("shAnaliz")
        wn.set_vba_name("shNokta")

    F = lambda **k: wb.add_format(dict({"font_name": "Calibri", "font_size": 11}, **k))
    sub = F(font_color="#595959", font_size=9)
    lab = F(bold=True, bg_color="#D9E1F2", border=1, valign="vcenter")
    val = F(border=1, valign="vcenter", bold=True, num_format="@")
    head = F(bold=True, bg_color="#D9E1F2", border=1, align="center", valign="vcenter",
             text_wrap=True)
    big = F(bold=True, font_size=18, align="center", valign="vcenter", border=2,
            font_color="#FFFFFF", bg_color="#A6A6A6")
    small = F(align="center", font_color="#595959", font_size=9)

    for c, w in enumerate((18, 20, 20, 14, 11, 12, 13, 16)):
        ws.set_column(c, c, w)
    ws.set_row(0, 26)
    ws.write("A2", "No report loaded.", sub)

    # 4 fields: Part No / Serial No, CMM / Date
    for r, (l1, l2) in ((2, ("Part No", "Serial No")), (3, ("CMM", "Date"))):
        ws.write(r, 0, l1, lab)
        ws.merge_range(r, 1, r, 3, "", val)
        ws.write(r, 4, l2, lab)
        ws.merge_range(r, 5, r, 6, "", val)
        ws.set_row(r, 20)
    ws.write("H2", "OVERALL RESULT", small)
    ws.merge_range("H3:H4", "-", big)
    ws.write("H5", "", small)
    ws.set_row(4, 14)

    heads = ("PLATFORM NO.", "OUT PROFILE .010\nPOINT COUNT", "OUT PROFILE .020\nPOINT COUNT",
             "AVERAGE\n(±.004)", "TOTAL POINT\nCOUNT", "5% LIMIT", "MAX |DEV|", "PASS/FAIL")
    for c, h in enumerate(heads):
        ws.write(HEAD_ROW - 1, c, h, head)
    ws.set_row(HEAD_ROW - 1, 32)
    ws.freeze_panes(HEAD_ROW, 0)

    ws.insert_button("A1", {"macro": "RaporYukle", "caption": "LOAD REPORT",
                            "width": 120, "height": 26, "x_offset": 2, "y_offset": 4})
    ws.insert_button("B1", {"macro": "SonucuKaydet", "caption": "SAVE RESULT (.xlsx)",
                            "width": 150, "height": 26, "x_offset": 2, "y_offset": 4})
    ws.insert_button("D1", {"macro": "Temizle", "caption": "CLEAR",
                            "width": 90, "height": 26, "x_offset": 2, "y_offset": 4})
    ws.activate()
    for sh in (ws, wn):
        sh.set_landscape()
        sh.set_paper(9)
        sh.fit_to_pages(1, 0)

    wn.write("A1", "PLATFORM", head)
    wn.freeze_panes(1, 1)
    wn.set_column(0, 0, 11)
    wb.close()


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(HERE), "PLATFORM_DEGERLENDIRME.xlsm")
    build(out)
    print(out)
