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

VERSION = "2.0.0"
HEAD_ROW = 10            # Excel row of the table header (same as in the macro)


def build(out, macro=True):
    with open(os.path.join(HERE, "PlatformAnaliz.bas"), encoding="ascii") as f:
        src = f.read()
    if out.lower().endswith(".xlsx"):
        macro = False
    binp = os.path.join(tempfile.mkdtemp(), "vbaProject.bin")
    vba_bin.build([("ThisWorkbook", "doc_wb", ""), ("shAnaliz", "doc_ws", ""),
                   ("shNokta", "doc_ws", ""), ("PlatformAnaliz", "std", src)], binp)

    wb = xlsxwriter.Workbook(out)
    ws = wb.add_worksheet("ANALİZ")
    wn = wb.add_worksheet("NOKTALAR")
    if macro:
        wb.set_vba_name("ThisWorkbook")
        wb.add_vba_project(binp)
        ws.set_vba_name("shAnaliz")
        wn.set_vba_name("shNokta")

    F = lambda **k: wb.add_format(dict({"font_name": "Calibri", "font_size": 11}, **k))
    title = F(bold=True, font_size=16)
    sub = F(font_color="#595959", font_size=9)
    lab = F(bold=True, bg_color="#D9E1F2", border=1, valign="vcenter")
    val = F(border=1, valign="vcenter", bold=True, num_format="@")
    head = F(bold=True, bg_color="#D9E1F2", border=1, align="center", valign="vcenter",
             text_wrap=True)
    big = F(bold=True, font_size=28, align="center", valign="vcenter", border=2,
            font_color="#FFFFFF", bg_color="#A6A6A6")
    small = F(align="center", font_color="#595959")

    for c, w in enumerate((18, 20, 20, 14, 13, 10, 11, 16)):
        ws.set_column(c, c, w)
    ws.write("A1", "PLATFORM DEĞERLENDİRME", title)
    ws.write("A2", "Rapor yüklenmedi.  EXCEL YÜKLE düğmesine basın.", sub)
    ws.write("E1", "v%s  ·  created by EYALTIRIKLI" % VERSION, sub)
    ws.set_row(0, 30)

    labels = (("Part No", "Serial No"), ("Operation", "Operator"), ("CMM", "NC W/S"),
              ("Spindle", "Date"), ("Company", "Location"))
    for i, (l1, l2) in enumerate(labels):
        r = 3 + i
        ws.write(r, 0, l1, lab)
        ws.merge_range(r, 1, r, 3, "", val)
        ws.write(r, 4, l2, lab)
        ws.merge_range(r, 5, r, 6, "", val)
        ws.set_row(r, 20)
    ws.write("H3", "GENEL SONUÇ", F(bold=True, align="center", font_color="#595959"))
    ws.merge_range("H4:H8", "—", big)
    ws.write("H9", "", small)

    heads = ("PLATFORM NO.", "OUT PROFILE .010\nPOINT COUNT", "OUT PROFILE .020\nPOINT COUNT",
             "AVERAGE\n(±.004)", "TOTAL POINT\nCOUNT", "%5 LIMIT", "MAX |DEV|", "PASS/FAIL")
    for c, h in enumerate(heads):
        ws.write(HEAD_ROW - 1, c, h, head)
    ws.set_row(HEAD_ROW - 1, 36)
    ws.freeze_panes(HEAD_ROW, 0)

    ws.insert_button("B1", {"macro": "RaporYukle", "caption": "EXCEL YÜKLE",
                            "width": 140, "height": 30, "x_offset": 4, "y_offset": 4})
    ws.insert_button("C1", {"macro": "SonucuKaydet", "caption": "SONUCU KAYDET (.xlsx)",
                            "width": 160, "height": 30, "x_offset": 14, "y_offset": 4})
    ws.activate()
    for sh in (ws, wn):
        sh.set_landscape()
        sh.set_paper(9)
        sh.fit_to_pages(1, 0)

    wn.write("A1", "NOKTA", head)
    wn.freeze_panes(1, 1)
    wn.set_column(0, 0, 8)
    wb.close()


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(HERE), "PLATFORM_DEGERLENDIRME.xlsm")
    build(out)
    print(out)
