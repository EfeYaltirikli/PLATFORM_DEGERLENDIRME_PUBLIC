# -*- coding: utf-8 -*-
"""python3 tests/test_core.py [out.xlsx] - builds a report like Hexagon's and checks the rules."""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import openpyxl
import plat_core as core

ok = bad = 0
def check(name, got, want):
    global ok, bad
    if got == want: ok += 1
    else:
        bad += 1; print("HATA", name, "->", got, "beklenen", want)

def make(path, cols):
    wb = openpyxl.Workbook()
    s = wb.active; s.title = "Summary"
    s["B2"] = "Part No"; s["B3"] = "1234M56G01_Stg1"; s.merge_cells("B3:C3")
    s["D2"] = "Serial No"; s["D3"] = "99X_1234M56G02_OP100"
    s["H2"] = "Operation"; s["H4"] = "AIRFOIL_INSPECTION"
    s["B5"] = "Operator"; s["B6"] = "ABC"; s["C5"] = "CMM"; s["C6"] = 12345
    s["D5"] = "NC W/S"; s["D6"] = "ZIGZAG"; s["E5"] = "Spindle"; s["E6"] = 550
    s["H5"] = "Date"; s["H6"] = "MON JAN-05-2026 10:00"
    s["B7"] = "Company"; s["D7"] = "Location"; s["D8"] = "XYZ"
    p = wb.create_sheet("Out of Tolerance"); p["A1"] = "x"
    p = wb.create_sheet("Platform")
    p["A1"] = "Summary Page"; p["B2"] = "Platform"
    for j, (name, vals) in enumerate(cols, 2):
        p.cell(row=3, column=j, value=name); p.cell(row=4, column=j, value="CNX403")
        for i, v in enumerate(vals):
            p.cell(row=5 + i, column=j, value=v)
    n = max(len(v) for _n, v in cols)
    for i in range(n): p.cell(row=5 + i, column=1, value=i + 1)
    # tolerance lines under the table - must be ignored
    p.cell(row=5 + n + 1, column=1, value="Upper Tol")
    p.cell(row=5 + n + 2, column=1, value="Lower Tol")
    for j in range(2, 2 + len(cols)):
        p.cell(row=5 + n + 1, column=j, value=0.05); p.cell(row=5 + n + 2, column=j, value=-0.05)
    wb.create_sheet("Tip Profile")
    wb.save(path)

N = 40
cols = [
    ("SPF_1", [0.002] * N),                                   # 1 pass
    ("PLATFORM_2", [0.001] * 38 + [0.006, -0.0099]),          # 2 c010=2 (5%=2) -> 2<2 no -> FAIL
    ("PLATFORM_3", [0.001] * 39 + [0.007]),                   # 3 c010=1 <2 -> PASS
    ("PLATFORM_4", [0.001] * 39 + [-0.0101]),                 # 4 > .010 -> FAIL
    ("PLATFORM_5", [0.005] * 20 + [-0.005] * 20),             # 5 exact .005 inside, avg 0 -> PASS
    ("PLATFORM_6", [0.004] * N),                              # 6 avg exactly .004 -> PASS
    ("PLATFORM_7", [0.0041] * N),                             # 7 avg .0041 -> FAIL
    ("PLATFORM_8", ["0.0010"] * 30 + [None] * 10),            # 8 text values, 30 points -> PASS
    ("PLATFORM_9", [0.001] * 39 + [0.010]),                   # 9 exactly .010 -> c010=1 -> PASS
]
d = tempfile.mkdtemp()
src = os.path.join(d, "rapor.xlsx")
make(src, cols)
rep = core.load(src)
P = rep["platforms"]
check("platform sayısı", len(P), 9)
check("nokta", rep["points"], N)
check("pass", [p["pass"] for p in P], [True, False, True, False, True, True, False, True, True])
check("c010", [p["c010"] for p in P], [0, 2, 1, 0, 0, 0, 0, 0, 1])
check("c020", [p["c020"] for p in P], [0, 0, 0, 1, 0, 0, 0, 0, 0])
check("n8", P[7]["n"], 30)
check("isim", P[1]["name"], "PLATFORM_2")
check("genel", rep["pass"], False)
i = rep["info"]
check("part", i["part"], "1234M56G01_Stg1")
check("serial", i["serial"], "99X_1234M56G02_OP100")
check("operation", i["operation"], "AIRFOIL_INSPECTION")
check("cmm", i["cmm"], "12345")
check("spindle", i["spindle"], "550")
check("location", i["location"], "XYZ")
check("company", i["company"], "")
check("operator", i["operator"], "ABC")
out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(d, core.default_name(rep))
core.export_xlsx(rep, out)
w = openpyxl.load_workbook(out)
ws = w["PLATFORM"]
check("A2", ws["A2"].value, "PLATFORM #1")
check("D3", ws["D3"].value, "FAIL")
check("D2 renk", ws["D2"].fill.fgColor.rgb[-6:], "00B050")
check("D3 renk", ws["D3"].fill.fgColor.rgb[-6:], "FF0000")
check("sayfalar", w.sheetnames, ["PLATFORM", "PARÇA BİLGİSİ", "NOKTALAR"])
check("dosya adı", core.default_name(rep), "1234M56G01_Stg1_99X_1234M56G02_OP100_PLATFORM.xlsx")
# only pass so far: all pass
make(src, cols[:1]); check("tek pass", core.load(src)["pass"], True)
print("tamam %d / hata %d" % (ok, bad))
print("ORNEK", src, out)
