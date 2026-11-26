# -*- coding: utf-8 -*-
"""PLATFORM DEGERLENDIRME - reading the Hexagon report (.xlsx), judging every
platform and writing the result to Excel.  No Qt in here.

Rule for one platform (one column of the Platform sheet), with N = the
number of points in it:
    |dev| <= 0.005            -> inside the .010 profile, fine
    0.005 < |dev| <= 0.010    -> "OUT PROFILE .010" point
    |dev| > 0.010             -> "OUT PROFILE .020" point
    PASS  when  .020 points == 0  and  .010 points < 5 % of N
          and  |average of the deviations| <= 0.004  (the .008 profile)
"""

import os
import re

VERSION = "1.0.1"
APP_TITLE = "PLATFORM DEĞERLENDİRME"
APP_CREDIT = "created by EYALTIRIKLI"

TOL_IN = 0.005           # +/- .005  = the .010 profile
TOL_OUT = 0.010          # +/- .010  = the .020 profile
PCT = 0.05               # allowed share of points between the two
TOL_AVG = 0.004          # +/- .004 = the .008 profile, for the average
EPS = 1e-9               # exactly .005 / .010 counts as inside

# the header block of the Summary sheet: label in the report -> key
INFO_FIELDS = (
    ("Part No", "part"), ("Serial No", "serial"), ("Operation", "operation"),
    ("Operator", "operator"), ("CMM", "cmm"), ("NC W/S", "ncws"),
    ("Spindle", "spindle"), ("Date", "date"), ("Company", "company"),
    ("Location", "location"),
)
INFO_LABELS = dict((k, l) for l, k in INFO_FIELDS)


class ReadError(Exception):
    pass


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def _num(v):
    """A cell as a number, or None.  Text like '0.0026' or '-0,0002' too."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if hasattr(v, "strftime"):
        return v.strftime("%d.%m.%Y %H:%M")
    return str(v).strip()


def _rows(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _cell(rows, r, c):
    if 0 <= r < len(rows) and 0 <= c < len(rows[r]):
        return rows[r][c]
    return None


def _find_sheet(wb, word):
    names = wb.sheetnames
    for n in names:
        if _norm(n) == word:
            return wb[n]
    for n in names:
        if word in _norm(n):
            return wb[n]
    return None


# ---------------------------------------------------------------------------
#  Summary sheet: part information
# ---------------------------------------------------------------------------

def read_info(rows):
    """Every label of INFO_FIELDS is looked up; its value is the first
    filled cell under it (a few rows down), else the one to its right.
    Part No falls back on B3."""
    labels = set(_norm(l) for l, _k in INFO_FIELDS)
    found = {}
    for r, row in enumerate(rows[:60]):
        for c, v in enumerate(row[:40]):
            if v is None:
                continue
            n = _norm(v)
            for lab, key in INFO_FIELDS:
                if key in found or n != _norm(lab):
                    continue
                val = ""
                for dr in range(1, 4):
                    x = _cell(rows, r + dr, c)
                    if x is not None and _text(x):
                        if _norm(x) not in labels:
                            val = _text(x)
                        break
                if not val:
                    for dc in range(1, 4):
                        x = _cell(rows, r, c + dc)
                        if x is not None and _text(x):
                            if _norm(x) not in labels:
                                val = _text(x)
                            break
                found[key] = val
    info = dict((k, found.get(k, "")) for _l, k in INFO_FIELDS)
    if not info["part"]:
        info["part"] = _text(_cell(rows, 2, 1))          # B3
    return info


# ---------------------------------------------------------------------------
#  Platform sheet: the deviations
# ---------------------------------------------------------------------------

def _is_int(v, want=None):
    x = _num(v)
    if x is None or abs(x - round(x)) > 1e-9:
        return False
    return want is None or int(round(x)) == want


def read_platforms(rows):
    """Points run top to bottom (point no. in column A: 1, 2, 3 ...),
    platforms left to right from column B.  Only those numbered rows are
    data - the tolerance lines under the table are left out.
    Returns (platforms, point_count) with every platform
    {"no", "name", "col", "values": [float or None per point row]}."""
    start = None
    for r, row in enumerate(rows):
        if _is_int(_cell(rows, r, 0), 1) and any(_num(x) is not None for x in row[1:]):
            start = r
            break
    if start is None:
        raise ReadError("Platform sayfasında veri bulunamadı "
                        "(A sütununda 1, 2, 3 ... nokta sırası bekleniyor).")
    end = start
    while _is_int(_cell(rows, end + 1, 0), end - start + 2):
        end += 1
    data = rows[start:end + 1]
    width = max(len(r) for r in data)
    cols = [c for c in range(1, width)
            if any(_num(_cell(data, i, c)) is not None for i in range(len(data)))]
    if not cols:
        raise ReadError("Platform sayfasında sayısal sapma değeri yok.")
    plats = []
    for i, c in enumerate(cols):
        name = ""
        for hr in range(start - 1, max(-1, start - 4), -1):
            t = _text(_cell(rows, hr, c))
            # the row right above is the CMM name (CNX403) - the feature
            # name (PLATFORM_..) is the one above that when it is there
            if t and not re.fullmatch(r"CNX?\d+", t, re.I):
                name = t
                break
        vals = [_num(_cell(data, k, c)) for k in range(len(data))]
        plats.append({"no": i + 1, "name": name, "col": c, "values": vals})
    return plats, len(data)


def judge(values):
    """One platform's numbers and verdict."""
    pts = [v for v in values if v is not None]
    n = len(pts)
    c010 = sum(1 for v in pts if TOL_IN + EPS < abs(v) <= TOL_OUT + EPS)
    c020 = sum(1 for v in pts if abs(v) > TOL_OUT + EPS)
    # 1.0.1: the 5 % is a whole number of points, rounded half up
    # (40 -> 2, 30 -> 1.5 -> 2, 50 -> 2.5 -> 3, 45 -> 2.25 -> 2)
    limit = int(n * PCT + 0.5)
    avg = (sum(pts) / n) if n else 0.0
    avg_ok = abs(avg) <= TOL_AVG + EPS
    ok = n > 0 and c020 == 0 and c010 < limit and avg_ok
    return {"n": n, "c010": c010, "c020": c020, "limit": limit,
            "avg": avg, "avg_ok": avg_ok,
            "max": max((abs(v) for v in pts), default=0.0),
            "pct": (100.0 * c010 / n) if n else 0.0,
            "pass": ok}


def level(v):
    """0 inside, 1 between .005 and .010, 2 over .010 - for colouring."""
    if v is None:
        return None
    a = abs(v)
    if a > TOL_OUT + EPS:
        return 2
    if a > TOL_IN + EPS:
        return 1
    return 0


def load(path):
    """The whole report: {"path", "info", "platforms", "points", "pass"}."""
    if not os.path.exists(path):
        raise ReadError("Dosya yok: %s" % path)
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise ReadError("Excel açılamadı (.xlsx olmalı): %s" % e)
    try:
        ws = _find_sheet(wb, "platform")
        if ws is None:
            raise ReadError("'Platform' sayfası yok.  Sayfalar: %s" % ", ".join(wb.sheetnames))
        plats, npts = read_platforms(_rows(ws))
        sm = _find_sheet(wb, "summary")
        if sm is None:
            sm = wb[wb.sheetnames[0]]
        info = read_info(_rows(sm))
    finally:
        try:
            wb.close()
        except Exception:
            pass
    for p in plats:
        p.update(judge(p["values"]))
    return {"path": path, "info": info, "platforms": plats, "points": npts,
            "pass": all(p["pass"] for p in plats)}


# ---------------------------------------------------------------------------
#  Excel output
# ---------------------------------------------------------------------------

PASS_FILL = "00B050"
FAIL_FILL = "FF0000"
LEVEL_FILL = {0: "C6EFCE", 1: "FFD966", 2: "FF7C80"}


def default_name(rep):
    i = rep["info"]
    parts = [x for x in (i.get("part"), i.get("serial")) if x]
    base = "_".join(parts) if parts else os.path.splitext(os.path.basename(rep["path"]))[0]
    base = re.sub(r'[\\/:*?"<>|]+', "_", base).strip(" ._") or "RAPOR"
    return "%s_PLATFORM.xlsx" % base


def export_xlsx(rep, path):
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    thin = Side(style="thin", color="000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")
    bold = Font(name="Calibri", size=11, bold=True)
    norm = Font(name="Calibri", size=11)
    head_fill = PatternFill("solid", fgColor="D9E1F2")

    def fill(c):
        return PatternFill("solid", fgColor=c)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PLATFORM"
    heads = ("PLATFORM NO.", "OUT PROFILE .010\nPOINT COUNT", "OUT PROFILE .020\nPOINT COUNT",
             "PASS/FAIL", "AVERAGE\n(±.004)", "TOTAL POINT\nCOUNT", "%5 LIMIT", "MAX |DEV|",
             "FEATURE")
    for c, h in enumerate(heads, 1):
        x = ws.cell(row=1, column=c, value=h)
        x.font, x.alignment, x.border, x.fill = bold, center, box, head_fill
    ws.row_dimensions[1].height = 36
    for i, p in enumerate(rep["platforms"], 2):
        vals = ("PLATFORM #%d" % p["no"], p["c010"], p["c020"],
                "PASS" if p["pass"] else "FAIL", round(p["avg"], 5), p["n"],
                p["limit"], round(p["max"], 4), p["name"])
        for c, v in enumerate(vals, 1):
            x = ws.cell(row=i, column=c, value=v)
            x.font, x.border = norm, box
            x.alignment = left if c in (1, 9) else center
        ws.cell(row=i, column=5).number_format = "0.0000"
        ws.cell(row=i, column=8).number_format = "0.0000"
        if not p["avg_ok"]:
            ws.cell(row=i, column=5).fill = fill(LEVEL_FILL[2])
        pf = ws.cell(row=i, column=4)
        pf.fill = fill(PASS_FILL if p["pass"] else FAIL_FILL)
        pf.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        if p["c020"]:
            ws.cell(row=i, column=3).fill = fill(LEVEL_FILL[2])
        if p["c010"]:
            ws.cell(row=i, column=2).fill = fill(LEVEL_FILL[1] if p["c010"] < p["limit"]
                                                 else LEVEL_FILL[2])
    r = len(rep["platforms"]) + 3
    ok = sum(1 for p in rep["platforms"] if p["pass"])
    x = ws.cell(row=r, column=1, value="GENEL SONUÇ")
    x.font, x.border = bold, box
    x = ws.cell(row=r, column=4, value="PASS" if rep["pass"] else "FAIL")
    x.fill = fill(PASS_FILL if rep["pass"] else FAIL_FILL)
    x.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    x.alignment, x.border = center, box
    ws.cell(row=r, column=2, value="%d / %d platform PASS" % (ok, len(rep["platforms"]))).font = norm
    ws.cell(row=r + 1, column=1,
            value="Kural: |sapma| ≤ .005 içeride;  .005 < |sapma| ≤ .010 nokta sayısı "
                  "< toplam noktanın %5'i (tam sayıya yuvarlanır);  |sapma| > .010 nokta yok;  |ortalama| ≤ .004.").font = \
        Font(name="Calibri", size=9, italic=True, color="595959")
    for c, w in enumerate((18, 20, 20, 14, 12, 13, 10, 11, 30), 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "B2"

    # part information
    wi = wb.create_sheet("PARÇA BİLGİSİ")
    for i, (lab, key) in enumerate(INFO_FIELDS, 1):
        a = wi.cell(row=i, column=1, value=lab)
        a.font, a.border, a.fill = bold, box, head_fill
        b = wi.cell(row=i, column=2, value=rep["info"].get(key, ""))
        b.font, b.border = norm, box
    n = len(INFO_FIELDS)
    wi.cell(row=n + 2, column=1, value="Kaynak dosya").font = bold
    wi.cell(row=n + 2, column=2, value=os.path.basename(rep["path"])).font = norm
    wi.cell(row=n + 3, column=1, value="Program").font = bold
    wi.cell(row=n + 3, column=2, value="%s %s  ·  %s" % (APP_TITLE, VERSION, APP_CREDIT)).font = norm
    wi.column_dimensions["A"].width = 16
    wi.column_dimensions["B"].width = 44

    # every point, coloured
    wp = wb.create_sheet("NOKTALAR")
    x = wp.cell(row=1, column=1, value="NOKTA")
    x.font, x.alignment, x.border, x.fill = bold, center, box, head_fill
    for j, p in enumerate(rep["platforms"], 2):
        x = wp.cell(row=1, column=j, value="#%d" % p["no"])
        x.font, x.alignment, x.border = bold, center, box
        x.fill = fill(PASS_FILL if p["pass"] else FAIL_FILL)
        x.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        wp.column_dimensions[get_column_letter(j)].width = 9
    for k in range(rep["points"]):
        x = wp.cell(row=k + 2, column=1, value=k + 1)
        x.font, x.alignment, x.border = bold, center, box
        for j, p in enumerate(rep["platforms"], 2):
            v = p["values"][k]
            if v is None:
                continue
            x = wp.cell(row=k + 2, column=j, value=v)
            x.number_format = "0.0000"
            x.font, x.alignment, x.border = norm, center, box
            x.fill = fill(LEVEL_FILL[level(v)])
    wp.freeze_panes = "B2"
    wb.save(path)
