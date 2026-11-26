# -*- coding: utf-8 -*-
"""PLATFORM DEGERLENDIRME - load a Hexagon report (.xlsx), see the part
information and the PASS / FAIL of every platform, export it to Excel.
Same look as WIP TRACKER."""

import os
import sys

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QSettings, Qt, QTimer
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QFileDialog, QFrame,
                               QGridLayout, QHBoxLayout, QHeaderView, QLabel,
                               QMainWindow, QMessageBox, QPushButton, QSplitter,
                               QStatusBar, QTableView, QVBoxLayout, QWidget)

import plat_core as core
import plat_theme

TH = plat_theme.theme("night")

# PASS / FAIL and the point levels: the same strong colours in both themes
PASS_BG, PASS_FG = "#1E9E4A", "#FFFFFF"
FAIL_BG, FAIL_FG = "#D32F2F", "#FFFFFF"


def lvl_colors():
    if TH["name"] == "day":
        return {0: ("#D5F0DC", "#14532D"), 1: ("#FBE3A0", "#6B4A00"), 2: ("#F6C1BE", "#8E1F1A")}
    return {0: ("#1F3D29", "#CFF2DA"), 1: ("#5A4A1C", "#FFE7A6"), 2: ("#5A2A27", "#FFD6D2")}


def label(text="", obj=None, muted=False):
    w = QLabel(text)
    if obj:
        w.setObjectName(obj)
    if muted:
        w.setProperty("muted", True)
    return w


def button(text, slot=None, primary=False, tip=None):
    b = QPushButton(text)
    b.setCursor(Qt.PointingHandCursor)
    if primary:
        b.setProperty("primary", True)
    if tip:
        b.setToolTip(tip)
    if slot:
        b.clicked.connect(lambda _=False: slot())
    return b


def draw_mark(p, x, y, s, col):
    """The mark: a measured profile inside its tolerance band."""
    p.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor(col), max(1.2, s / 18))
    pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)
    p.drawLine(int(x + s * .08), int(y + s * .28), int(x + s * .92), int(y + s * .28))
    p.drawLine(int(x + s * .08), int(y + s * .72), int(x + s * .92), int(y + s * .72))
    pen.setWidthF(max(1.6, s / 11))
    p.setPen(pen)
    pts = [(.08, .52), (.25, .40), (.42, .58), (.58, .44), (.75, .56), (.92, .46)]
    for (a, b), (c, d) in zip(pts, pts[1:]):
        p.drawLine(int(x + s * a), int(y + s * b), int(x + s * c), int(y + s * d))


def app_icon():
    pm = QPixmap(64, 64)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#1F63D6"))
    p.drawRoundedRect(0, 0, 64, 64, 14, 14)
    draw_mark(p, 6, 6, 52, "#FFFFFF")
    p.end()
    return QIcon(pm)


class Mark(QWidget):
    def __init__(self, size=40):
        super().__init__()
        self.setFixedSize(size, size)

    def paintEvent(self, ev):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(TH["accent"]))
        p.drawRoundedRect(self.rect(), 9, 9)
        s = self.width()
        draw_mark(p, s * .1, s * .1, s * .8, "#FFFFFF")
        p.end()


class ThemeButton(QPushButton):
    def __init__(self, slot):
        super().__init__("")
        self.setProperty("icon", True)
        self.setFixedSize(40, 40)
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip("Gündüz / gece")
        self.clicked.connect(lambda _=False: slot())

    def paintEvent(self, ev):
        super().paintEvent(ev)
        import math
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        pen = QPen(QColor(TH["text"]), 1.4)
        pen.setCapStyle(Qt.RoundCap)
        p.setPen(pen)
        cx, cy = self.width() / 2, self.height() / 2
        if TH["name"] == "night":
            p.drawEllipse(int(cx - 4), int(cy - 4), 8, 8)
            for k in range(8):
                a = k * math.pi / 4
                p.drawLine(int(cx + 7 * math.cos(a)), int(cy + 7 * math.sin(a)),
                           int(cx + 9.5 * math.cos(a)), int(cy + 9.5 * math.sin(a)))
        else:
            from PySide6.QtGui import QPainterPath
            a = QPainterPath()
            a.addEllipse(cx - 7, cy - 7, 14, 14)
            b = QPainterPath()
            b.addEllipse(cx - 3, cy - 10, 14, 14)
            p.drawPath(a.subtracted(b))
        p.end()


# ---------------------------------------------------------------------------
#  tables
# ---------------------------------------------------------------------------

def _f4(v):
    return "%.4f" % v


class ResultModel(QAbstractTableModel):
    HEADS = ("PLATFORM NO.", "SONUÇ", "OUT .010\n(.005–.010)", "OUT .020\n(> .010)",
             "ORTALAMA\n(±.004)", "NOKTA", "%5 SINIR", "MAX |DEV|", "ÖZELLİK")

    def __init__(self):
        super().__init__()
        self.rows = []

    def set(self, rows):
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(self.HEADS)

    def headerData(self, sec, orient, role=Qt.DisplayRole):
        if orient == Qt.Horizontal and role == Qt.DisplayRole:
            return self.HEADS[sec]
        return None

    def data(self, idx, role=Qt.DisplayRole):
        p = self.rows[idx.row()]
        c = idx.column()
        lv = lvl_colors()
        if role == Qt.DisplayRole:
            avg = p["avg"] if abs(p["avg"]) >= 5e-6 else 0.0
            return ("PLATFORM #%d" % p["no"], "PASS" if p["pass"] else "FAIL",
                    p["c010"], p["c020"], "%.5f" % avg, p["n"],
                    p["limit"], _f4(p["max"]), p["name"])[c]
        if role == Qt.TextAlignmentRole:
            return int((Qt.AlignLeft if c in (0, 8) else Qt.AlignHCenter) | Qt.AlignVCenter)
        if role == Qt.BackgroundRole:
            if c == 1:
                return QColor(PASS_BG if p["pass"] else FAIL_BG)
            if c == 2 and p["c010"]:
                return QColor(lv[1 if p["c010"] < p["limit"] else 2][0])
            if c == 3 and p["c020"]:
                return QColor(lv[2][0])
            if c == 4 and not p["avg_ok"]:
                return QColor(lv[2][0])
        if role == Qt.ForegroundRole:
            if c == 1:
                return QColor(PASS_FG)
            if c == 2 and p["c010"]:
                return QColor(lv[1 if p["c010"] < p["limit"] else 2][1])
            if c == 3 and p["c020"]:
                return QColor(lv[2][1])
            if c == 4 and not p["avg_ok"]:
                return QColor(lv[2][1])
        if role == Qt.FontRole and c in (0, 1):
            f = QFont(plat_theme.FONT, 10)
            f.setBold(True)
            return f
        if role == Qt.ToolTipRole:
            why = []
            if p["c020"]:
                why.append("%d nokta ±.010 dışında" % p["c020"])
            if p["c010"] >= p["limit"]:
                why.append(".005–.010 arası %d nokta (sınır < %d)" % (p["c010"], p["limit"]))
            if not p["avg_ok"]:
                why.append("ortalama %.5f  (±.004 dışında)" % p["avg"])
            return ("FAIL: " + ";  ".join(why)) if why else \
                "PASS  -  .005–.010 arası %d / %d nokta (%%%.1f), ortalama %.5f" % (
                    p["c010"], p["n"], p["pct"], p["avg"])
        return None


class PointModel(QAbstractTableModel):
    def __init__(self):
        super().__init__()
        self.plats = []
        self.n = 0

    def set(self, plats, n):
        self.beginResetModel()
        self.plats, self.n = plats, n
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return self.n

    def columnCount(self, parent=QModelIndex()):
        return len(self.plats)

    def headerData(self, sec, orient, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return ("#%d" % self.plats[sec]["no"]) if orient == Qt.Horizontal else str(sec + 1)
        if orient == Qt.Horizontal and role == Qt.ToolTipRole:
            p = self.plats[sec]
            return "PLATFORM #%d  %s  -  %s" % (p["no"], p["name"], "PASS" if p["pass"] else "FAIL")
        if orient == Qt.Horizontal and role == Qt.ForegroundRole:
            return QColor(PASS_BG if self.plats[sec]["pass"] else FAIL_BG)
        return None

    def data(self, idx, role=Qt.DisplayRole):
        v = self.plats[idx.column()]["values"][idx.row()]
        if v is None:
            return None
        if role == Qt.DisplayRole:
            return _f4(v)
        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignCenter)
        lv = core.level(v)
        if role == Qt.BackgroundRole and lv:
            return QColor(lvl_colors()[lv][0])
        if role == Qt.ForegroundRole and lv:
            return QColor(lvl_colors()[lv][1])
        return None


def make_table(model):
    t = QTableView()
    t.setModel(model)
    t.setFrameShape(QFrame.NoFrame)
    t.setAlternatingRowColors(True)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.verticalHeader().setDefaultSectionSize(28)
    t.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
    return t


# ---------------------------------------------------------------------------
#  window
# ---------------------------------------------------------------------------

class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        global TH
        self.settings = QSettings("EYALTIRIKLI", "PLATFORM_DEGERLENDIRME")
        TH = plat_theme.theme(self.settings.value("theme", "night"))
        self.setWindowTitle("%s  v%s" % (core.APP_TITLE, core.VERSION))
        self.setWindowIcon(app_icon())
        self.resize(1400, 860)
        self.setAcceptDrops(True)
        self.rep = None

        root = QWidget()
        root.setObjectName("Page")
        self.setCentralWidget(root)
        v = QVBoxLayout(root)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        v.addWidget(self._top_bar())
        body = QWidget()
        body.setObjectName("Body")
        b = QVBoxLayout(body)
        b.setContentsMargins(18, 14, 18, 12)
        b.setSpacing(10)
        v.addWidget(body, 1)

        head = QHBoxLayout()
        tl = QVBoxLayout()
        tl.setSpacing(2)
        tl.addWidget(label("Platform değerlendirme", obj="PageTitle"))
        self.sub = label("Hexagon raporu yükleyin (.xlsx) - ya da dosyayı pencereye sürükleyin",
                         obj="PageSub")
        tl.addWidget(self.sub)
        head.addLayout(tl)
        head.addStretch(1)
        self.export_btn = button("Excel'e aktar", self.export)
        self.export_btn.setEnabled(False)
        head.addWidget(self.export_btn)
        head.addWidget(button("Excel yükle", self.choose, primary=True))
        b.addLayout(head)

        # part information + overall verdict
        top = QHBoxLayout()
        top.setSpacing(10)
        info = QFrame()
        info.setProperty("card", True)
        g = QGridLayout(info)
        g.setContentsMargins(16, 12, 16, 12)
        g.setHorizontalSpacing(26)
        g.setVerticalSpacing(8)
        self.info_lbl = {}
        layout = (("part", 0, 0, 2), ("serial", 0, 2, 2), ("operation", 0, 4, 1),
                  ("operator", 1, 0, 1), ("cmm", 1, 1, 1), ("ncws", 1, 2, 1),
                  ("spindle", 1, 3, 1), ("date", 1, 4, 1),
                  ("company", 2, 0, 2), ("location", 2, 2, 3))
        for key, r, c, span in layout:
            box = QVBoxLayout()
            box.setSpacing(0)
            box.addWidget(label(core.INFO_LABELS[key].upper(), muted=True))
            val = label("—")
            val.setTextInteractionFlags(Qt.TextSelectableByMouse)
            size = "15pt" if key in ("part", "serial") else "11pt"
            val.setStyleSheet("font-size:%s;font-weight:700;" % size)
            box.addWidget(val)
            self.info_lbl[key] = val
            g.addLayout(box, r, c, 1, span)
        for c in range(5):
            g.setColumnStretch(c, 1)
        top.addWidget(info, 1)

        self.verdict = QFrame()
        self.verdict.setProperty("card", True)
        self.verdict.setFixedWidth(260)
        vv = QVBoxLayout(self.verdict)
        vv.setContentsMargins(16, 12, 16, 12)
        vv.addWidget(label("GENEL SONUÇ", muted=True), 0, Qt.AlignHCenter)
        self.verdict_big = label("—")
        self.verdict_big.setAlignment(Qt.AlignCenter)
        vv.addWidget(self.verdict_big, 1)
        self.verdict_sub = label("", muted=True)
        self.verdict_sub.setAlignment(Qt.AlignCenter)
        vv.addWidget(self.verdict_sub)
        top.addWidget(self.verdict)
        b.addLayout(top)

        rule = label("Kural (her platform):  |sapma| ≤ .005 içeride  ·  .005 < |sapma| ≤ .010 "
                     "nokta sayısı < toplam noktanın %5'i (yuvarlanmış)  ·  |sapma| > .010 nokta yok  ·  "
                     "|ortalama| ≤ .004.   Tam .004 / .005 / .010 kabul.", muted=True)
        rule.setWordWrap(True)
        b.addWidget(rule)

        self.rm = ResultModel()
        self.pm = PointModel()
        self.rt = make_table(self.rm)
        self.pt = make_table(self.pm)
        self.rt.verticalHeader().hide()
        self.rt.horizontalHeader().setFixedHeight(46)
        self.pt.horizontalHeader().setFixedHeight(46)
        self.pt.horizontalHeader().setDefaultSectionSize(70)
        self.pt.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.rt.selectionModel().currentRowChanged.connect(self._goto_point_col)
        split = QSplitter(Qt.Horizontal)
        for t, title in ((self.rt, "PLATFORMLAR"), (self.pt, "NOKTALAR  (yukarıdan aşağı nokta, "
                                                             "soldan sağa platform)")):
            card = QFrame()
            card.setProperty("card", True)
            cl = QVBoxLayout(card)
            cl.setContentsMargins(1, 8, 1, 1)
            cl.setSpacing(6)
            h = label(title)
            h.setProperty("section", True)
            h.setContentsMargins(12, 0, 0, 0)
            cl.addWidget(h)
            cl.addWidget(t, 1)
            split.addWidget(card)
        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 2)
        self.split = split
        self.empty = label("Henüz rapor yok.  Sağ üstten  Excel yükle.", muted=True)
        self.empty.setAlignment(Qt.AlignCenter)
        b.addWidget(self.empty, 1)
        b.addWidget(split, 1)
        split.hide()

        sb = QStatusBar()
        self.setStatusBar(sb)
        self.status = QLabel("Hazır.")
        sb.addWidget(self.status, 1)
        self.apply_theme()

    def _top_bar(self):
        bar = QFrame()
        bar.setObjectName("TopBar")
        bar.setFixedHeight(62)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(14, 6, 16, 6)
        lay.setSpacing(12)
        self.mark = Mark(40)
        lay.addWidget(self.mark)
        sep = QFrame()
        sep.setFixedSize(1, 34)
        sep.setObjectName("TopSep")
        lay.addWidget(sep)
        txt = QVBoxLayout()
        txt.setSpacing(0)
        txt.addWidget(label(core.APP_TITLE, obj="AppTitle"))
        txt.addWidget(label("v%s  ·  %s" % (core.VERSION, core.APP_CREDIT), obj="AppCredit"))
        lay.addLayout(txt)
        lay.addStretch(1)
        self.file_lbl = label("", obj="LastUpdated")
        lay.addWidget(self.file_lbl)
        self.theme_btn = ThemeButton(self.toggle_theme)
        lay.addWidget(self.theme_btn)
        return bar

    # ---- theme ----
    def toggle_theme(self):
        global TH
        TH = plat_theme.theme("day" if TH["name"] == "night" else "night")
        self.apply_theme()
        self.settings.setValue("theme", TH["name"])

    def apply_theme(self):
        QApplication.instance().setPalette(plat_theme.palette(TH))
        self.setStyleSheet(plat_theme.stylesheet(TH))
        self._paint_verdict()
        for t in (self.rt, self.pt):
            t.viewport().update()
        self.mark.update()
        self.theme_btn.update()

    def _paint_verdict(self):
        if not self.rep:
            self.verdict_big.setStyleSheet("font-size:30pt;font-weight:800;color:%s;"
                                           % TH["muted"])
            self.verdict.setStyleSheet("")
            return
        ok = self.rep["pass"]
        bg = PASS_BG if ok else FAIL_BG
        self.verdict.setStyleSheet("QFrame[card=\"true\"]{background:%s;border:none;"
                                   "border-radius:10px;} QLabel{color:#FFFFFF;}" % bg)
        self.verdict_big.setStyleSheet("font-size:34pt;font-weight:800;color:#FFFFFF;")

    # ---- loading ----
    def dragEnterEvent(self, ev):
        if ev.mimeData().hasUrls():
            ev.acceptProposedAction()

    def dropEvent(self, ev):
        for u in ev.mimeData().urls():
            p = u.toLocalFile()
            if p.lower().endswith((".xlsx", ".xlsm")):
                self.load(p)
                return

    def choose(self):
        start = self.settings.value("dir", os.path.expanduser("~"))
        path, _f = QFileDialog.getOpenFileName(self, "Hexagon raporu seçin", start,
                                               "Excel (*.xlsx *.xlsm);;Tümü (*)")
        if path:
            self.settings.setValue("dir", os.path.dirname(path))
            self.load(path)

    def load(self, path):
        self.status.setText("Okunuyor: %s" % path)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            rep = core.load(path)
        except core.ReadError as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "Okunamadı", str(e))
            self.status.setText("Okunamadı.")
            return
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "Okunamadı", "%s: %s" % (type(e).__name__, e))
            self.status.setText("Okunamadı.")
            return
        QApplication.restoreOverrideCursor()
        self.show_report(rep)

    def show_report(self, rep):
        self.rep = rep
        for k, w in self.info_lbl.items():
            w.setText(rep["info"].get(k) or "—")
        plats = rep["platforms"]
        ok = sum(1 for p in plats if p["pass"])
        self.verdict_big.setText("PASS" if rep["pass"] else "FAIL")
        self.verdict_sub.setText("%d / %d platform PASS" % (ok, len(plats)))
        self._paint_verdict()
        self.rm.set(plats)
        self.pm.set(plats, rep["points"])
        hh = self.rt.horizontalHeader()
        for c in range(self.rm.columnCount()):
            hh.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(self.rm.columnCount() - 1, QHeaderView.Stretch)
        self.empty.hide()
        self.split.show()
        self.export_btn.setEnabled(True)
        name = os.path.basename(rep["path"])
        self.sub.setText("%s  ·  %d platform  ·  %d nokta sırası" % (name, len(plats), rep["points"]))
        self.file_lbl.setText("Dosya: <b>%s</b>" % name)
        self.status.setText("Okundu: %s" % rep["path"])

    def _goto_point_col(self, cur, _prev):
        if cur.isValid() and self.pm.columnCount():
            self.pt.scrollTo(self.pm.index(0, cur.row()))
            self.pt.selectColumn(cur.row())

    # ---- export ----
    def export(self):
        if not self.rep:
            return
        start = self.settings.value("out_dir", os.path.dirname(self.rep["path"]))
        path, _f = QFileDialog.getSaveFileName(
            self, "Excel'e aktar", os.path.join(start, core.default_name(self.rep)),
            "Excel (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            core.export_xlsx(self.rep, path)
        except Exception as e:
            QMessageBox.warning(self, "Kaydedilemedi", str(e))
            return
        self.settings.setValue("out_dir", os.path.dirname(path))
        self.status.setText("Kaydedildi: %s" % path)


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "EYALTIRIKLI.PlatformDegerlendirme")
        except Exception:
            pass
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont(plat_theme.FONT, 10))
    app.setWindowIcon(app_icon())
    w = Window()
    w.show()
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        QTimer.singleShot(0, lambda: w.load(sys.argv[1]))
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
