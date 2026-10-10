# -*- coding: utf-8 -*-
"""ŠIHTARICA IDSS/IMH – desktop aplikacija (Windows, offline). Qt/PySide6."""
import json, os, sys, traceback, datetime as dt
from pathlib import Path
from PySide6.QtCore import Qt, QThread, Signal, QSize, QUrl
from PySide6.QtGui import QColor, QFont, QDesktopServices, QIcon, QPixmap, QPainter, QBrush
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QGridLayout, QStackedWidget, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QTabWidget,
    QLineEdit, QSpinBox, QRadioButton, QButtonGroup, QCheckBox, QTreeWidget, QTreeWidgetItem, QSplitter, QProgressBar,
    QAbstractItemView, QSizePolicy, QGraphicsDropShadowEffect, QStyledItemDelegate, QDialog)
import engine as E
import izvoz as I

# Windows konzola koristi cp1252 (bez č/ć/đ/š/ž): ispis nikad ne smije srušiti aplikaciju
for _tok in (sys.stdout, sys.stderr):
    try: _tok.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

APP = "Šihtarica IDSS · IMH"
CFG = Path(os.environ.get("APPDATA", Path.home())) / "SihtaricaIDSS" / "postavke.json"
def resurs(ime): return Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / ime

PLAVA, TAMNA, AKCENT = "#035EA1", "#0B2540", "#08ABE6"
QSS = f"""
* {{ font-family: 'Segoe UI', 'Carlito', sans-serif; font-size: 10pt; color: #1F2D3A; }}
QMainWindow, #sadrzaj {{ background: #F3F6F9; }}
#meni {{ background: {TAMNA}; }}
#meni QLabel {{ color: #C9D6E3; }}
#logo {{ color: white; font-size: 17pt; font-weight: 700; }}
#podlogo {{ color: #7FA6C9; font-size: 9pt; letter-spacing: 1px; }}
#nav {{ outline: none; background: transparent; color: #C9D6E3; border: none; text-align: left; padding: 11px 18px; font-size: 10.5pt; border-left: 3px solid transparent; }}
#nav:hover {{ background: #12345A; color: white; }}
#nav:checked {{ background: #133A61; color: white; border-left: 3px solid {AKCENT}; font-weight: 600; }}
#fajlKartica {{ background: #12345A; border-radius: 8px; }}
#fajlNaslov {{ color: #7FA6C9; font-size: 8.5pt; }}
#fajlIme {{ color: white; font-size: 9pt; font-weight: 600; }}
#naslov {{ font-size: 18pt; font-weight: 700; color: {TAMNA}; }}
#podnaslov {{ color: #5B6B7A; font-size: 10pt; }}
#kartica {{ background: white; border: 1px solid #E1E8EF; border-radius: 10px; }}
#kpiBroj {{ font-size: 20pt; font-weight: 700; color: {TAMNA}; }}
#kpiOznaka {{ color: #6B7C8C; font-size: 8.5pt; text-transform: uppercase; }}
#kpiUpoz {{ font-size: 20pt; font-weight: 700; color: #B26A00; }}
QPushButton {{ background: white; border: 1px solid #C9D4DE; border-radius: 7px; padding: 7px 14px; }}
QPushButton:hover {{ border-color: {PLAVA}; color: {PLAVA}; }}
QPushButton#primarno {{ background: {PLAVA}; color: white; border: none; font-weight: 600; padding: 10px 22px; font-size: 10.5pt; }}
QPushButton#primarno:hover {{ background: #024B80; }}
QPushButton#primarno:disabled {{ background: #A9BCCB; }}
QTableWidget, QTreeWidget {{ background: white; border: 1px solid #E1E8EF; border-radius: 8px; gridline-color: #EEF2F6;
    selection-background-color: #D6E8F7; selection-color: #1F2D3A; alternate-background-color: #F7FAFC; }}
QHeaderView::section {{ background: #EAF1F7; color: #33475B; border: none; border-right: 1px solid #DDE5EC; padding: 6px; font-weight: 600; font-size: 9pt; }}
QTabWidget::pane {{ border: none; }}
QTabBar::tab {{ background: transparent; padding: 8px 16px; color: #5B6B7A; border-bottom: 2px solid transparent; }}
QTabBar::tab:selected {{ color: {PLAVA}; border-bottom: 2px solid {PLAVA}; font-weight: 600; }}
QComboBox, QLineEdit, QSpinBox {{ background: white; border: 1px solid #C9D4DE; border-radius: 6px; padding: 5px 8px; }}
QComboBox:focus, QLineEdit:focus, QSpinBox:focus {{ border-color: {PLAVA}; }}
QProgressBar {{ border: none; background: #E1E8EF; border-radius: 3px; height: 6px; }}
QProgressBar::chunk {{ background: {PLAVA}; border-radius: 3px; }}
#traka {{ background: white; border-top: 1px solid #E1E8EF; }}
#info {{ color: #5B6B7A; }}
#uspjeh {{ background: #E8F5EC; border: 1px solid #B9E0C5; border-radius: 8px; color: #1B5E20; padding: 10px; }}
#greska {{ background: #FDECEC; border: 1px solid #F3C2C2; border-radius: 8px; color: #8E1B1B; padding: 10px; }}
#savjet {{ background: #FFF8E1; border: 1px solid #F5DE9A; border-radius: 8px; color: #6D4C00; padding: 10px; }}
"""

def ucitaj_postavke():
    try: p = json.loads(CFG.read_text(encoding="utf-8"))
    except Exception: p = {}
    p.setdefault("tumacenje", dict(E.PODRAZUMIJEVANE_POSTAVKE)); return p
def sacuvaj_postavke(p):
    CFG.parent.mkdir(parents=True, exist_ok=True); CFG.write_text(json.dumps(p, ensure_ascii=False, indent=2), encoding="utf-8")

# ------------------------------------------------------------------ pomoćni widgeti
class Radnik(QThread):
    napredak = Signal(str); gotovo = Signal(object); greska = Signal(str)
    def __init__(self, fn, *a, **k): super().__init__(); self.fn, self.a, self.k = fn, a, k
    def run(self):
        try: self.gotovo.emit(self.fn(*self.a, **self.k))
        except (E.GreskaFormata, E.GreskaUpisa) as e: self.greska.emit(str(e))
        except PermissionError: self.greska.emit("Fajl je otvoren u drugom programu (npr. Excelu). Zatvorite ga i pokušajte ponovo.")
        except Exception: self.greska.emit("Neočekivana greška. Evidencija nije promijenjena.\n\n" + traceback.format_exc(limit=4))

def kartica():
    f = QFrame(); f.setObjectName("kartica")
    sj = QGraphicsDropShadowEffect(blurRadius=18, xOffset=0, yOffset=2); sj.setColor(QColor(16, 42, 67, 22)); f.setGraphicsEffect(sj)
    return f

def lbl(t, ime=None, wrap=False):
    l = QLabel(t)
    if ime: l.setObjectName(ime)
    l.setWordWrap(wrap); return l

class KPI(QFrame):
    def __init__(self, oznaka):
        super().__init__(); self.setObjectName("kartica")
        v = QVBoxLayout(self); v.setContentsMargins(16, 12, 16, 12); v.setSpacing(2)
        self.broj = lbl("–", "kpiBroj"); v.addWidget(self.broj); v.addWidget(lbl(oznaka, "kpiOznaka"))
    def postavi(self, v, upozorenje=False):
        self.broj.setText(str(v)); self.broj.setObjectName("kpiUpoz" if upozorenje and v not in (0, "0") else "kpiBroj")
        self.broj.style().unpolish(self.broj); self.broj.style().polish(self.broj)

class ZonaZaFajl(QFrame):
    odabran = Signal(str)
    def __init__(self):
        super().__init__(); self.setAcceptDrops(True); self.setCursor(Qt.PointingHandCursor); self.setMinimumHeight(130)
        v = QVBoxLayout(self); v.setAlignment(Qt.AlignCenter)
        self.ikona = lbl("⇪"); self.ikona.setStyleSheet(f"font-size: 28pt; color: {PLAVA};"); self.ikona.setAlignment(Qt.AlignCenter)
        self.t1 = lbl("Prevucite ovdje „All Report” izvoz s uređaja"); self.t1.setStyleSheet("font-size: 12pt; font-weight: 600;")
        self.t2 = lbl("ili kliknite za odabir fajla  ·  .xls ili .xlsx", "podnaslov")
        for w in (self.ikona, self.t1, self.t2): w.setAlignment(Qt.AlignCenter); v.addWidget(w)
        self._stil(False)
    def kompaktno(self, da):
        self.ikona.setVisible(not da); self.setMinimumHeight(64 if da else 130); self.setMaximumHeight(72 if da else 16777215)
        self.t1.setText("Učitajte drugi izvoz – prevucite ili kliknite" if da else "Prevucite ovdje „All Report” izvoz s uređaja")
        self.t2.setVisible(not da)
    def _stil(self, aktivno):
        self.setStyleSheet(f"ZonaZaFajl {{ background: {'#E8F2FB' if aktivno else 'white'}; border: 2px dashed {PLAVA if aktivno else '#B7C7D6'}; border-radius: 12px; }}")
    def mousePressEvent(self, e):
        p, _ = QFileDialog.getOpenFileName(self, "Odaberite All Report izvoz", "", "Izvoz uređaja (*.xls *.xlsx)")
        if p: self.odabran.emit(p)
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls() and e.mimeData().urls()[0].toLocalFile().lower().endswith((".xls", ".xlsx")):
            e.acceptProposedAction(); self._stil(True)
    def dragLeaveEvent(self, e): self._stil(False)
    def dropEvent(self, e):
        self._stil(False); self.odabran.emit(e.mimeData().urls()[0].toLocalFile())

class Obavijest(QLabel):
    def __init__(self):
        super().__init__(); self.setWordWrap(True); self.hide(); self.setTextInteractionFlags(Qt.TextSelectableByMouse); self.setMargin(10)
    def pokazi(self, tekst, vrsta="uspjeh"):
        self.setObjectName(vrsta); self.setText(tekst); self.style().unpolish(self); self.style().polish(self); self.show()

def tabela(kolone, rastegni=None):
    t = QTableWidget(0, len(kolone)); t.setHorizontalHeaderLabels(kolone)
    t.verticalHeader().setVisible(False); t.setAlternatingRowColors(True); t.setShowGrid(False)
    t.setSelectionBehavior(QAbstractItemView.SelectRows); t.verticalHeader().setDefaultSectionSize(30)
    h = t.horizontalHeader(); h.setHighlightSections(False)
    for i in range(len(kolone)): h.setSectionResizeMode(i, QHeaderView.ResizeToContents)
    if rastegni is not None: h.setSectionResizeMode(rastegni, QHeaderView.Stretch)
    return t

def stavka(t, uredivo=False, boja=None):
    it = QTableWidgetItem("" if t is None else str(t))
    if not uredivo: it.setFlags(it.flags() & ~Qt.ItemIsEditable)
    if boja: it.setBackground(QColor(boja))
    return it

class Zaglavlje(QWidget):
    def __init__(self, naslov, podnaslov):
        super().__init__(); v = QVBoxLayout(self); v.setContentsMargins(0, 0, 0, 6); v.setSpacing(2)
        v.addWidget(lbl(naslov, "naslov")); v.addWidget(lbl(podnaslov, "podnaslov", True))

# ------------------------------------------------------------------ stranica: UVOZ
class StranicaUvoz(QWidget):
    upisano = Signal()
    def __init__(self, glavni):
        super().__init__(); self.g = glavni; self.plan = None; self.radnik = None
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 0); v.setSpacing(14)
        v.addWidget(Zaglavlje("Uvoz s uređaja", "Izvoz „All Report” s uređaja FA01H prevodi se u propisanu šihtaricu. "
                              "Prevodi se sve što uređaj sadrži; ono što nedostaje ostaje prazno i označeno – ništa se ne izmišlja."))
        self.zona = ZonaZaFajl(); self.zona.odabran.connect(self.analiziraj); v.addWidget(self.zona)
        self.obavijest = Obavijest(); v.addWidget(self.obavijest)
        self.napredak = QProgressBar(); self.napredak.setRange(0, 0); self.napredak.setTextVisible(False); self.napredak.hide(); v.addWidget(self.napredak)
        self.lbl_napr = lbl("", "info"); self.lbl_napr.hide(); v.addWidget(self.lbl_napr)
        # rezultat analize
        self.rez = QWidget(); rv = QVBoxLayout(self.rez); rv.setContentsMargins(0, 0, 0, 0); rv.setSpacing(12)
        self.lbl_izvor = lbl("", "podnaslov", True); rv.addWidget(self.lbl_izvor)
        kp = QHBoxLayout(); kp.setSpacing(12)
        self.kpi = {k: KPI(n) for k, n in (("radnika", "Radnika"), ("novih", "Novih šihtarica"), ("dopuna", "Dopuna postojećih"),
                                             ("bez_kategorije", "Dani bez kategorije"), ("bez_satnice", "Bez satnice"), ("neslaganja", "Neslaganja u izvozu"))}
        for w in self.kpi.values(): kp.addWidget(w)
        rv.addLayout(kp)
        self.tabovi = QTabWidget()
        self.t_radnici = tabela(["ID", "Ime i prezime", "Ustanova", "Radno mjesto", "Sedmično (h)", "Odgovorno lice", "Odjel na uređaju", "Status"], 7)
        self.t_radnici.itemChanged.connect(self._radnik_izmijenjen)
        w = QWidget(); wl = QVBoxLayout(w); wl.setContentsMargins(0, 8, 0, 0)
        wl.addWidget(lbl("Dvoklikom uredite radno mjesto, sedmičnu satnicu i odgovorno lice. Podaci se trajno pamte u registru radnika. "
                         "Bez sedmične satnice odsustva se evidentiraju bez broja sati.", "info", True)); wl.addWidget(self.t_radnici)
        self.tabovi.addTab(w, "Radnici")
        w = QWidget(); wl = QVBoxLayout(w); wl.setContentsMargins(0, 8, 0, 0)
        red = QHBoxLayout(); red.addWidget(lbl("Označite redove i dodijelite kategoriju:", "info"))
        self.cb_kat = QComboBox(); self.cb_kat.addItems(list(E.KATEGORIJE)); self.cb_kat.setMinimumWidth(240); red.addWidget(self.cb_kat)
        b = QPushButton("Primijeni na označene"); b.clicked.connect(self._primijeni_kategoriju); red.addWidget(b); red.addStretch()
        wl.addLayout(red)
        wl.addWidget(lbl("Kategorija nije obavezna. Dan bez kategorije upisuje se kao „Nema podataka u izvozu uređaja” i može se kasnije dopuniti.", "info", True))
        self.t_kat = tabela(["ID", "Radnik", "Datum", "Dan", "Razlog", "Kategorija"], 5); wl.addWidget(self.t_kat)
        self.tabovi.addTab(w, "Dani bez prijava")
        w = QWidget(); wl = QVBoxLayout(w); wl.setContentsMargins(0, 8, 0, 0)
        wl.addWidget(lbl("Pomoćni listovi izvoza (raspored, zbirni pregled, nepravilnosti) koji se ne slažu s prijavama. Ne upisuju se u šihtaricu – "
                         "vrijede prijave. Sve se trajno bilježi u listu „Upozorenja”.", "info", True))
        self.t_nesl = tabela(["ID", "Datum", "Izvor", "Opis"], 3); wl.addWidget(self.t_nesl)
        self.tabovi.addTab(w, "Neslaganja u izvozu")
        rv.addWidget(self.tabovi, 1)
        self.rez.hide(); v.addWidget(self.rez, 1)
        v.addStretch(0)
        # akcijska traka
        traka = QFrame(); traka.setObjectName("traka"); th = QHBoxLayout(traka); th.setContentsMargins(0, 12, 0, 16)
        self.lbl_akcija = lbl("", "info", True); th.addWidget(self.lbl_akcija, 1)
        self.btn_upis = QPushButton("Prevedi i upiši u evidenciju"); self.btn_upis.setObjectName("primarno"); self.btn_upis.setEnabled(False)
        self.btn_upis.clicked.connect(self.upisi); th.addWidget(self.btn_upis)
        v.addWidget(traka)

    def _rad(self, poruka, fn, *a, gotovo, **k):
        self.napredak.show(); self.lbl_napr.setText(poruka); self.lbl_napr.show(); self.btn_upis.setEnabled(False); self.zona.setEnabled(False)
        self.radnik = Radnik(fn, *a, **k)
        self.radnik.napredak.connect(self.lbl_napr.setText)
        def kraj(): self.napredak.hide(); self.lbl_napr.hide(); self.zona.setEnabled(True)
        self.radnik.gotovo.connect(lambda r: (kraj(), gotovo(r)))
        self.radnik.greska.connect(lambda t: (kraj(), self.obavijest.pokazi(t, "greska"), self._osvjezi_dugme()))
        self.radnik.start()

    def analiziraj(self, putanja):
        ev = self.g.evidencija()
        if not ev: self.obavijest.pokazi("Prvo kreirajte ili otvorite evidencijski fajl (lijevo dolje).", "savjet"); return
        self.obavijest.hide()
        post = self.g.post["tumacenje"]
        self._rad("Čitam i analiziram izvoz…", lambda: E.napravi_plan(E.citaj_izvor(putanja, post), ev, post), gotovo=self.prikazi)

    def prikazi(self, plan):
        self.plan = plan; izv = plan["izvor"]; s = E.sazetak_plana(plan)
        self.lbl_izvor.setText(f"<b>{izv['naziv']}</b> &nbsp;·&nbsp; period {izv['od']:%d.%m.%Y} – {izv['do']:%d.%m.%Y} &nbsp;·&nbsp; format: {izv['format']}"
                               f" &nbsp;·&nbsp; {len(izv['zapisi'])} dnevnih zapisa")
        for k, w in self.kpi.items(): w.postavi(s[k], upozorenje=k in ("bez_kategorije", "bez_satnice", "neslaganja"))
        t = self.t_radnici; t.blockSignals(True); t.setRowCount(0)
        for r in plan["radnici"].values():
            i = t.rowCount(); t.insertRow(i)
            st = ("Novi radnik" if r["novi"] else "U registru") + ("" if E.norma_radnika(r) else " · nema satnice")
            boja = "#FFF6D5" if r["novi"] or not E.norma_radnika(r) else None
            for c, (v, ur) in enumerate(((r["id"], 0), (r["ime"], 1), (None, 0), (r.get("radno_mjesto"), 1), (r.get("sedmicno"), 1),
                                         (r.get("odgovorno_lice") or self.g.post["tumacenje"]["odgovorno_lice"].get(r["inst"], ""), 1), (r.get("odjel"), 0), (st, 0))):
                if c == 2: continue
                t.setItem(i, c, stavka(v, ur, boja if c == 7 else None))
            cb = QComboBox(); cb.addItems(["IDSS", "IMH"]); cb.setCurrentText(r["inst"])
            cb.currentTextChanged.connect(lambda txt, rid=r["id"]: plan["radnici"][rid].update(inst=txt))
            if any(st_["postojeci"] for st_ in plan["stavke"] if st_["id"] == r["id"]): cb.setEnabled(False)
            t.setCellWidget(i, 2, cb)
        t.blockSignals(False)
        k = self.t_kat; k.setRowCount(0)
        for n, x in enumerate(plan["kategorizacija"]):
            i = k.rowCount(); k.insertRow(i)
            for c, v in enumerate((x["id"], x["ime"], f"{x['datum']:%d.%m.%Y}", E.DANI[x["datum"].weekday()], x["razlog"])): k.setItem(i, c, stavka(v))
            cb = QComboBox(); cb.addItems(list(E.KATEGORIJE)); cb.setCurrentText(x["kategorija"])
            cb.currentTextChanged.connect(lambda txt, n=n: (plan["kategorizacija"][n].update(kategorija=txt), self._osvjezi_kpi()))
            k.setCellWidget(i, 5, cb)
        self.tabovi.setTabText(1, f"Dani bez prijava ({len(plan['kategorizacija'])})")
        n = self.t_nesl; n.setRowCount(0)
        for rid, d, izvor, op in plan["kontrola"]:
            i = n.rowCount(); n.insertRow(i)
            for c, v in enumerate((rid, f"{d:%d.%m.%Y}" if d else "", izvor, op)): n.setItem(i, c, stavka(v))
        self.tabovi.setTabText(2, f"Neslaganja u izvozu ({len(plan['kontrola'])})")
        self.rez.show(); self.zona.kompaktno(True); self._osvjezi_dugme()

    def _osvjezi_kpi(self):
        if self.plan:
            s = E.sazetak_plana(self.plan)
            for k in ("bez_kategorije", "bez_satnice"): self.kpi[k].postavi(s[k], True)

    def _radnik_izmijenjen(self, it):
        if not self.plan: return
        rid = self.t_radnici.item(it.row(), 0).text(); r = self.plan["radnici"][rid]; v = it.text().strip()
        if it.column() == 1 and v: r["ime"] = v
        elif it.column() == 3: r["radno_mjesto"] = v
        elif it.column() == 5: r["odgovorno_lice"] = v
        elif it.column() == 4:
            if v == "": r["sedmicno"] = None
            else:
                try:
                    x = float(v.replace(",", "."))
                    if not 0 < x <= 60: raise ValueError
                    r["sedmicno"] = int(x) if x.is_integer() else x
                except ValueError:
                    QMessageBox.warning(self, APP, "Sedmična satnica mora biti broj između 0 i 60 (npr. 40).")
                    self.t_radnici.blockSignals(True); it.setText("" if r.get("sedmicno") is None else str(r["sedmicno"])); self.t_radnici.blockSignals(False)
        st = ("Novi radnik" if r["novi"] else "U registru") + ("" if E.norma_radnika(r) else " · nema satnice")
        self.t_radnici.blockSignals(True); s_ = self.t_radnici.item(it.row(), 7); s_.setText(st)
        s_.setBackground(QColor("#FFF6D5") if r["novi"] or not E.norma_radnika(r) else QColor("white")); self.t_radnici.blockSignals(False)
        self._osvjezi_kpi()

    def _primijeni_kategoriju(self):
        redovi = sorted({i.row() for i in self.t_kat.selectedIndexes()})
        if not redovi: self.obavijest.pokazi("Označite jedan ili više redova u tabeli.", "savjet"); return
        for r in redovi: self.t_kat.cellWidget(r, 5).setCurrentText(self.cb_kat.currentText())

    def _osvjezi_dugme(self):
        p = self.plan
        if not p: self.btn_upis.setEnabled(False); self.lbl_akcija.setText(""); return
        if p["vec_uvezen"]:
            self.btn_upis.setEnabled(False); self.lbl_akcija.setText(f"Ovaj isti izvoz je već upisan ({p['vec_uvezen']}). Nema ničeg novog."); return
        s = E.sazetak_plana(p)
        self.btn_upis.setEnabled(True)
        self.lbl_akcija.setText(f"Biće kreirano <b>{s['novih']}</b> novih šihtarica i dopunjeno <b>{s['dopuna']}</b> postojećih. "
                                "Postojeći upisi se nikad ne mijenjaju; prije upisa pravi se sigurnosna kopija.")

    def upisi(self):
        ev = self.g.evidencija(); p = self.plan
        if not (ev and p): return
        s = E.sazetak_plana(p)
        if QMessageBox.question(self, APP, f"Upisati u evidenciju?\n\nNove šihtarice: {s['novih']}\nDopuna postojećih: {s['dopuna']}\n"
                                f"Dani bez kategorije: {s['bez_kategorije']}\n\nRaniji upisi ostaju nepromijenjeni.") != QMessageBox.Yes: return
        self.btn_upis.setEnabled(False)
        def gotovo(r):
            self.plan = None; self.rez.hide(); self.lbl_akcija.setText(""); self.zona.kompaktno(False)
            self.obavijest.pokazi(f"✔  Upisano: {r['novih']} novih šihtarica, {r['dopunjenih']} dopunjenih ({r['dana_dopunjeno']} dana). "
                                  f"Upozorenja zabilježena: {r['upozorenja']}. Istorija provjerena ćeliju po ćeliju – nepromijenjena.\n"
                                  f"Sigurnosna kopija prethodnog stanja: {Path(r['kopija']).name}", "uspjeh")
            self.upisano.emit()
        self._rad("Upisujem…", E.upisi, p, ev, gotovo=gotovo, napredak=lambda t: self.radnik.napredak.emit(t))

# ------------------------------------------------------------------ stranica: EVIDENCIJA
KOL_OBRASCA = ["R.br.", "Datum", "Dan", "Početak\nrada", "Završetak\nrada", "Početak\npauze", "Završetak\npauze", "Ukupno\nsati",
               "Prekov.", "Teren.", "Pripr.", "Godišnji", "Bolov.", "Praznik", "Sedm.\nodmor", "Napomena"]
class StranicaEvidencija(QWidget):
    def __init__(self, glavni):
        super().__init__(); self.g = glavni; self.model = None
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 20); v.setSpacing(12)
        h = QHBoxLayout(); h.addWidget(Zaglavlje("Evidencija", "Sve šihtarice po periodima. Pregled je samo za čitanje – ispravke se rade u Excelu uz napomenu o korekciji."), 1)
        b = QPushButton("Osvježi"); b.clicked.connect(self.ucitaj); h.addWidget(b, 0, Qt.AlignTop)
        b = QPushButton("Otvori u Excelu"); b.clicked.connect(self.g.otvori_evidenciju); h.addWidget(b, 0, Qt.AlignTop)
        v.addLayout(h)
        sp = QSplitter(); sp.setHandleWidth(10)
        lijevo = QWidget(); lijevo.setMinimumWidth(330); lv = QVBoxLayout(lijevo); lv.setContentsMargins(0, 0, 0, 0)
        self.trazi = QLineEdit(); self.trazi.setPlaceholderText("Traži radnika…"); self.trazi.textChanged.connect(self._filtriraj); lv.addWidget(self.trazi)
        self.stablo = QTreeWidget(); self.stablo.setHeaderLabels(["Period / radnik", "Sati", "⚑"]); self.stablo.headerItem().setToolTip(2, "Broj dana s napomenom")
        hd = self.stablo.header(); hd.setStretchLastSection(False); hd.setSectionResizeMode(0, QHeaderView.Stretch)
        hd.setSectionResizeMode(1, QHeaderView.ResizeToContents); hd.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.stablo.setRootIsDecorated(True); self.stablo.setAlternatingRowColors(True)
        self.stablo.currentItemChanged.connect(self._odabran); lv.addWidget(self.stablo)
        sp.addWidget(lijevo)
        desno = QWidget(); dv = QVBoxLayout(desno); dv.setContentsMargins(0, 0, 0, 0); dv.setSpacing(8)
        self.lbl_sih = lbl("Odaberite radnika lijevo.", "podnaslov", True); dv.addWidget(self.lbl_sih)
        self.t = tabela(KOL_OBRASCA, 15); self.t.setWordWrap(True); self.t.horizontalHeader().setMinimumHeight(42)
        hh = self.t.horizontalHeader()
        for c, w in enumerate([40, 56, 92, 64, 70, 64, 70, 64, 62, 58, 58, 64, 58, 62, 62]): hh.setSectionResizeMode(c, QHeaderView.Fixed); self.t.setColumnWidth(c, w)
        dv.addWidget(self.t)
        self.lbl_skriveno = lbl("", "info"); dv.addWidget(self.lbl_skriveno)
        sp.addWidget(desno); sp.setSizes([300, 960]); v.addWidget(sp, 1)

    def ucitaj(self):
        ev = self.g.evidencija(); self.stablo.clear(); self.t.setRowCount(0)
        if not ev: return
        try: self.model = I.procitaj_model(ev)
        except Exception as e: self.lbl_sih.setText(f"Evidenciju nije moguće pročitati: {e}"); return
        grupe = {}
        for i, p in enumerate(self.model["sihtarice"]):
            k = f"{p['mjesec_naziv']} {p['godina']}"
            if k not in grupe:
                grupe[k] = QTreeWidgetItem([k, "", ""]); f = grupe[k].font(0); f.setBold(True); grupe[k].setFont(0, f)
                self.stablo.insertTopLevelItem(0, grupe[k])
            nap = sum(1 for d in p["dani"] if d["napomena"])
            it = QTreeWidgetItem([f"{p['radnik']['ime']}  ·  {p['ustanova']}", f"{p['ukupno']['ukupno_sati']:.2f}", str(nap) if nap else ""])
            it.setData(0, Qt.UserRole, i); it.setTextAlignment(1, Qt.AlignRight)
            if nap: it.setForeground(2, QColor("#B26A00"))
            grupe[k].addChild(it)
        if self.stablo.topLevelItemCount(): self.stablo.topLevelItem(0).setExpanded(True)
        self.lbl_sih.setText(f"{len(self.model['sihtarice'])} šihtarica u evidenciji. Odaberite radnika lijevo.")

    def _filtriraj(self, t):
        t = t.lower()
        for i in range(self.stablo.topLevelItemCount()):
            g = self.stablo.topLevelItem(i); vid = 0
            for j in range(g.childCount()):
                c = g.child(j); s = t in c.text(0).lower(); c.setHidden(not s); vid += s
            g.setHidden(vid == 0); g.setExpanded(bool(t) or g.isExpanded())

    def _odabran(self, it, _=None):
        if not it or it.data(0, Qt.UserRole) is None: return
        p = self.model["sihtarice"][it.data(0, Qt.UserRole)]; r = p["radnik"]
        self.lbl_sih.setText(f"<span style='font-size:13pt;font-weight:700;color:{TAMNA}'>{r['ime']}</span> &nbsp; {p['ustanova']} · {p['mjesec_naziv']} {p['godina']}"
                             f" · radno mjesto: {r['radno_mjesto'] or '—'} · sedmično: {r['ugovoreno_sedmicno'] or '—'} h · odgovorno lice: {r['odgovorno_lice'] or '—'}"
                             f" · list „{p['list']}”")
        f = lambda x: "" if x in (None, "") else (f"{x:.2f}" if isinstance(x, float) else str(x))
        t = self.t; t.setRowCount(0)
        for d in p["dani"]:
            i = t.rowCount(); t.insertRow(i)
            datum = dt.date.fromisoformat(d["datum"]); vik = datum.weekday() >= 5
            vals = (d["rb"], datum.strftime("%d.%m."), d["dan"], d["pocetak_rada"], d["zavrsetak_rada"], d["pocetak_pauze"], d["zavrsetak_pauze"],
                    f(d["ukupno_sati"]), f(d["prekovremeni"]), f(d["terenski"]), f(d["pripravnost"]), f(d["godisnji"]), f(d["bolovanje"]),
                    f(d["praznik"]), f(d["sedmicni_odmor"]), d["napomena"])
            for c, v in enumerate(vals):
                s = stavka(v)
                if c < 15: s.setTextAlignment(Qt.AlignCenter)
                if vik and c < 3: s.setForeground(QColor("#8A99A6"))
                if c == 15 and v: s.setBackground(QColor("#FFE7A3" if v in E.MARKERI or "nije" in v or "Nedostaje" in v else "#FFF6D5")); s.setToolTip(v)
                t.setItem(i, c, s)
        i = t.rowCount(); t.insertRow(i); u = p["ukupno"]
        for c, v in ((0, "UKUPNO"), (7, f(u["ukupno_sati"])), (8, f(u["prekovremeni"])), (9, f(u["terenski"])), (10, f(u["pripravnost"])),
                     (11, f(u["godisnji"])), (12, f(u["bolovanje"])), (13, f(u["praznik"])), (14, f(u["sedmicni_odmor"]))):
            s = stavka(v, boja="#FFF6D5"); fn = s.font(); fn.setBold(True); s.setFont(fn); s.setTextAlignment(Qt.AlignCenter); t.setItem(i, c, s)
        for c in range(16):
            if not t.item(i, c): t.setItem(i, c, stavka("", boja="#FFF6D5"))
        # kolone bez ijednog podatka se sklanjaju radi preglednosti (osim osnovnih)
        skriveno = []
        for c in range(5, 15):
            prazno = all(not (t.item(r, c) and t.item(r, c).text()) for r in range(t.rowCount() - 1))
            t.setColumnHidden(c, prazno and c != 7)
            if prazno and c != 7: skriveno.append(KOL_OBRASCA[c].replace("\n", " "))
        self.lbl_skriveno.setText(("Prazne kolone sklonjene iz pregleda: " + ", ".join(skriveno) + ". U evidenciji i izvozima obrazac je uvijek kompletan.") if skriveno else "")
        t.resizeRowsToContents()

# ------------------------------------------------------------------ stranica: IZVOZ
class StranicaIzvoz(QWidget):
    def __init__(self, glavni):
        super().__init__(); self.g = glavni; self.radnik = None
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 20); v.setSpacing(14)
        v.addWidget(Zaglavlje("Izvoz podataka", "Izvoz čita isključivo evidenciju – podaci u izvozu su identični šihtaricama."))
        k = kartica(); g = QGridLayout(k); g.setContentsMargins(20, 18, 20, 18); g.setHorizontalSpacing(16); g.setVerticalSpacing(12)
        g.addWidget(lbl("<b>Obuhvat</b>"), 0, 0)
        self.cb_per = QComboBox(); self.cb_per.setMinimumWidth(200); g.addWidget(lbl("Period"), 1, 0); g.addWidget(self.cb_per, 1, 1)
        self.cb_inst = QComboBox(); self.cb_inst.addItems(["Sve ustanove", "IDSS", "IMH"]); g.addWidget(lbl("Ustanova"), 2, 0); g.addWidget(self.cb_inst, 2, 1)
        g.addWidget(lbl("<b>Formati</b>"), 3, 0)
        self.chk = {}
        for i, (k_, n, opis) in enumerate((("xlsx", "Excel (.xlsx)", "Propisani obrazac s formulama, mjesečni pregled, registar, dnevnik i upozorenja"),
                                           ("pdf", "PDF (.pdf)", "Spremno za štampu i potpis: pregled + jedna stranica po šihtarici (A4 položeno)"),
                                           ("json", "JSON (.json)", "Strukturirani podaci za druge sisteme (obračun plata, arhiva)"))):
            c = QCheckBox(n); c.setChecked(True); self.chk[k_] = c
            g.addWidget(c, 4 + i, 0); g.addWidget(lbl(opis, "info"), 4 + i, 1, 1, 2)
        g.addWidget(lbl("<b>Odredište</b>"), 7, 0)
        self.le_folder = QLineEdit(str(Path.home() / "Documents")); g.addWidget(self.le_folder, 8, 0, 1, 2)
        b = QPushButton("Promijeni…"); b.clicked.connect(self._folder); g.addWidget(b, 8, 2)
        g.setColumnStretch(1, 1); v.addWidget(k)
        self.napredak = QProgressBar(); self.napredak.setRange(0, 0); self.napredak.hide(); v.addWidget(self.napredak)
        self.obavijest = Obavijest(); v.addWidget(self.obavijest)
        h = QHBoxLayout(); h.addStretch()
        self.btn_folder = QPushButton("Otvori folder"); self.btn_folder.hide(); self.btn_folder.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(self.le_folder.text()))); h.addWidget(self.btn_folder)
        self.btn = QPushButton("Izvezi"); self.btn.setObjectName("primarno"); self.btn.clicked.connect(self.izvezi); h.addWidget(self.btn)
        v.addLayout(h); v.addStretch()

    def osvjezi(self):
        ev = self.g.evidencija(); self.cb_per.clear(); self.cb_per.addItem("Svi periodi", None)
        if ev:
            for p in reversed(I.dostupni_periodi(ev)):
                g, m = p.split("-"); self.cb_per.addItem(f"{E.MJESECI[int(m)-1]} {g}", p)

    def _folder(self):
        p = QFileDialog.getExistingDirectory(self, "Odredište izvoza", self.le_folder.text())
        if p: self.le_folder.setText(p)

    def izvezi(self):
        ev = self.g.evidencija()
        if not ev: self.obavijest.pokazi("Nema otvorene evidencije.", "savjet"); return
        formati = [k for k, c in self.chk.items() if c.isChecked()]
        if not formati: self.obavijest.pokazi("Odaberite barem jedan format.", "savjet"); return
        per = self.cb_per.currentData(); inst = None if self.cb_inst.currentIndex() == 0 else self.cb_inst.currentText()
        folder = Path(self.le_folder.text()); folder.mkdir(parents=True, exist_ok=True)
        osnova = f"Sihtarica_{per or 'svi-periodi'}_{inst or 'IDSS-IMH'}_{dt.datetime.now():%Y%m%d_%H%M}"
        def posao():
            out = []
            for f in formati:
                cilj = folder / f"{osnova}.{f}"
                {"xlsx": I.izvezi_xlsx, "pdf": I.izvezi_pdf, "json": I.izvezi_json}[f](ev, cilj, per, inst); out.append(cilj.name)
            return out
        self.btn.setEnabled(False); self.napredak.show(); self.obavijest.hide()
        self.radnik = Radnik(posao)
        def kraj(): self.btn.setEnabled(True); self.napredak.hide()
        self.radnik.gotovo.connect(lambda r: (kraj(), self.obavijest.pokazi("✔  Izvezeno u " + str(folder) + ":\n" + "\n".join(r), "uspjeh"), self.btn_folder.show()))
        self.radnik.greska.connect(lambda t: (kraj(), self.obavijest.pokazi(t, "greska")))
        self.radnik.start()

# ------------------------------------------------------------------ stranica: POSTAVKE
class StranicaPostavke(QWidget):
    def __init__(self, glavni):
        super().__init__(); self.g = glavni
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 20); v.setSpacing(14)
        v.addWidget(Zaglavlje("Postavke tumačenja", "Kako se prijave s uređaja prevode u šihtaricu. Postavke korištene pri svakom upisu bilježe se u „Dnevnik uvoza”."))
        k = kartica(); kv = QVBoxLayout(k); kv.setContentsMargins(20, 18, 20, 18); kv.setSpacing(10)
        kv.addWidget(lbl("<b>Logika prijave</b>"))
        kv.addWidget(lbl("Radnik se na ulazu u školu/vrtić i pri izlasku prijavljuje otiskom prsta ili šifrom. Uređaj bilježi samo trenutak prijave, "
                         "bez oznake ulaz/izlaz. Zato je prva prijava u danu <b>dolazak</b>, zadnja <b>odlazak</b>, a parovi prijava između njih su "
                         "<b>međuizlasci</b>. Kad smisao nije jednoznačan (jedna ili neparan broj prijava), vremena se ne raspoređuju nego doslovno "
                         "navode u napomeni, a dan se označava za provjeru. Ako izvoz sadrži način prijave, prijava šifrom se navodi u napomeni.", "info", True))
        h = QHBoxLayout(); h.addWidget(lbl("Dvije prijave unutar"))
        self.sp = QSpinBox(); self.sp.setRange(0, 15); self.sp.setSuffix(" min"); h.addWidget(self.sp)
        h.addWidget(lbl("smatraju se jednom (radnik je ponovio prst ili šifru).")); h.addStretch(); kv.addLayout(h)
        kv.addWidget(lbl("<b>Jedan međuizlazak u danu (4 prijave)</b>"))
        self.rb_p = QRadioButton("Upisati u kolone „Početak/Završetak pauze” – obrazac ga odbija od dnevnih sati")
        self.rb_n = QRadioButton("Samo navesti u napomeni – ne odbija se od dnevnih sati")
        grp = QButtonGroup(self); grp.addButton(self.rb_p); grp.addButton(self.rb_n); kv.addWidget(self.rb_p); kv.addWidget(self.rb_n)
        kv.addWidget(lbl("Dva ili više međuizlazaka se nikad ne odbijaju automatski – svi se navode u napomeni i označavaju za provjeru.", "info", True))
        kv.addWidget(lbl("<b>Odgovorno lice (ako nije upisano kod radnika)</b>"))
        gr = QGridLayout(); self.le = {}
        for i, inst in enumerate(("IDSS", "IMH")):
            gr.addWidget(lbl(inst), i, 0); self.le[inst] = QLineEdit(); gr.addWidget(self.le[inst], i, 1)
        kv.addLayout(gr); v.addWidget(k)
        self.obavijest = Obavijest(); v.addWidget(self.obavijest)
        h = QHBoxLayout(); h.addStretch(); b = QPushButton("Sačuvaj postavke"); b.setObjectName("primarno"); b.clicked.connect(self.sacuvaj); h.addWidget(b)
        v.addLayout(h); v.addStretch(); self.ucitaj()

    def ucitaj(self):
        t = self.g.post["tumacenje"]
        self.sp.setValue(int(t.get("prag_ponovljenog_otiska_min", 2)))
        (self.rb_p if t.get("medjuizlazak", "pauza") == "pauza" else self.rb_n).setChecked(True)
        for inst, le in self.le.items(): le.setText(t.get("odgovorno_lice", {}).get(inst, ""))

    def sacuvaj(self):
        self.g.post["tumacenje"] = dict(prag_ponovljenog_otiska_min=self.sp.value(), medjuizlazak="pauza" if self.rb_p.isChecked() else "napomena",
                                        odgovorno_lice={k: v.text().strip() for k, v in self.le.items()})
        sacuvaj_postavke(self.g.post); self.obavijest.pokazi("✔  Postavke su sačuvane i primjenjuju se od sljedećeg uvoza.", "uspjeh")

# ------------------------------------------------------------------ glavni prozor
class Glavni(QMainWindow):
    def __init__(self):
        super().__init__(); self.post = ucitaj_postavke()
        self.setWindowTitle(APP); self.resize(1360, 860); self.setMinimumSize(1100, 700)
        c = QWidget(); h = QHBoxLayout(c); h.setContentsMargins(0, 0, 0, 0); h.setSpacing(0); self.setCentralWidget(c)
        meni = QFrame(); meni.setObjectName("meni"); meni.setFixedWidth(232); mv = QVBoxLayout(meni); mv.setContentsMargins(0, 22, 0, 16); mv.setSpacing(2)
        lg = QVBoxLayout(); lg.setContentsMargins(22, 0, 22, 18); lg.addWidget(lbl("Šihtarica", "logo")); lg.addWidget(lbl("IDSS · IMH  ·  FA01H", "podlogo")); mv.addLayout(lg)
        self.stog = QStackedWidget(); self.stog.setObjectName("sadrzaj")
        self.s_uvoz, self.s_ev, self.s_izv, self.s_post = StranicaUvoz(self), StranicaEvidencija(self), StranicaIzvoz(self), StranicaPostavke(self)
        self.nav = []
        for i, (ikona, n, s) in enumerate((("⇪", "Uvoz s uređaja", self.s_uvoz), ("▤", "Evidencija", self.s_ev),
                                           ("⇩", "Izvoz podataka", self.s_izv), ("⚙", "Postavke", self.s_post))):
            self.stog.addWidget(s)
            b = QPushButton(f"  {ikona}   {n}"); b.setObjectName("nav"); b.setCheckable(True); b.setCursor(Qt.PointingHandCursor); b.setFocusPolicy(Qt.NoFocus)
            b.clicked.connect(lambda _=False, i=i: self.idi(i)); mv.addWidget(b); self.nav.append(b)
        mv.addStretch()
        fk = QFrame(); fk.setObjectName("fajlKartica"); fv = QVBoxLayout(fk); fv.setContentsMargins(14, 12, 14, 12); fv.setSpacing(4)
        fv.addWidget(lbl("EVIDENCIJSKI FAJL", "fajlNaslov")); self.lbl_fajl = lbl("—", "fajlIme", True); fv.addWidget(self.lbl_fajl)
        hb = QHBoxLayout(); hb.setSpacing(6)
        for t, fn in (("Novi…", self.kreiraj), ("Otvori…", self.odaberi)):
            b = QPushButton(t); b.setStyleSheet("QPushButton{background:#1C4A78;color:white;border:none;padding:5px 10px;border-radius:5px;font-size:9pt}"
                                                "QPushButton:hover{background:#245C93}"); b.setFocusPolicy(Qt.NoFocus); b.clicked.connect(fn); hb.addWidget(b)
        fv.addLayout(hb)
        w = QWidget(); wl = QVBoxLayout(w); wl.setContentsMargins(14, 0, 14, 0); wl.addWidget(fk); mv.addWidget(w)
        mv.addWidget(lbl(f"   v{E.VERZIJA}", "podlogo"))
        h.addWidget(meni); h.addWidget(self.stog, 1)
        self.s_uvoz.upisano.connect(self.osvjezi_sve)
        self._osvjezi_fajl(); self.idi(0)

    def idi(self, i):
        for j, b in enumerate(self.nav): b.setChecked(i == j)
        self.stog.setCurrentIndex(i)
        if i == 1: self.s_ev.ucitaj()
        if i == 2: self.s_izv.osvjezi()

    def evidencija(self):
        p = self.post.get("evidencija"); return Path(p) if p and Path(p).exists() else None

    def _osvjezi_fajl(self):
        ev = self.evidencija()
        self.lbl_fajl.setText(ev.name if ev else "Nije odabran"); self.lbl_fajl.setToolTip(str(ev) if ev else "")
        if not ev: self.s_uvoz.obavijest.pokazi("Dobrodošli. Za početak kliknite „Novi…” (lijevo dolje) da kreirate evidencijski fajl iz propisanog obrasca, "
                                                "ili „Otvori…” ako evidencija već postoji.", "savjet")
        else: self.s_uvoz.obavijest.hide()

    def osvjezi_sve(self): self.s_izv.osvjezi()

    def kreiraj(self):
        p, _ = QFileDialog.getSaveFileName(self, "Novi evidencijski fajl", str(Path.home() / "Documents" / "Sihtarica_IDSS_IMH_evidencija.xlsx"), "Excel (*.xlsx)")
        if not p: return
        if Path(p).exists():
            QMessageBox.warning(self, APP, "Fajl s tim imenom već postoji. Postojeća evidencija se nikad ne prepisuje – otvorite je preko „Otvori…”."); return
        try: E.kreiraj_evidenciju(resurs("sablon_Sihtarica_IDSS_IMH.xlsx"), p)
        except Exception as e: QMessageBox.critical(self, APP, str(e)); return
        self.post["evidencija"] = p; sacuvaj_postavke(self.post); self._osvjezi_fajl()
        self.s_uvoz.obavijest.pokazi("✔  Evidencijski fajl je kreiran iz propisanog obrasca. Sada prevucite izvoz s uređaja.", "uspjeh")

    def odaberi(self):
        p, _ = QFileDialog.getOpenFileName(self, "Otvori evidenciju", str(Path.home() / "Documents"), "Excel (*.xlsx)")
        if not p: return
        try: E.provjeri_evidenciju(p)
        except Exception as e: QMessageBox.critical(self, APP, str(e)); return
        self.post["evidencija"] = p; sacuvaj_postavke(self.post); self._osvjezi_fajl(); self.idi(self.stog.currentIndex())

    def otvori_evidenciju(self):
        ev = self.evidencija()
        if ev: QDesktopServices.openUrl(QUrl.fromLocalFile(str(ev)))

def main():
    if "--provjera" in sys.argv:
        import provjera
        i = sys.argv.index("--provjera"); cilj = sys.argv[i + 1] if len(sys.argv) > i + 1 else str(Path.home() / "Desktop")
        Path(cilj).mkdir(parents=True, exist_ok=True)
        ok, _ = provjera.pokreni(cilj); sys.exit(0 if ok else 1)
    if "--samotest" in sys.argv:
        import samotest; sys.exit(samotest.pokreni(sys.argv[sys.argv.index("--samotest") + 1]))
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv); app.setStyleSheet(QSS); app.setApplicationName(APP)
    w = Glavni(); w.show(); sys.exit(app.exec())

if __name__ == "__main__":
    main()
