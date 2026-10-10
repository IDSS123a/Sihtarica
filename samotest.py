# -*- coding: utf-8 -*-
"""Dijagnostički test instalacije: Sihtarica_IDSS_IMH.exe --samotest <folder s izvozima>
Prolazi kroz cijeli tok u pravom prozoru aplikacije i snima ekrane i izvještaj u <folder>\\rezultat."""
import os, sys, time, traceback, ctypes, datetime as dt
from pathlib import Path

def pokreni(folder):
    folder = Path(folder); out = folder / "rezultat"; out.mkdir(exist_ok=True)
    os.environ["APPDATA"] = str(out / "appdata")
    log = []
    def zapis(ok, t):
        log.append(("OK   " if ok else "GREŠKA ") + t)
        try: print(log[-1], flush=True)
        except Exception: pass
    from PySide6.QtWidgets import QApplication, QMessageBox
    import app as A, engine as E, izvoz as I
    qa = QApplication(sys.argv); qa.setStyleSheet(A.QSS)
    QMessageBox.question = lambda *a, **k: QMessageBox.Yes
    w = A.Glavni(); w.resize(1360, 860); w.show()
    def cekaj(s):
        while getattr(s, "radnik", None) and s.radnik.isRunning(): qa.processEvents(); time.sleep(0.05)
        for _ in range(30): qa.processEvents()
    def snimi(n): qa.processEvents(); w.grab().save(str(out / f"{n}.png"))
    try:
        zapis(True, f"Sistem: {sys.platform}, Python {sys.version.split()[0]}, zamrznuto (exe): {getattr(sys, 'frozen', False)}")
        ev = out / "Evidencija_test.xlsx"
        if ev.exists(): ev.unlink()
        E.kreiraj_evidenciju(A.resurs("sablon_Sihtarica_IDSS_IMH.xlsx"), ev); w.post["evidencija"] = str(ev); w._osvjezi_fajl()
        zapis(ev.exists(), f"Evidencija kreirana iz ugrađenog obrasca: {ev}")
        snimi("01_pocetak")
        def uvoz(ime, ocekuj_gresku=False, kategorija=None):
            u = w.s_uvoz; u.analiziraj(str(folder / ime)); cekaj(u)
            if ocekuj_gresku:
                zapis(u.obavijest.isVisible() and u.obavijest.objectName() == "greska", f"{ime}: odbijen s porukom – {u.obavijest.text()[:90]}"); return
            if u.plan is None: zapis(False, f"{ime}: analiza nije uspjela – {u.obavijest.text()[:200]}"); return
            for r in range(u.t_radnici.rowCount()):
                if not u.t_radnici.item(r, 4).text(): u.t_radnici.item(r, 4).setText("40")
            if kategorija and u.t_kat.rowCount():
                u.tabovi.setCurrentIndex(1); u.t_kat.selectRow(0); u.cb_kat.setCurrentText(kategorija); u._primijeni_kategoriju()
            snimi(f"02_analiza_{Path(ime).stem}")
            u.upisi(); cekaj(u)
            zapis(u.obavijest.objectName() == "uspjeh", f"{ime}: {u.obavijest.text().splitlines()[0][:150]}")
        uvoz("All_Report_mockup_October_2026.xlsx", kategorija="Neplaćeno odsustvo")
        u = w.s_uvoz; u.analiziraj(str(folder / "All_Report_mockup_October_2026.xlsx")); cekaj(u)
        zapis(not u.btn_upis.isEnabled() and "već upisan" in u.lbl_akcija.text(), "Ponovni uvoz istog fajla: upis onemogućen")
        uvoz("Nov_1-15.xlsx"); uvoz("Nov_full.xlsx"); uvoz("Dec_grid.xlsx")
        uvoz("All_Report.xls", ocekuj_gresku=True); snimi("03_prazni_xls")
        # zaključan fajl (kao kad je otvoren u Excelu): CreateFileW bez dijeljenja
        if sys.platform == "win32":
            k32 = ctypes.windll.kernel32; k32.CreateFileW.restype = ctypes.c_void_p
            h = k32.CreateFileW(str(ev), 0x80000000 | 0x40000000, 0, None, 3, 0x80, None)
            try:
                try:
                    p = E.napravi_plan(E.citaj_izvor(str(folder / "Dec_grid.xlsx")), ev); p["vec_uvezen"] = None
                    E.upisi(p, ev); zapis(False, "Zaključan fajl: upis je prošao (nije smio)")
                except E.GreskaUpisa as e: zapis("Excel" in str(e) or "zaključan" in str(e), f"Zaključan fajl (simuliran Excel): {str(e)[:90]}")
            finally: k32.CloseHandle(ctypes.c_void_p(h))
        w.idi(1); qa.processEvents(); st = w.s_ev.stablo
        zapis(st.topLevelItemCount() == 3, f"Evidencija: {st.topLevelItemCount()} perioda, {len(w.s_ev.model['sihtarice'])} šihtarica")
        g = st.topLevelItem(st.topLevelItemCount() - 1); g.setExpanded(True)
        for j in range(g.childCount()):
            if "Nermin" in g.child(j).text(0): st.setCurrentItem(g.child(j))
        snimi("04_evidencija")
        w.idi(2); w.s_izv.le_folder.setText(str(out / "izvoz")); w.s_izv.cb_per.setCurrentIndex(1); w.s_izv.izvezi(); cekaj(w.s_izv)
        f = sorted((out / "izvoz").glob("*")); zapis(len(f) == 3, "Izvoz: " + ", ".join(x.name for x in f)); snimi("05_izvoz")
        w.idi(3); snimi("06_postavke")
        import openpyxl
        wb = openpyxl.load_workbook(ev); s = wb["2026-10 IMH MK019"]
        zapis(s["L28"].value == 8 and s["D7"].value == dt.time(8, 0), f"Provjera ćelija: MK019 22.10. godišnji={s['L28'].value} h, 01.10. dolazak={s['D7'].value}")
        zapis(len(list((ev.parent / "Sigurnosne_kopije").glob("*.xlsx"))) >= 4, "Sigurnosne kopije postoje")
    except Exception:
        zapis(False, "Izuzetak:\n" + traceback.format_exc())
    (out / "izvjestaj.txt").write_text("\n".join(log), encoding="utf-8")
    w.close(); return 0 if all(l.startswith("OK") for l in log) else 1
