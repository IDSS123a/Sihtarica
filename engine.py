# -*- coding: utf-8 -*-
"""
ŠIHTARICA IDSS/IMH – jezgro prevođenja (v2)

Izvor (2): izvoz uređaja FA01H "All Report" (.xls / .xlsx)
Cilj  (1): jedan trajni evidencijski fajl zasnovan na propisanom obrascu Sihtarica_IDSS_IMH.xlsx

NAČELA
 • Prevodi se SVE što postoji u izvoru; što nedostaje ostaje prazno i jasno označeno – ništa se ne izmišlja.
 • Logika prijave: radnik se na ulazu/izlazu prijavljuje otiskom prsta ili šifrom. Uređaj bilježi samo
   trenutke prijave (bez oznake ulaz/izlaz). Zato: 1. prijava = dolazak, zadnja = odlazak, a parovi između
   njih su međuizlasci. Kad je smisao neodređen (jedna ili neparan broj prijava), vrijeme se NE raspoređuje
   u kolonu dolaska/odlaska nego se doslovno navodi u napomeni.
 • Istorija je nepromjenjiva: svaka ranije popunjena ćelija ostaje identična. Dozvoljeno je samo popuniti
   prazne dane (dopuna). Dokaz se provodi poređenjem svake ćelije prije i poslije; ako ne prođe, upis se poništava.
"""
import collections, datetime as dt, getpass, hashlib, os, re, shutil, tempfile
from copy import copy, deepcopy
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.utils import get_column_letter

VERZIJA = "2.1.0"
MJESECI = ["Januar","Februar","Mart","April","Maj","Juni","Juli","August","Septembar","Oktobar","Novembar","Decembar"]
DANI = ["Ponedjeljak","Utorak","Srijeda","Četvrtak","Petak","Subota","Nedjelja"]
OBRAZAC = {"IDSS": "IDSS - šihtarica", "IMH": "IMH - šihtarica"}
L_UPUTSTVO, L_PREGLED, L_REGISTAR, L_DNEVNIK, L_UPOZORENJA = (
    "Uputstvo", "Mjesečni pregled", "Registar radnika", "Dnevnik uvoza", "Upozorenja")
PRVI_RED = 7
NAZIV_LISTA = re.compile(r"^(\d{4})-(\d{2}) (IDSS|IMH) (.+)$")
KOLONE_PODATAKA = list("DEFGIJKLMNO")
MARKER_NEMA = "Nema podataka u izvozu uređaja"
MARKER_VAN = "Dan nije obuhvaćen izvozom uređaja"
MARKERI = {MARKER_NEMA, MARKER_VAN}

KATEGORIJE = collections.OrderedDict([
    ("Bez kategorije", None), ("Godišnji odmor", "L"), ("Bolovanje", "M"), ("Praznik / neradni dan", "N"),
    ("Sedmični odmor", "O"), ("Terenski rad / službeni put", "J"), ("Plaćeno odsustvo", None),
    ("Neplaćeno odsustvo", None), ("Opravdano odsustvo – drugo", None), ("Neopravdano odsustvo", None)])

PODRAZUMIJEVANE_POSTAVKE = dict(
    prag_ponovljenog_otiska_min=2,          # dvije prijave unutar N minuta = jedna (radnik ponovio prst/šifru)
    medjuizlazak="pauza",                   # "pauza" -> F/G (odbija se po formuli obrasca); "napomena" -> samo P
    odgovorno_lice={"IDSS": "", "IMH": ""},
)
BOJA = dict(plava="035EA1", svijetloplava="08ABE6", traka="F7FAFC", zuta="FFF6D5", amber="FFE7A3", tekst="263238")

class GreskaFormata(Exception): pass
class GreskaUpisa(Exception): pass

# ======================================================================= pomoćne
def _datum(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    if isinstance(v, str):
        for f in ("%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y."):
            try: return dt.datetime.strptime(v.strip(), f).date()
            except ValueError: pass
    raise GreskaFormata(f"neprepoznat datum {v!r}")

def _vrijeme(v):
    if v in (None, ""): return None
    if isinstance(v, dt.datetime): return dt.time(v.hour, v.minute)
    if isinstance(v, dt.time): return dt.time(v.hour, v.minute)
    if isinstance(v, (int, float)) and 0 <= v < 1:
        m = int(Decimal(str(v * 1440)).quantize(Decimal(1), ROUND_HALF_UP)); return dt.time(m // 60 % 24, m % 60)
    if isinstance(v, str) and re.fullmatch(r"\s*\d{1,2}:\d{2}(:\d{2})?\s*", v):
        h, m = v.strip().split(":")[:2]
        if int(h) < 24 and int(m) < 60: return dt.time(int(h), int(m))
    raise GreskaFormata(f"neprepoznato vrijeme {v!r}")

def _min(t): return t.hour * 60 + t.minute
def _t(m): return dt.time(m // 60, m % 60)
def _hm(t): return t.strftime("%H:%M")
def prijava_txt(n): return f"{n} prijava" if n % 10 == 1 and n % 100 != 11 else (f"{n} prijave" if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14) else f"{n} prijava")
def _id(v): return re.sub(r"\.0$", "", str(v).strip())
def zaokruzi(x, n=2): return float(Decimal(str(x)).quantize(Decimal(1).scaleb(-n), ROUND_HALF_UP))
def zadnji_dan(g, m): return (dt.date(g + m // 12, m % 12 + 1, 1) - dt.timedelta(days=1)).day

def sati_obrasca(d, e, f, g):
    """Identično formuli obrasca: ROUND(MOD(E-D,1)*24 - IF(F,G; MOD(G-F,1)*24; 0); 2)."""
    if not (d and e): return None
    h = ((_min(e) - _min(d)) % 1440) / 60
    if f and g: h -= ((_min(g) - _min(f)) % 1440) / 60
    return zaokruzi(h)

def sigurna_kopija(izvor, cilj):
    """Kopija sadržaja + vremenskih oznaka (kao shutil.copy2), bez Windows funkcije CopyFile2."""
    shutil.copyfile(izvor, cilj); shutil.copystat(izvor, cilj)
    if Path(izvor).stat().st_size != Path(cilj).stat().st_size: raise GreskaUpisa("Sigurnosna kopija nije potpuna – upis prekinut.")

def sha256_fajla(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 16), b""): h.update(b)
    return h.hexdigest()

_NEDOZVOLJENO = re.compile(r"[\[\]:*?/\\]")
def naziv_lista(g, m, inst, rid):
    cist = _NEDOZVOLJENO.sub("-", str(rid))
    return f"{g}-{m:02d} {inst} {cist}"[:31]

# ======================================================================= čitanje izvora (2)
def _otvori(path):
    p = str(path)
    if p.lower().endswith(".xls"):
        import xlrd
        b = xlrd.open_workbook(p)
        out = {}
        for s in b.sheets():
            rows = []
            for r in range(s.nrows):
                row = []
                for c in range(s.ncols):
                    cell = s.cell(r, c)
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        t = xlrd.xldate_as_tuple(cell.value, b.datemode)
                        row.append(dt.time(t[3], t[4]) if t[:3] == (0, 0, 0) else dt.datetime(*t))
                    else: row.append(cell.value)
                rows.append(row)
            out[s.name] = rows
        return out
    if not p.lower().endswith((".xlsx", ".xlsm")):
        raise GreskaFormata("Podržani su samo Excel izvozi uređaja (.xls, .xlsx).")
    wb = openpyxl.load_workbook(p, data_only=True)
    return {ws.title: [[c.value for c in row] for row in ws.iter_rows()] for ws in wb}

def _period(rows):
    for r in rows[:5]:
        for v in r:
            if isinstance(v, str):
                m = re.search(r"(\d{2})/(\d{2})/(\d{4})\s*~\s*(\d{2})/(\d{2})/(\d{4})", v)
                if m:
                    d1, m1, y1, d2, m2, y2 = map(int, m.groups())
                    od, do = dt.date(y1, m1, d1), dt.date(y2, m2, d2)
                    if do < od: raise GreskaFormata("Period izvoza je neispravan (kraj prije početka).")
                    return od, do
    raise GreskaFormata("U listu 'Attendance Logs' nije pronađen period izvoza (dd/mm/gggg ~ dd/mm/gggg).")

_METODA = [(re.compile(r"finger|fp|otisak|prst", re.I), "otisak prsta"),
           (re.compile(r"pass|pw|pin|šifra|sifra|lozink", re.I), "šifra"),
           (re.compile(r"card|kartic|rfid", re.I), "kartica"), (re.compile(r"face|lice", re.I), "lice")]
def _metoda(v):
    if v in (None, ""): return None
    for rx, n in _METODA:
        if rx.search(str(v)): return n
    return str(v).strip()

def _citaj_tabelu(rows, hdr_i, od, do):
    h = {str(n).strip().lower(): i for i, n in enumerate(rows[hdr_i]) if n not in (None, "")}
    def kol(*imena):
        for n in imena:
            if n.lower() in h: return h[n.lower()]
        return None
    K = dict(datum=kol("Date"), id=kol("Employee ID", "ID"), ime=kol("Employee name", "Name"),
             odjel=kol("Department", "Dept.", "Dept"), ulaz=kol("Clock-in", "Clock in", "In"),
             izlaz=kol("Clock-out", "Clock out", "Out"), status=kol("Status"), ot=kol("Overtime (hours)", "Overtime"),
             sati=kol("Worked hours"), metoda=kol("Verify", "Verify mode", "Verification", "Verify type", "Način prijave"))
    nema = [n for n, k in (("Date", "datum"), ("Employee ID", "id"), ("Clock-in", "ulaz"), ("Clock-out", "izlaz")) if K[k] is None]
    if nema: raise GreskaFormata("U listu 'Attendance Logs' nedostaju kolone: " + ", ".join(nema))
    zapisi = []
    for n, r in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
        if not any(v not in (None, "") for v in r): continue
        g = lambda k: (r[K[k]] if K[k] is not None and K[k] < len(r) else None)
        try:
            ot = g("ot")
            ot = ot.total_seconds() / 3600 if isinstance(ot, dt.timedelta) else float(ot or 0)
            z = dict(red=n, datum=_datum(g("datum")), id=_id(g("id")), ime=str(g("ime") or "").strip(),
                     odjel=str(g("odjel") or "").strip(), ulaz=_vrijeme(g("ulaz")), izlaz=_vrijeme(g("izlaz")),
                     otisci=None, ponovljeni=0, status=str(g("status") or "").strip(), ot=zaokruzi(ot),
                     sati_uredjaj=g("sati"), metoda=_metoda(g("metoda")))
        except (GreskaFormata, ValueError, TypeError) as e:
            raise GreskaFormata(f"Attendance Logs, red {n}: {e}")
        if z["id"] in ("", "None"): raise GreskaFormata(f"Attendance Logs, red {n}: nedostaje ID radnika.")
        if not od <= z["datum"] <= do:
            raise GreskaFormata(f"Attendance Logs, red {n}: datum {z['datum']:%d.%m.%Y} je izvan perioda izvoza.")
        zapisi.append(z)
    return zapisi

def _citaj_mrezu(rows, dan_red, od, do, prag):
    kol_dan = {c: int(v) for c, v in enumerate(rows[dan_red]) if isinstance(v, (int, float)) and float(v).is_integer() and 1 <= v <= 31}
    dani = [od + dt.timedelta(days=i) for i in range((do - od).days + 1)]
    if len(kol_dan) != len(dani):
        raise GreskaFormata(f"Broj kolona dana ({len(kol_dan)}) ne odgovara periodu izvoza ({len(dani)} dana).")
    kol_datum = {c: dani[i] for i, c in enumerate(sorted(kol_dan))}
    for c, d in kol_datum.items():
        if kol_dan[c] != d.day: raise GreskaFormata("Redoslijed kolona dana ne odgovara periodu izvoza.")
    def polje(row, oznaka):
        for i, v in enumerate(row):
            if isinstance(v, str) and re.match(rf"^\s*{oznaka}\s*[:：]", v, re.I):
                ost = re.sub(rf"^\s*{oznaka}\s*[:：]\s*", "", v, flags=re.I).strip()
                if ost: return ost
                for w in row[i + 1:]:
                    if w not in (None, ""): return _id(w)
        return None
    zapisi, i = [], dan_red + 1
    while i < len(rows):
        rid = polje(rows[i], r"(ID|No\.?|User ID)")
        if rid is None: i += 1; continue
        ime, odjel = polje(rows[i], "Name") or "", polje(rows[i], r"Dept\.?") or ""
        if i + 1 >= len(rows): raise GreskaFormata(f"Attendance Logs: radnik {rid} nema red s prijavama.")
        red = rows[i + 1]
        for c, datum in kol_datum.items():
            v = red[c] if c < len(red) else None
            tekst = _hm(v) if isinstance(v, (dt.time, dt.datetime)) else str(v or "")
            sirovo = sorted(_min(_vrijeme(x)) for x in re.findall(r"\b\d{1,2}:\d{2}\b", tekst))
            otisci = []
            for m in sirovo:
                if otisci and m - otisci[-1] <= prag: continue
                otisci.append(m)
            zapisi.append(dict(red=i + 2, datum=datum, id=rid, ime=ime, odjel=odjel, ulaz=None, izlaz=None,
                               otisci=otisci, ponovljeni=len(sirovo) - len(otisci), status="", ot=0.0,
                               sati_uredjaj=None, metoda=None))
        i += 2
    return zapisi

def citaj_izvor(path, postavke=None):
    post = {**PODRAZUMIJEVANE_POSTAVKE, **(postavke or {})}
    L = _otvori(path)
    if "Attendance Logs" not in L:
        raise GreskaFormata("Fajl ne sadrži list 'Attendance Logs' – ovo nije 'All Report' izvoz uređaja.")
    rows = L["Attendance Logs"]
    od, do = _period(rows)
    if not any(v not in (None, "") for r in rows[2:] for v in r):
        raise GreskaFormata(f"Izvoz za period {od:%d.%m.%Y}–{do:%d.%m.%Y} ne sadrži nijednu prijavu radnika "
                            "(list 'Attendance Logs' je prazan). Na uređaju provjerite period i izvezite ponovo.")
    hdr = next((i for i, r in enumerate(rows[:15]) if any(str(v).strip().lower() in ("clock-in", "clock in") for v in r if v)), None)
    if hdr is not None:
        zapisi, fmt = _citaj_tabelu(rows, hdr, od, do), "tabelarni (ulaz/izlaz)"
    else:
        n = (do - od).days + 1
        dan = next((i for i, r in enumerate(rows[:10]) if sum(1 for v in r if isinstance(v, (int, float)) and 1 <= v <= 31) >= n), None)
        if dan is None: raise GreskaFormata("List 'Attendance Logs' nije u prepoznatom formatu. Ništa nije prevedeno.")
        zapisi, fmt = _citaj_mrezu(rows, dan, od, do, post["prag_ponovljenog_otiska_min"]), "mrežni (prijave po danima)"
    if not zapisi: raise GreskaFormata("U izvozu nije pronađen nijedan radnik.")
    vidjeno, imena = set(), collections.defaultdict(set)
    for z in zapisi:
        k = (z["id"], z["datum"])
        if k in vidjeno: raise GreskaFormata(f"Izvoz sadrži dva zapisa za radnika {k[0]} na dan {k[1]:%d.%m.%Y}.")
        vidjeno.add(k); imena[z["id"]].add(z["ime"])
    for rid, s in imena.items():
        if len(s) > 1: raise GreskaFormata(f"Radnik {rid} ima različita imena u izvozu: {', '.join(sorted(s))}.")
    return dict(od=od, do=do, zapisi=zapisi, listovi=L, format=fmt, putanja=str(path),
                naziv=Path(path).name, sha256=sha256_fajla(path))

# ======================================================================= tumačenje dana
def tumaci_dan(z, datum, u_izvozu, post, norma, kategorija=None):
    """Vraća (celije{kolona: vrijednost}, napomene[str], upozorenja[str]). Ništa se ne izmišlja."""
    c, nap, upz = {}, [], []
    vikend = datum.weekday() >= 5
    if z is not None and not z["ulaz"] and not z["izlaz"] and not z["otisci"] and not z["status"] and not z["ot"]:
        z = None
    if z is None:
        if not u_izvozu: return c, [MARKER_VAN], []
        if kategorija and kategorija != "Bez kategorije":
            return _kategorija(c, nap, kategorija, norma)
        return (c, [MARKER_NEMA], []) if not vikend else (c, [], [])
    # ---- prijave
    if z["otisci"] is None:                                   # izvor zna šta je ulaz, a šta izlaz
        if z["ulaz"]: c["D"] = z["ulaz"]
        if z["izlaz"]: c["E"] = z["izlaz"]
        if bool(z["ulaz"]) != bool(z["izlaz"]):
            nap.append("Nedostaje prijava " + ("odlaska" if z["ulaz"] else "dolaska")); upz.append(nap[-1])
    else:
        o = z["otisci"]; n = len(o)
        if n == 1:
            nap.append(f"Jedna prijava u {_hm(_t(o[0]))} – nije moguće utvrditi dolazak ili odlazak"); upz.append(nap[-1])
        elif n >= 2:
            c["D"], c["E"] = _t(o[0]), _t(o[-1])
            sred = o[1:-1]
            if n == 4:
                a, b = _t(sred[0]), _t(sred[1])
                if post["medjuizlazak"] == "pauza": c["F"], c["G"] = a, b; nap.append(f"Međuizlazak {_hm(a)}–{_hm(b)}")
                else: nap.append(f"Međuizlazak {_hm(a)}–{_hm(b)} (nije odbijen)")
            elif n > 2:
                nap.append(f"{prijava_txt(n)} ({', '.join(_hm(_t(x)) for x in o)}) – međuizlasci nisu odbijeni")
                upz.append(("Neparan broj prijava" if n % 2 else "Više međuizlazaka") + f": {nap[-1]}")
        if z["ponovljeni"]: nap.append(f"Ponovljena prijava zanemarena ({z['ponovljeni']}×)")
    if z["metoda"] == "šifra": nap.append("Prijava šifrom")
    if (c.get("D") or c.get("E")) and vikend: nap.append("Rad na neradni dan")
    if z["ot"]: c["I"] = z["ot"]; nap.append("Prekovremeni prema uređaju – potvrditi nalog")
    st = z["status"]
    if st in ("", "Prisutan") or st.startswith("Vikend"): pass
    elif st == "Kašnjenje": nap.append("Kašnjenje prema uređaju")
    elif not (c.get("D") or c.get("E")):
        if st == "Godišnji odmor" and kategorija in (None, "Godišnji odmor"):
            return _kategorija(c, nap, "Godišnji odmor", norma, "prema uređaju")
        if kategorija and kategorija != "Bez kategorije":
            return _kategorija(c, nap, kategorija, norma, f"uneseno ručno; uređaj: {st}")
        if st == "Godišnji odmor": return _kategorija(c, nap, "Godišnji odmor", norma, "uređaj")
        nap.append(f"{st} (uređaj) – kategorija nije unesena")
    else: nap.append(f"Status uređaja: {st}")
    if c.get("D") and c.get("E") and isinstance(z["sati_uredjaj"], (int, float)):
        h = sati_obrasca(c["D"], c["E"], c.get("F"), c.get("G"))
        if abs(h - float(z["sati_uredjaj"])) > 0.01: upz.append(f"Uređaj navodi {z['sati_uredjaj']} h, obrazac računa {h} h")
    return c, nap, upz

def _kategorija(c, nap, kat, norma, izvor="ručno"):
    kol = KATEGORIJE.get(kat)
    if kol:
        if norma: c[kol] = norma; nap.append(f"{kat} ({izvor})")
        else: nap.append(f"{kat} ({izvor}) – sati nisu upisani: u registru nema ugovorene sedmične satnice")
    else: nap.append(f"{kat} ({izvor})")
    return c, nap, ([] if norma or not kol else [nap[-1]])

def norma_radnika(r):
    try: return zaokruzi(float(r["sedmicno"]) / 5) if r.get("sedmicno") not in (None, "") else None
    except (TypeError, ValueError): return None

# ======================================================================= evidencija (1)
_T = Side(style="thin", color="B7C4CF"); _OKVIR = Border(left=_T, right=_T, top=_T, bottom=_T)
POMOCNI = {
    L_REGISTAR: (["ID uređaja", "Ime i prezime", "Ustanova", "Radno mjesto", "Ugovoreno sedmično (h)",
                  "Odgovorno lice", "Odjel na uređaju", "Upisan u registar"], [12, 26, 11, 30, 14, 26, 22, 17]),
    L_DNEVNIK: (["Vrijeme upisa", "Korisnik", "Izvorni fajl", "SHA-256 izvora", "Period izvoza", "Format izvoza",
                 "Novih šihtarica", "Dopunjenih", "Dopunjenih dana", "Upozorenja", "Postavke tumačenja", "Sigurnosna kopija"],
                [18, 13, 32, 18, 22, 22, 11, 11, 11, 11, 34, 38]),
    L_UPOZORENJA: (["Period", "ID", "Radnik", "Datum", "Izvor", "Opis", "Riješeno / potpis"], [9, 10, 24, 11, 18, 78, 20]),
}

def _stil_pomocnog(wb, naziv):
    if naziv in wb.sheetnames: return wb[naziv]
    ws = wb.create_sheet(naziv)
    zag, sir = POMOCNI[naziv]
    ws.append([naziv.upper()]); ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(zag))
    t = ws.cell(1, 1); t.font = Font(name="Aptos Display", size=14, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=BOJA["plava"]); t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    for c, (txt, w) in enumerate(zip(zag, sir), 1):
        x = ws.cell(2, c, txt)
        x.font = Font(name="Aptos", size=9, bold=True, color="FFFFFF"); x.fill = PatternFill("solid", fgColor=BOJA["svijetloplava"])
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); x.border = _OKVIR
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[2].height = 32
    ws.freeze_panes = "A3"; ws.auto_filter.ref = f"A2:{get_column_letter(len(zag))}2"
    ws.page_setup.orientation = "landscape"; ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:2"
    return ws

def _dodaj_red(ws, vrijednosti, formati=None):
    r = ws.max_row + 1
    for c, v in enumerate(vrijednosti, 1):
        x = ws.cell(r, c, v); x.border = _OKVIR
        x.font = Font(name="Calibri", size=10)
        x.alignment = Alignment(vertical="top", wrap_text=True)
        if r % 2 == 0: x.fill = PatternFill("solid", fgColor=BOJA["traka"])
        if formati and c in formati: x.number_format = formati[c]
    n = len(POMOCNI.get(ws.title, ([],))[0]) or len(vrijednosti)
    ws.auto_filter.ref = f"A2:{get_column_letter(n)}{r}"
    return r

def kreiraj_evidenciju(sablon, cilj):
    cilj = Path(cilj)
    if cilj.exists(): raise GreskaUpisa(f"Fajl već postoji i neće biti prepisan: {cilj}")
    wb = openpyxl.load_workbook(sablon)
    for n in (L_UPUTSTVO, L_PREGLED, *OBRAZAC.values()):
        if n not in wb.sheetnames: raise GreskaUpisa(f"Šablon nema propisani list '{n}'.")
    for n in (L_REGISTAR, L_DNEVNIK, L_UPOZORENJA): _stil_pomocnog(wb, n)
    wb.save(cilj)

PORUKA_ZAKLJUCANO = "Evidencija je otvorena u drugom programu (npr. Excelu). Zatvorite je i pokušajte ponovo – ništa nije promijenjeno."

def provjeri_evidenciju(path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True); s = wb.sheetnames; wb.close()
    except PermissionError:
        raise GreskaUpisa(PORUKA_ZAKLJUCANO)
    nema = [n for n in (L_PREGLED, *OBRAZAC.values()) if n not in s]
    if nema: raise GreskaUpisa("Fajl nije evidencija zasnovana na propisanom obrascu (nedostaje: " + ", ".join(nema) + ").")
    return s

def procitaj_registar(wb_ili_put):
    zatvori = not hasattr(wb_ili_put, "sheetnames")
    wb = openpyxl.load_workbook(wb_ili_put, read_only=True) if zatvori else wb_ili_put
    reg = {}
    if L_REGISTAR in wb.sheetnames:
        for r in wb[L_REGISTAR].iter_rows(min_row=3, values_only=True):
            if r and r[0] not in (None, ""):
                reg[_id(r[0])] = dict(id=_id(r[0]), ime=r[1] or "", inst=str(r[2] or "").strip().upper(),
                                      radno_mjesto=r[3] or "", sedmicno=r[4], odgovorno_lice=r[5] or "", odjel=r[6] or "")
    if zatvori: wb.close()
    return reg

def uvezeni_izvori(path):
    wb = openpyxl.load_workbook(path, read_only=True); out = {}
    if L_DNEVNIK in wb.sheetnames:
        for r in wb[L_DNEVNIK].iter_rows(min_row=3, values_only=True):
            if r and r[3]: out[str(r[3])] = r[0]
    wb.close(); return out

# ----------------------------------------------------------------------- plan
def napravi_plan(izv, evidencija, postavke=None):
    post = {**PODRAZUMIJEVANE_POSTAVKE, **(postavke or {})}
    listovi = set(provjeri_evidenciju(evidencija))
    reg = procitaj_registar(evidencija)
    radnici = collections.OrderedDict()
    for z in sorted(izv["zapisi"], key=lambda z: (z["id"], z["datum"])):
        if z["id"] in radnici: continue
        r = dict(reg[z["id"]]) if z["id"] in reg else None
        novi = r is None
        if novi:
            r = dict(id=z["id"], ime=z["ime"] or f"Radnik {z['id']}", inst="IMH" if "IMH" in z["odjel"].upper() else "IDSS",
                     radno_mjesto="", sedmicno=None, odgovorno_lice="", odjel=z["odjel"])
        r.update(novi=novi, ime_uredjaj=z["ime"], odjel=z["odjel"] or r.get("odjel", ""))
        radnici[z["id"]] = r
    po = collections.defaultdict(dict)
    for z in izv["zapisi"]: po[(z["id"], z["datum"].year, z["datum"].month)][z["datum"]] = z
    mjeseci = sorted({(z["datum"].year, z["datum"].month) for z in izv["zapisi"]})
    stavke, kategorizacija = [], []
    for g, m in mjeseci:
        for rid, r in radnici.items():
            dani = po.get((rid, g, m))
            if not dani: continue
            postojeci = next((n for n in listovi if NAZIV_LISTA.match(n) and n.startswith(f"{g}-{m:02d} ")
                              and n.split(" ", 2)[2] == naziv_lista(g, m, "X", rid).split(" ", 2)[2]), None)
            stavke.append(dict(godina=g, mjesec=m, id=rid, dani=dani, postojeci=postojeci,
                               naziv=postojeci or naziv_lista(g, m, r["inst"], rid)))
            for d in range(1, zadnji_dan(g, m) + 1):
                datum = dt.date(g, m, d)
                if not izv["od"] <= datum <= izv["do"]: continue
                z = dani.get(datum)
                prazan = z is None or (not z["ulaz"] and not z["izlaz"] and not z["otisci"])
                if not prazan: continue
                st = z["status"] if z else ""
                if z is None and datum.weekday() >= 5: continue
                if st.startswith("Vikend") or (st == "" and datum.weekday() >= 5): continue
                prijedlog = "Godišnji odmor" if st == "Godišnji odmor" else "Bez kategorije"
                kategorizacija.append(dict(id=rid, ime=r["ime"], datum=datum, razlog=f"Uređaj: {st}" if st else "Nema prijava",
                                           kategorija=prijedlog))
    kontrola = unakrsna_kontrola(izv)
    for r in radnici.values():
        if not r["novi"] and r["ime_uredjaj"] and r["ime"] != r["ime_uredjaj"]:
            kontrola.append((r["id"], None, "Registar", f"Ime u registru „{r['ime']}”, na uređaju „{r['ime_uredjaj']}” – koristi se ime iz registra"))
    vec = uvezeni_izvori(evidencija).get(izv["sha256"][:16])
    return dict(izvor=izv, radnici=radnici, stavke=stavke, kategorizacija=kategorizacija, kontrola=kontrola,
                postavke=post, vec_uvezen=vec)

def sazetak_plana(p):
    return dict(novih=sum(1 for s in p["stavke"] if not s["postojeci"]), dopuna=sum(1 for s in p["stavke"] if s["postojeci"]),
                radnika=len(p["radnici"]), novih_radnika=sum(r["novi"] for r in p["radnici"].values()),
                bez_kategorije=sum(1 for k in p["kategorizacija"] if k["kategorija"] == "Bez kategorije"),
                bez_satnice=sum(1 for r in p["radnici"].values() if norma_radnika(r) is None),
                neslaganja=len(p["kontrola"]))

# ----------------------------------------------------------------------- unakrsna kontrola izvora
def unakrsna_kontrola(izv):
    out, L = [], izv["listovi"]
    po = {(z["id"], z["datum"]): z for z in izv["zapisi"]}
    od, do = izv["od"], izv["do"]
    def ima_rada(z):
        return bool(z and (z["ulaz"] or z["izlaz"] or z["otisci"]))
    try:
        vik25 = False
        for row in L.get("Schedule information sheet", [])[4:]:
            if not row or not row[0]: continue
            rid = _id(row[0])
            for d in range((do - od).days + 1):
                datum, c = od + dt.timedelta(days=d), 3 + d
                if c >= len(row): break
                kod, z = row[c], po.get((rid, datum))
                if kod in (26, 26.0) and ima_rada(z):
                    out.append((rid, datum, "Raspored uređaja", "Raspored navodi službeni put, a postoje prijave na uređaju"))
                elif kod in (25, 25.0):
                    if datum.weekday() >= 5: vik25 = True
                    elif ima_rada(z): out.append((rid, datum, "Raspored uređaja", "Raspored navodi odsustvo, a postoje prijave na uređaju"))
        if vik25: out.append(("SVI", None, "Raspored uređaja", "Vikendi su u rasporedu uređaja označeni kodom 25 (odsustvo). Na šihtaricu nema uticaja."))
    except Exception as e: out.append(("SVI", None, "Raspored uređaja", f"List nije moguće provjeriti ({e})"))
    try:
        agg = collections.defaultdict(lambda: dict(h=0.0, ot=0.0))
        for z in izv["zapisi"]:
            d = z["ulaz"] or (_t(z["otisci"][0]) if z["otisci"] and len(z["otisci"]) >= 2 else None)
            e = z["izlaz"] or (_t(z["otisci"][-1]) if z["otisci"] and len(z["otisci"]) >= 2 else None)
            if d and e: agg[z["id"]]["h"] += sati_obrasca(d, e, None, None)
            agg[z["id"]]["ot"] += z["ot"]
        for row in L.get("Attendance summary", [])[4:]:
            if not row or not row[0] or len(row) < 15: continue
            rid = _id(row[0]); a = agg[rid]
            if isinstance(row[4], (int, float)) and abs(row[4] - a["h"]) > 0.05:
                out.append((rid, None, "Zbirni pregled uređaja", f"Stvarni sati: zbirni pregled {row[4]:g} h, prijave {a['h']:.2f} h"))
            ot = row[9].total_seconds() / 3600 if isinstance(row[9], dt.timedelta) else None
            if ot is not None and abs(ot - a["ot"]) > 0.01:
                out.append((rid, None, "Zbirni pregled uređaja", f"Prekovremeni: zbirni pregled {ot:g} h, prijave {a['ot']:g} h"))
    except Exception as e: out.append(("SVI", None, "Zbirni pregled uređaja", f"List nije moguće provjeriti ({e})"))
    try:
        for row in L.get("Abnormal", [])[4:]:
            if not row or not row[0]: continue
            rid, datum, t = _id(row[0]), _datum(row[3]), _vrijeme(row[4])
            z = po.get((rid, datum)); d = z and (z["ulaz"] or (_t(z["otisci"][0]) if z["otisci"] else None))
            if d != t: out.append((rid, datum, "Nepravilnosti uređaja", f"List nepravilnosti navodi dolazak {_hm(t)}, prijave: {(_hm(d) if d else 'nema')} – vrijede prijave"))
    except Exception as e: out.append(("SVI", None, "Nepravilnosti uređaja", f"List nije moguće provjeriti ({e})"))
    return out

# ----------------------------------------------------------------------- upis
def _kloniraj_obrazac(wb, inst, naziv):
    src = wb[OBRAZAC[inst]]
    ws = wb.copy_worksheet(src); ws.title = naziv
    ws.freeze_panes = src.freeze_panes
    for dv in src.data_validations.dataValidation: ws.add_data_validation(deepcopy(dv))
    ws.print_area = "A1:Q43"; ws.print_title_rows = "1:6"
    for a in ("orientation", "paperSize", "fitToWidth", "fitToHeight", "scale"): setattr(ws.page_setup, a, getattr(src.page_setup, a))
    ws.page_margins = copy(src.page_margins); ws.print_options = copy(src.print_options)
    ws.sheet_properties.pageSetUpPr = copy(src.sheet_properties.pageSetUpPr)
    ws.sheet_properties.tabColor = BOJA["plava"] if inst == "IDSS" else "7CB342"
    ws.oddFooter.left.text = "&8Šihtarica IDSS/IMH – generisano iz uređaja FA01H"; ws.oddFooter.right.text = "&8Stranica &P od &N"
    return ws

def _prazan_dan(ws, red):
    return all(ws[f"{k}{red}"].value in (None, "") for k in KOLONE_PODATAKA) and ws[f"P{red}"].value in (None, "", *MARKERI)

def _isto(a, b):
    if isinstance(a, dt.datetime): a = dt.time(a.hour, a.minute)
    if isinstance(b, dt.datetime): b = dt.time(b.hour, b.minute)
    if isinstance(a, (int, float)) and isinstance(b, (int, float)): return abs(a - b) < 1e-9
    return a == b

_AMBER = PatternFill("solid", fgColor=BOJA["amber"])

def _upisi_mjesec(ws, s, r, odluke, post, izv, dopuna):
    g, m = s["godina"], s["mjesec"]
    zaglavlje = {"B3": MJESECI[m - 1], "D3": g, "F3": r["ime"], "L3": r.get("radno_mjesto") or None,
                 "D4": r.get("sedmicno") if r.get("sedmicno") not in ("", None) else None,
                 "F4": r.get("odgovorno_lice") or post["odgovorno_lice"].get(r["inst"]) or None}
    for k, v in zaglavlje.items():
        if v is not None and ws[k].value in (None, ""): ws[k] = v
    norma = norma_radnika(r)
    upz, sukobi, popunjeno = [], [], 0
    for d in range(1, zadnji_dan(g, m) + 1):
        datum, red = dt.date(g, m, d), PRVI_RED + d - 1
        u_izvozu = izv["od"] <= datum <= izv["do"]
        z = s["dani"].get(datum)
        cel, nap, u = tumaci_dan(z, datum, u_izvozu, post, norma, odluke.get((r["id"], datum)))
        if dopuna and not _prazan_dan(ws, red):
            for k, v in cel.items():
                if not _isto(ws[f"{k}{red}"].value, v):
                    sukobi.append((datum, f"Kolona {k}: u evidenciji {ws[f'{k}{red}'].value!s}, novi izvoz {v!s} – zadržano postojeće"))
            continue
        if dopuna and not cel and (not nap or nap[0] in MARKERI): 
            if nap and ws[f"P{red}"].value in MARKERI and ws[f"P{red}"].value != nap[0]: ws[f"P{red}"] = nap[0]
            continue
        for k, v in cel.items(): ws[f"{k}{red}"] = v
        if dopuna and not nap and ws[f"P{red}"].value in MARKERI:
            ws[f"P{red}"] = None; ws[f"P{red}"].fill = copy(ws[f"O{red}"].fill)
        if nap:
            ws[f"P{red}"] = "; ".join(nap)
            ws[f"P{red}"].fill = _AMBER if (u or any(x in MARKERI for x in nap)) else copy(ws[f"O{red}"].fill)
        if dopuna and (cel or nap): popunjeno += 1
        upz += [(datum, x) for x in u]
    return upz, sukobi, popunjeno

def _red_pregleda(ws, list_naziv, inst, ime, g, m, red=None):
    if red is None:
        red = 6
        while any(ws.cell(red, c).value not in (None, "") for c in range(1, 8)): red += 1
    uzor = 6 if red % 2 == 0 else 7
    if red > 25:
        ws.row_dimensions[red].height = ws.row_dimensions[uzor].height
        for c in range(1, 8): ws.cell(red, c)._style = copy(ws.cell(uzor, c)._style)
    q = "'" + list_naziv.replace("'", "''") + "'!"
    ws.cell(red, 1, inst); ws.cell(red, 2, ime); ws.cell(red, 3, MJESECI[m - 1]); ws.cell(red, 4, g)
    ws.cell(red, 5, f"={q}H38"); ws.cell(red, 6, f"={q}I38")
    ws.cell(red, 7, f'=TRIM(IF({q}L38>0,"Godišnji "&{q}L38&" h; ","")&IF({q}M38>0,"Bolovanje "&{q}M38&" h; ","")'
                    f'&IF({q}N38>0,"Praznik "&{q}N38&" h; ","")&IF({q}J38>0,"Terenski "&{q}J38&" h; ","")'
                    f'&IF({q}O38>0,"Sedmični odmor "&{q}O38&" h; ",""))')
    for c in (5, 6): ws.cell(red, c).number_format = "0.00"
    for dv in ws.data_validations.dataValidation:
        mm = re.fullmatch(r"([AC])6:\1(\d+)", str(dv.sqref))
        if mm and int(mm.group(2)) < red: dv.sqref = MultiCellRange(f"{mm.group(1)}6:{mm.group(1)}{red}")
    ws.print_area = f"A1:G{max(red, 25)}"
    return red

def _sredi_redoslijed(wb):
    prednji = [L_UPUTSTVO, L_PREGLED, L_REGISTAR, L_DNEVNIK, L_UPOZORENJA, *OBRAZAC.values()]
    perioda = sorted([w for w in wb.worksheets if NAZIV_LISTA.match(w.title)], key=lambda w: w.title)
    ostali = [w for w in wb.worksheets if w.title not in prednji and not NAZIV_LISTA.match(w.title)]
    wb._sheets = [wb[n] for n in prednji if n in wb.sheetnames] + ostali + perioda
    wb.active = 0
    for w in wb.worksheets: w.sheet_view.tabSelected = (w.title == L_UPUTSTVO)

def _snimak(wb):
    return {ws.title: {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value not in (None, "")} for ws in wb.worksheets}

def dokaz_istorije(prije, poslije):
    """Svaka ranije popunjena ćelija mora ostati ista. Izuzeci: oznake praznog dana (P) i polja registra B–G."""
    greske = []
    for lst, celije in prije.items():
        if lst not in poslije: greske.append(f"List „{lst}” je nestao"); continue
        nove = poslije[lst]
        for koord, v in celije.items():
            if koord in nove and _isto(nove[koord], v): continue
            if NAZIV_LISTA.match(lst) and koord[0] == "P" and v in MARKERI: continue
            if lst == L_REGISTAR and koord[0] in "BCDEFG" and int(re.sub(r"\D", "", koord)) >= 3: continue
            greske.append(f"„{lst}”!{koord}: bilo „{v}”, sada „{nove.get(koord)}”")
            if len(greske) > 20: return greske
    return greske

def upisi(plan, evidencija, folder_kopija=None, napredak=lambda t: None):
    evidencija = Path(evidencija)
    try:
        with open(evidencija, "r+b"): pass
    except PermissionError:
        raise GreskaUpisa(PORUKA_ZAKLJUCANO)
    if plan["vec_uvezen"]:
        raise GreskaUpisa(f"Ovaj isti izvoz je već upisan ({plan['vec_uvezen']}). Ništa novo nije za upis.")
    napredak("Pravim sigurnosnu kopiju…")
    fk = Path(folder_kopija or evidencija.parent / "Sigurnosne_kopije"); fk.mkdir(exist_ok=True)
    kopija, n = fk / f"{evidencija.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx", 1
    while kopija.exists():                     # dva upisa u istoj sekundi ne smiju prepisati kopiju
        n += 1; kopija = fk / f"{evidencija.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}_{n}.xlsx"
    sigurna_kopija(evidencija, kopija)
    napredak("Učitavam evidenciju…")
    wb = openpyxl.load_workbook(evidencija)
    prije = _snimak(wb)
    for n in (L_REGISTAR, L_DNEVNIK, L_UPOZORENJA): _stil_pomocnog(wb, n)
    post, izv = plan["postavke"], plan["izvor"]
    # registar
    wsr = wb[L_REGISTAR]
    redovi = {_id(wsr.cell(i, 1).value): i for i in range(3, wsr.max_row + 1) if wsr.cell(i, 1).value not in (None, "")}
    for rid, r in plan["radnici"].items():
        vr = [rid, r["ime"], r["inst"], r.get("radno_mjesto") or None,
              r.get("sedmicno") if r.get("sedmicno") not in ("", None) else None, r.get("odgovorno_lice") or None, r.get("odjel") or None]
        if rid in redovi:
            for c, v in enumerate(vr[1:], 2):
                if v not in (None, ""): wsr.cell(redovi[rid], c, v)
        else: _dodaj_red(wsr, vr + [dt.datetime.now().strftime("%d.%m.%Y %H:%M")])
    odluke = {(k["id"], k["datum"]): k["kategorija"] for k in plan["kategorizacija"]}
    wsu, wsp = wb[L_UPOZORENJA], wb[L_PREGLED]
    postojece_upz = {tuple(str(x or "") for x in r[:6]) for r in wsu.iter_rows(min_row=3, values_only=True)}
    def upozori(per, rid, ime, datum, izvor, opis):
        v = [per, rid, ime, datum.strftime("%d.%m.%Y") if datum else "", izvor, opis, None]
        if tuple(str(x or "") for x in v[:6]) not in postojece_upz:
            _dodaj_red(wsu, v); postojece_upz.add(tuple(str(x or "") for x in v[:6])); return 1
        return 0
    novih, dopunjenih, dana_dop, br_upz, nazivi = 0, 0, 0, 0, []
    stavke = sorted(plan["stavke"], key=lambda s: (s["godina"], s["mjesec"], plan["radnici"][s["id"]]["inst"], plan["radnici"][s["id"]]["ime"]))
    for i, s in enumerate(stavke, 1):
        r = plan["radnici"][s["id"]]
        napredak(f"Prevodim {i}/{len(stavke)}: {r['ime']} – {MJESECI[s['mjesec']-1]} {s['godina']}")
        per = f"{s['godina']}-{s['mjesec']:02d}"
        if s["postojeci"]:
            ws = wb[s["postojeci"]]
            upz, sukobi, pop = _upisi_mjesec(ws, s, r, odluke, post, izv, dopuna=True)
            if pop: dopunjenih += 1; dana_dop += pop
            for datum, t in sukobi: br_upz += upozori(per, s["id"], r["ime"], datum, "Dopuna", t)
        else:
            inst = r["inst"]
            ws = _kloniraj_obrazac(wb, inst, naziv_lista(s["godina"], s["mjesec"], inst, s["id"]))
            upz, _, _ = _upisi_mjesec(ws, s, r, odluke, post, izv, dopuna=False)
            _red_pregleda(wsp, ws.title, inst, r["ime"], s["godina"], s["mjesec"])
            novih += 1; nazivi.append(ws.title)
        for datum, t in upz: br_upz += upozori(per, s["id"], r["ime"], datum, "Prijave na uređaju", t)
    per_izv = f"{izv['od']:%Y-%m}"
    for rid, datum, izvor, opis in plan["kontrola"]:
        br_upz += upozori(per_izv, rid, plan["radnici"].get(rid, {}).get("ime", ""), datum, izvor, opis)
    _dodaj_red(wb[L_DNEVNIK], [dt.datetime.now().strftime("%d.%m.%Y %H:%M:%S"), getpass.getuser(), izv["naziv"], izv["sha256"][:16],
               f"{izv['od']:%d.%m.%Y} – {izv['do']:%d.%m.%Y}", izv["format"], novih, dopunjenih, dana_dop, br_upz,
               f"ponovljena prijava ≤{post['prag_ponovljenog_otiska_min']} min; međuizlazak: {post['medjuizlazak']}", kopija.name])
    _sredi_redoslijed(wb)
    wb.calculation.fullCalcOnLoad = True
    napredak("Provjeravam da istorija nije promijenjena…")
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=evidencija.parent); os.close(fd)
    try:
        wb.save(tmp)
        greske = dokaz_istorije(prije, _snimak(openpyxl.load_workbook(tmp)))
        if greske: raise GreskaUpisa("Dokaz nepromjenjivosti istorije nije prošao – upis je poništen, evidencija je netaknuta.\n• " + "\n• ".join(greske[:10]))
        napredak("Spremam evidenciju…")
        os.replace(tmp, evidencija)
    except PermissionError:
        raise GreskaUpisa("Evidencija je zaključana (otvorena u Excelu?). Ništa nije promijenjeno.")
    finally:
        if os.path.exists(tmp): os.remove(tmp)
    return dict(novih=novih, dopunjenih=dopunjenih, dana_dopunjeno=dana_dop, upozorenja=br_upz, kopija=str(kopija), listovi=nazivi)
