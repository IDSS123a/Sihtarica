# Šihtarica IDSS · IMH — v2.1

Desktop aplikacija (Windows, radi bez interneta) koja prevodi izvoz uređaja FA01H („All Report”, .xls/.xlsx)
u propisani obrazac šihtarice i vodi **jedan trajni evidencijski fajl** s istorijom svih perioda.

## Gotova aplikacija (bez instalacije bilo čega)
Svaki push na `main` automatski, na pravom Windows računaru (GitHub Actions):
sastavi `Sihtarica_IDSS_IMH.exe`, pokrene 11 provjera na izvornom kodu i na gotovom .exe-u, i objavi ih.
**Actions → posljednje pokretanje → Artifacts → `Sihtarica_IDSS_IMH_Windows`** (zip s .exe-om i rezultatom provjere).
Rezultat provjere vidi se i direktno na stranici pokretanja (Summary).

## Pokretanje
Gotova aplikacija je **`Sihtarica_IDSS_IMH.exe`** – jedan fajl, bez instalacije, radi bez interneta. Pri prvom pokretanju
Windows može prikazati „Windows je zaštitio vaš PC” (aplikacija nije digitalno potpisana): **Više informacija → Ipak pokreni**.

Ponovna izrada iz izvornog koda (samo ako se kod mijenja): instalirajte Python 3.12 s python.org („Add python.exe to PATH”)
i dvokliknite `build.bat` – napravi `dist\Sihtarica_IDSS_IMH.exe` i odmah ga provjeri.

## Provjera na vašem računaru (preporučeno jednom)
U komandnoj liniji: `Sihtarica_IDSS_IMH.exe --provjera C:\Provjera`
Aplikacija sama napravi testne izvoze, prođe 11 provjera (uključujući fajl otvoren u Excelu i putanje s č/ć/š)
i upiše `C:\Provjera\rezultat_provjere.txt`. Stvarna evidencija se ne dira.

## Rad
1. **Novi…** (lijevo dolje) kreira evidenciju iz propisanog obrasca; **Otvori…** otvara postojeću. Aplikacija je pamti.
2. **Uvoz s uređaja**: prevucite All Report. Pregledajte radnike (satnica, radno mjesto, odgovorno lice – pamti se trajno),
   po želji dodijelite kategoriju danima bez prijava, zatim **Prevedi i upiši u evidenciju**.
3. **Evidencija**: pregled svih šihtarica po periodima. **Izvoz podataka**: .xlsx, .pdf, .json za period/ustanovu ili sve.

## Pravila prevođenja (ništa se ne izmišlja)
| Situacija u izvozu | Šihtarica |
|---|---|
| 2 prijave | Početak = prva, završetak = druga |
| 4 prijave | Početak/završetak + međuizlazak u kolone pauze (ili samo napomena – Postavke) |
| 1 prijava | Kolone vremena prazne; vrijeme doslovno u napomeni, dan označen |
| 3 ili 5+ prijava | Prva/zadnja prijava; sve prijave u napomeni; međuizlasci se ne odbijaju; dan označen |
| Ponovljena prijava (≤2 min) | Računa se jednom; navedeno u napomeni |
| Prijava šifrom (ako izvoz sadrži način prijave) | Navedeno u napomeni |
| Nema prijava (radni dan) | „Nema podataka u izvozu uređaja” – može se kasnije dopuniti |
| Dan izvan perioda izvoza | „Dan nije obuhvaćen izvozom uređaja” – može se kasnije dopuniti |
| Status uređaja (godišnji, odsustvo…) | Godišnji: sati = sedmična satnica / 5; ostalo: napomena ili ručna kategorija |
| Prekovremeni prema uređaju | Kolona I + napomena „potvrditi nalog” |

## Garancije istorije
- Svaka ranije popunjena ćelija ostaje identična — provjerava se ćeliju po ćeliju nakon svakog upisa; ako provjera ne prođe, upis se poništava.
- Dopuna popunjava samo prazne dane; razlika u već popunjenom danu se ne upisuje nego bilježi u „Upozorenja”.
- Isti izvoz (isti SHA-256) ne može se upisati dvaput. Sigurnosna kopija prije svakog upisa (`Sigurnosne_kopije`).
- „Dnevnik uvoza” bilježi vrijeme, korisnika, izvor, njegov SHA-256, postavke tumačenja i sigurnosnu kopiju.

## Fajlovi
`app.py` (interfejs) · `engine.py` (prevođenje i upis) · `izvoz.py` (xlsx/json/pdf) · `sablon_Sihtarica_IDSS_IMH.xlsx` (propisani obrazac) · `fonts/` (Carlito, SIL OFL) · `provjera.py` (11 automatskih provjera) · `samotest.py` (provjera kroz prozor aplikacije)
