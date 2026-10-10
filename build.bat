@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo   Sihtarica IDSS/IMH - izrada Windows aplikacije (.exe)
echo ============================================================
set PY=
py -3.12 --version >nul 2>nul && set PY=py -3.12
if not defined PY python --version 2>nul | findstr /c:"3.12" >nul && set PY=python
if not defined PY (echo Potreban je Python 3.12 s python.org ^(oznacite "Add python.exe to PATH"^). & pause & exit /b 1)
if not exist .venv %PY% -m venv .venv || (echo Greska pri kreiranju okruzenja. & pause & exit /b 1)
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt || (echo Greska pri instalaciji paketa. & pause & exit /b 1)
pyinstaller --noconfirm --clean --onefile --windowed --name Sihtarica_IDSS_IMH ^
  --add-data "sablon_Sihtarica_IDSS_IMH.xlsx;." --add-data "fonts;fonts" app.py || (echo Greska pri izradi. & pause & exit /b 1)
echo.
echo Provjera gotove aplikacije...
dist\Sihtarica_IDSS_IMH.exe --provjera "%~dp0provjera_rezultat"
type "%~dp0provjera_rezultat\rezultat_provjere.txt"
echo.
echo GOTOVO: dist\Sihtarica_IDSS_IMH.exe
pause
