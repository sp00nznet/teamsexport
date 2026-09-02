@echo off
rem Builds dist\teamsexport.exe (console) and dist\teamsexport-gui.exe (windowed).
rem Builds in a venv: PyInstaller refuses to run if the obsolete `pathlib` PyPI
rem backport is installed globally, and this sidesteps whatever else is in there.
rem --collect-submodules is needed because iter_records() imports ccl lazily, so
rem PyInstaller's static analysis doesn't see it.
setlocal
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe python -m venv .venv || exit /b 1
set PY=.venv\Scripts\python.exe
%PY% -m pip install -q --upgrade pip pyinstaller -r requirements.txt || exit /b 1
%PY% -m PyInstaller --noconfirm --onefile --console  --name teamsexport ^
    --collect-submodules ccl_chromium_reader teamsexport.py || exit /b 1
%PY% -m PyInstaller --noconfirm --onefile --windowed --name teamsexport-gui ^
    --collect-submodules ccl_chromium_reader gui.py || exit /b 1
echo.
dir /b dist
