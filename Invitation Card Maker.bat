@echo off
rem Double-click to start Invitation Card Maker.
rem The first run (and the first run after an update) installs what the app needs
rem into a private Python environment in .venv. This takes a minute.
setlocal
cd /d "%~dp0"

set "VENV=.venv"
set "MARKER=%VENV%\requirements.installed"

rem Skip setup when the installed requirements match requirements.txt
if exist "%MARKER%" fc /b requirements.txt "%MARKER%" >nul 2>nul && goto launch

echo Setting up Invitation Card Maker...
echo.

if exist "%VENV%\Scripts\python.exe" goto install

rem Find Python 3 (the "py" launcher first, then "python" on PATH)
set "PY="
py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY python --version >nul 2>nul && set "PY=python"
if not defined PY (
    echo Python 3 was not found.
    echo Install it from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during setup, then run this file again.
    echo.
    pause
    exit /b 1
)
%PY% -m venv "%VENV%" || goto failed

:install
"%VENV%\Scripts\python.exe" -m pip install --upgrade pip --quiet --disable-pip-version-check
"%VENV%\Scripts\python.exe" -m pip install -r requirements.txt --quiet --disable-pip-version-check || goto failed
copy /y requirements.txt "%MARKER%" >nul
echo Setup complete.

:launch
start "" "%VENV%\Scripts\pythonw.exe" invitation_app.py
exit /b 0

:failed
echo.
echo Setup failed. Check your internet connection and try again.
echo To start over, delete the .venv folder next to this file.
echo.
pause
exit /b 1
