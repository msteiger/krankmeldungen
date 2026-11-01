@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Installation fehlgeschlagen. Bitte Python 3 installieren und erneut starten.
  pause
  exit /b 1
)
python mailplanung.py
if errorlevel 1 pause
