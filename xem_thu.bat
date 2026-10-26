@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
set OCAR_DATA=%~dp0..\DuLieu
set PY=python
where python >nul 2>nul || set PY=py
%PY% tools\build_monitor.py || (pause & exit /b 1)
start "" "docs\index.html"
pause
