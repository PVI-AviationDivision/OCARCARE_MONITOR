@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
set OCAR_DATA=%~dp0..\DuLieu
echo ==============================================
echo   OCARCARE MONITOR - Cap nhat bao cao hang ngay
echo ==============================================
set PY=python
where python >nul 2>nul || set PY=py

rem ---- Den het 16/11/2026 (khoa so 15/11): cap nhat ca trang Sales Contest (dung chung du lieu DuLieu\)
set CONTEST=
for /d %%F in ("D:\2. Sales Contest*") do if exist "%%F\OCarCare_Dashboard\update.bat" set "CONTEST=%%F\OCarCare_Dashboard"
for /f %%d in ('powershell -NoProfile -Command "(Get-Date).ToString('yyyyMMdd')"') do set TODAY=%%d
if %TODAY% LEQ 20261116 if defined CONTEST (
  echo.
  echo ---- [1/2] Trang Sales Contest ----
  call "%CONTEST%\update.bat" auto
  cd /d "%~dp0"
  echo.
  echo ---- [2/2] Trang Monitor ----
)

echo Dong bo voi GitHub...
git pull --rebase --autostash -q origin main
if errorlevel 1 echo [CANH BAO] Chua dong bo duoc voi GitHub - van tiep tuc
echo.

%PY% tools\build_monitor.py
if errorlevel 1 (
  echo.
  echo [LOI] Khong tao duoc du lieu. Kiem tra file trong thu muc DuLieu\
  if "%1"=="" pause
  exit /b 1
)

git add docs tools *.bat README.md .gitignore
git status --short | findstr /i ".xls matkhau data_kiemtra" && (echo [DUNG] Phat hien file khong duoc phep day len! & pause & exit /b 1)
git diff --cached --quiet
if not errorlevel 1 (
  echo Khong co thay doi so voi lan cap nhat truoc.
  goto end
)
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HHmm"') do set TS=%%i
git commit -q -m "Cap nhat du lieu %TS%"
git push -q
if errorlevel 1 (
  echo [LOI] Day len GitHub that bai. Kiem tra mang / dang nhap GitHub.
  if "%1"=="" pause
  exit /b 1
)
echo.
echo [OK] Da day len GitHub. Trang se cap nhat sau khoang 1 phut:
echo https://pvi-aviationdivision.github.io/OCARCARE_MONITOR/

:end
if "%1"=="" pause
