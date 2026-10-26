@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
set OCAR_DATA=%~dp0..\DuLieu
echo ==========================================================
echo   OCARCARE MONITOR - KET NOI GITHUB (chi chay 1 lan)
echo ==========================================================
set REPO=https://github.com/PVI-AviationDivision/OCARCARE_MONITOR.git

where git >nul 2>nul || (echo [LOI] May chua cai Git & pause & exit /b 1)
set PY=python
where python >nul 2>nul || set PY=py
%PY% --version || (echo [LOI] May chua cai Python & pause & exit /b 1)
%PY% -m pip install --quiet openpyxl

echo.
echo [1/4] Mat khau xem bao cao
if exist matkhau.txt (
  echo Da co file matkhau.txt - giu nguyen.
) else (
  set /p PW=Nhap mat khau muon dat cho trang bao cao: 
  call echo %%PW%%> matkhau.txt
  echo Da luu vao matkhau.txt ^(file nay KHONG day len GitHub^)
)

echo.
echo [2/4] Tao du lieu tu thu muc DuLieu\ ...
%PY% tools\build_monitor.py || (pause & exit /b 1)

echo.
echo [3/4] Ket noi repo GitHub...
if not exist ".git" git init -q
git config user.email >nul 2>nul || git config user.email "ocarcare-monitor@users.noreply.github.com"
git config user.name >nul 2>nul || git config user.name "OCarCare Monitor"
git remote remove origin >nul 2>nul
git remote add origin %REPO%
git fetch -q origin
git ls-remote --exit-code --heads origin main >nul 2>nul && git reset -q origin/main
git branch -M main

echo [4/4] Day len GitHub...
git add docs tools *.bat README.md .gitignore
git status --short | findstr /i ".xls matkhau data_kiemtra" && (echo [DUNG] Phat hien file khong duoc phep day len! & pause & exit /b 1)
git commit -q -m "Khoi tao OCARCARE Monitor" || (echo [LOI] Khong tao duoc commit & pause & exit /b 1)
git push -u origin main || (echo [LOI] Day len that bai - kiem tra quyen ghi vao repo & pause & exit /b 1)

echo.
echo [OK] Xong. Buoc cuoi: vao GitHub - repo OCARCARE_MONITOR - Settings - Pages
echo      Source: Deploy from a branch - Branch "main", thu muc "/docs" - Save.
echo Link: https://pvi-aviationdivision.github.io/OCARCARE_MONITOR/
pause
