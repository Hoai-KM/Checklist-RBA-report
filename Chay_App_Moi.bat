@echo off
rem ============================================================
rem  CHAY APP BAO CAO HR AUDIT RBA (Samsung SEVT) - PHIEN BAN MOI
rem  Bam dup chuot vao file nay de mo giao dien app.
rem  File nay tu dong:
rem    1) Kiem tra Python da cai dat (bao loi neu chua co)
rem    2) Kiem tra thu vien openpyxl (tu dong cai dat neu thieu)
rem    3) Khoi dong giao dien bang pythonw (KHONG hien cua so
rem       console den, chi co cua so app)
rem ============================================================
title HR Audit Report Editor v3.1 - dang khoi dong...
cd /d "%~dp0"

rem ---- 1) Kiem tra Python ----
where python >nul 2>nul
if errorlevel 1 goto nopython

rem ---- 2) Kiem tra thu vien openpyxl ----
python -c "import openpyxl" >nul 2>nul
if errorlevel 1 goto nodeps

rem ---- 3) Khoi dong app ----
:launch
set "PYW="
for /f "delims=" %%i in ('where python 2^>nul') do (
    if exist "%%~dpi\pythonw.exe" set "PYW=%%~dpi\pythonw.exe"
)
if defined PYW (
    start "" "%PYW%" run_app.py
) else (
    start "" python run_app.py
)
exit /b 0

:nodeps
echo [LOI] Chua cai thu vien openpyxl - dang tu dong cai dat...
python -m pip install -r requirements.txt
if errorlevel 1 goto pipfail
echo.
echo Da cai dat xong. Dang khoi dong app...
goto launch

:pipfail
echo.
echo [LOI] Khong cai dat duoc thu vien. Kiem tra ket noi mang roi thu lai,
echo hoac chay thu cong:  python -m pip install -r requirements.txt
pause
exit /b 1

:nopython
echo [LOI] Khong tim thay Python tren may tinh.
echo.
echo Hay cai dat Python 3.10 hoac moi hon tai:
echo     https://www.python.org/downloads/
echo Luu y: khi cai dat, tick chon "Add Python to PATH".
echo.
pause
exit /b 1