@echo off
rem ============================================================
rem  SINH BAO CAO HR AUDIT RBA (Samsung SEVT)
rem  Bam dup chuot vao file nay de sinh bao cáo.
rem ============================================================
title HR Audit Report Generator v3.1
cd /d "%~dp0"

echo ============================================
echo   HR Audit Report Generator v3.1
echo ============================================
echo.

rem ---- Kiem tra Python ----
where python >nul 2>nul
if errorlevel 1 (
    echo [LOI] Khong tim thay Python tren may tinh.
    echo Hay cai dat Python 3.10 hoac moi hon tai:
    echo     https://www.python.org/downloads/
    pause
    exit /b 1
)

rem ---- Kiem tra thu vien openpyxl ----
python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo [LOI] Chua cai thu vien openpyxl - dang tu dong cai dat...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [LOI] Khong cai dat duoc thu vien.
        pause
        exit /b 1
    )
)

rem ---- Sinh bao cáo ----
echo Dang sinh bao cao...
echo.

python -m hr_audit_app.cli ^
    --in "Copy of 2026.03.11_L10PO5_YOUSUNG_HR audit result report in 2H.2026.xlsx" ^
    --out "2026.08.26_L10PO5_YOUSUNG_HR audit result report in 2H.2026.xlsx" ^
    --audit-date 2026.08.26 ^
    --audit-type "Corrective Action Verification" ^
    --team-leader "KHUẤT MẠNH HOÀI" ^
    --team-member "NGUYỄN TIẾN THỊNH (SEVT)" ^
    --closing-date 2026-08-26 ^
    --interview-count 11 ^
    --interview-score 0.894 ^
    --profile-year 2025

echo.
echo ============================================
echo   HOAN THANH!
echo ============================================
echo.
echo File da duoc sinh: 2026.08.26_L10PO5_YOUSUNG_HR audit result report in 2H.2026.xlsx
echo.
echo Noi dung da co format moi:
echo   - 4 header in dam: AUDIT FINDINGS (ENGLISH), Legal references:,
echo     NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT), Căn cứ pháp lý:
echo   - Dong phan cach: ---------------------------------------
echo.
pause