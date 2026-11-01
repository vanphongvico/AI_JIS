@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

REM Chay truc tiep, khong can file exe. Danh cho may DA CAI PYTHON.

python --version >nul 2>&1
if errorlevel 1 goto KHONG_CO_PYTHON

python -c "import openpyxl, lxml, pypdfium2, PIL, cryptography" >nul 2>&1
if errorlevel 1 (
    echo Lan dau chay - dang cai thu vien can thiet, xin cho...
    python -m pip install openpyxl lxml pypdfium2 pillow cryptography
    if errorlevel 1 goto LOI_MANG
)

start "" pythonw "nguon\app_charpy.py"
exit /b 0

:KHONG_CO_PYTHON
echo.
echo   May nay chua cai Python.
echo.
echo   Cach 1 (de nhat): dung file  BaoCaoCoTinh.exe  trong thu muc nay.
echo   Cach 2: cai Python tai  https://www.python.org/downloads/
echo           nho tick "Add python.exe to PATH" khi cai.
echo.
pause
exit /b 1

:LOI_MANG
echo.
echo   Khong tai duoc thu vien. Kiem tra ket noi internet roi thu lai.
echo.
pause
exit /b 1
