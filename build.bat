@echo off
REM Build Tool SXB: dist\SXBTool.exe (ban 2 tab) va dist\SXBTool_4tab.exe (ban 4 tab).
REM Tu tim Python, tu cai thu vien con thieu, chay tu kiem tra roi moi build.
REM Muon dung Python cu the: set PYTHON="C:\duong dan\python.exe" roi chay build.bat
setlocal
cd /d "%~dp0"
title Build Tool SXB
echo === Build Tool SXB: SXBTool.exe (2 tab) va SXBTool_4tab.exe (4 tab) ===
echo.

REM ---- 1. Tim Python 3.9 tro len: bien PYTHON ^> py -3.11 ^> py -3 ^> python
if defined PYTHON (set "PY=%PYTHON%" & goto co_python)
for %%P in ("py -3.11" "py -3" "python") do (
    %%~P -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1 && (set "PY=%%~P" & goto co_python)
)
echo [LOI] Khong tim thay Python 3.9 tro len.
echo       Cai Python 3.11 tu https://www.python.org/downloads/ (tick "Add python.exe to PATH") roi chay lai.
goto loi
:co_python
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1
if errorlevel 1 (echo [LOI] %PY% khong chay duoc hoac phien ban duoi 3.9.& goto loi)
%PY% -c "import sys; print('Python', sys.version.split()[0], '-', sys.executable)"

REM ---- 2. pip va cac thu vien (ten import / ten goi pip)
%PY% -m pip --version >nul 2>&1 || %PY% -m ensurepip --upgrade >nul 2>&1
%PY% -m pip --version >nul 2>&1
if errorlevel 1 (echo [LOI] Python nay khong co pip.& goto loi)
echo Kiem tra thu vien:
call :kiem openpyxl openpyxl || goto loi
REM pymupdf ghim duoi 2.0: ban 2.0 keo theo onnxruntime, numpy... (cai lau, exe nang hon); tool da thu voi 1.26
call :kiem fitz "pymupdf>=1.24,<2" || goto loi
call :kiem docx python-docx || goto loi
call :kiem win32com.client pywin32 || goto loi
call :kiem PyInstaller pyinstaller || goto loi
echo.

REM ---- 3. Tu kiem tra truoc khi build
echo Tu kiem tra (sxb_tool.py --test)...
%PY% -X utf8 sxb_tool.py --test
if errorlevel 1 (echo [LOI] Tu kiem tra that bai - khong build.& goto loi)
echo.

REM ---- 4. Build 2 ban exe
echo Build dist\SXBTool.exe (2 tab)...
%PY% -m PyInstaller --noconfirm --onefile --noconsole --log-level WARN --name SXBTool --distpath dist --workpath build\work --specpath build sxb_tool.py
if errorlevel 1 goto loi_build
echo Build dist\SXBTool_4tab.exe (4 tab)...
%PY% -m PyInstaller --noconfirm --onefile --noconsole --log-level WARN --name SXBTool_4tab --distpath dist --workpath build\work --specpath build sxb_tool_4tab.py
if errorlevel 1 goto loi_build
echo.
echo === XONG ===
echo   dist\SXBTool.exe      : ban 2 tab (Doi chieu ^| Doi ten ^& luu)
echo   dist\SXBTool_4tab.exe : ban 4 tab (Quet kho ^| Doi chieu Excel ^| Kiem tra noi dung ^| Doi ten ^& luu)
pause
exit /b 0

:kiem
REM %1 = ten import, %2 = ten goi pip; thieu thi cai
%PY% -c "import %1" >nul 2>&1
if not errorlevel 1 (echo   [co san] %2& exit /b 0)
echo   [thieu]  %2 - dang cai...
%PY% -m pip install --disable-pip-version-check %2
if errorlevel 1 (echo [LOI] Khong cai duoc %2 - kiem tra ket noi Internet hoac quyen ghi thu muc Python.& exit /b 1)
%PY% -c "import %1" >nul 2>&1
if errorlevel 1 (echo [LOI] Da cai %2 nhung van khong import duoc %1.& exit /b 1)
echo   [da cai] %2
exit /b 0

:loi_build
echo [LOI] Build that bai - xem thong bao phia tren. Neu SXBTool.exe dang mo thi dong lai roi chay lai.
:loi
echo.
pause
exit /b 1
