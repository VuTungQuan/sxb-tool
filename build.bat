@echo off
REM Build SXBTool.exe (can pip install pyinstaller openpyxl pymupdf python-docx)
python -m PyInstaller --onefile --noconsole --name SXBTool --distpath dist --workpath build\work --specpath build sxb_tool.py
echo.
echo Xong: dist\SXBTool.exe
pause
