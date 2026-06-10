@echo off
chcp 65001 >nul
cd /d %~dp0

python -m pip install pyinstaller
pyinstaller -D main.py -n "PDF面单一条龙工具" ^
  --hidden-import zxingcpp ^
  --hidden-import pytesseract ^
  --distpath dist_stable ^
  --workpath build_stable ^
  --noconfirm --clean

echo.
echo EXE output: %cd%\dist_stable\PDF面单一条龙工具\PDF面单一条龙工具.exe
pause
