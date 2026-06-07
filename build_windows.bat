@echo off
setlocal
cd /d "%~dp0"

py -3 -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --clean --noconfirm drug_usage_trend.spec

echo.
echo Build complete.
echo EXE: dist\약품사용량추세.exe
pause
