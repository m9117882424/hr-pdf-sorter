@echo off
chcp 65001 > nul

cd /d "%~dp0"

if not exist "ru_decont.py" (
    echo [ERROR] File ru_decont.py not found in this folder.
    echo ru_decont_fast.py uses ru_decont.py as the core parser.
    echo Run: git pull origin main
    echo Or copy ru_decont.py from the repository into this folder.
    pause
    exit /b 1
)

if not exist "ru_decont_fast.py" (
    echo [ERROR] File ru_decont_fast.py not found in this folder.
    echo Run: git pull origin main
    pause
    exit /b 1
)

py ru_decont_fast.py --excel "excel_input\input.xlsx" --pdf-dir "pp_ru" --workers 3

pause
