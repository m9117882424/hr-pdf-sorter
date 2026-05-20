@echo off
chcp 65001 > nul

cd /d "%~dp0"

py ru_decont_fast.py --excel "excel_input\input.xlsx" --pdf-dir "pp_ru" --workers 3

pause
