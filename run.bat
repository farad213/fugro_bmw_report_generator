@echo off
setlocal enabledelayedexpansion

set "line1=pbar = tqdm(total=1)"
set "line2=pbar.update(1)"
set "file=venv/Lib/site-packages/docx2pdf/__init__.py"

set "found1="
set "found2="

for /f "tokens=*" %%a in ('type "%file%" ^| findstr /n "^"') do (
    set "line=%%a"
    set "line=!line:*:=!"
    if "!line!"=="!line1!" set "found1=true"
    if "!line!"=="!line2!" set "found2=true"
)

if defined found1 if defined found2 (
    for /f "usebackq delims=" %%a in ("%file%") do (
        if not "%%a"=="%line1%" if not "%%a"=="%line2%" (
            echo %%a>>"%file%.new"
        )
    )
    move /y "%file%.new" "%file%" >nul
)

endlocal

python main.py
pause