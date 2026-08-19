@echo off
REM ============================================================
REM  Synchronoss Unified Toolbox – one-click GUI launcher (Windows)
REM ============================================================
REM  Double-click this file to start the GUI without a command
REM  prompt window (uses pythonw / pyw).
REM ============================================================

cd /d "%~dp0"

REM Prefer the windowed Python launcher so no console appears
where pyw >nul 2>&1
if %ERRORLEVEL%==0 (
    start "" pyw -3 -m synchronoss_parser.toolbox_gui
    exit /b 0
)

where pythonw >nul 2>&1
if %ERRORLEVEL%==0 (
    start "" pythonw -m synchronoss_parser.toolbox_gui
    exit /b 0
)

REM Fall back to console Python so the user can see any error
where py >nul 2>&1
if %ERRORLEVEL%==0 (
    py -3 -m synchronoss_parser.toolbox_gui
) else (
    python -m synchronoss_parser.toolbox_gui
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo --------------------------------------------------------
    echo  The GUI failed to start.
    echo.
    echo  Common causes:
    echo    1. Python is not installed or not on PATH
    echo    2. Required packages are missing  (run:  pip install -r requirements.txt)
    echo    3. You are running an old copy of the files
    echo.
    echo  If the error mentions "No module named synchronoss_parser",
    echo  make sure you are launching from the folder that CONTAINS
    echo  the "synchronoss_parser" directory (this .bat does that
    echo  automatically).
    echo --------------------------------------------------------
    pause
)
