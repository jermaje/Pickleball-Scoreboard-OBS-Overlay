@echo off
REM Run the Pickleball scoreboard server from this folder.
REM Usage: run_scoreboard.bat [port] [autosave interval seconds]

cd /d "%~dp0"
set PORT=8000
if not "%1"=="" set PORT=%1
set AUTOSAVE=5
if not "%2"=="" set AUTOSAVE=%2
echo Removing obsolete state.json if present...
if exist state.json del /f state.json 2>nul
echo Checking for openpyxl...
python -c "import openpyxl" 2>nul || (
    echo openpyxl not found, installing...
    python -m pip install openpyxl || (
        echo Failed to install openpyxl. Please install it manually.
        pause
        exit /b 1
    )
)

echo Using state.xlsx for persistence
echo Starting Pickleball scoreboard server on port %PORT% with autosave every %AUTOSAVE% seconds...
python server.py %PORT% %AUTOSAVE%
