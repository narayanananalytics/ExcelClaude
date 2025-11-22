@echo off
REM Pre-Build Cleanup Script
REM Add this to your Visual Studio project's Pre-Build events
REM Project Properties -> Build Events -> Pre-build event command line

echo Running pre-build cleanup...

REM Close Excel if running
taskkill /F /IM EXCEL.EXE 2>nul
if %errorlevel% == 0 (
    echo Closed running Excel instance
    timeout /t 2 /nobreak >nul
)

REM Clear ClickOnce cache for this project
rundll32 dfshim CleanOnlineAppCache 2>nul

echo Pre-build cleanup complete
exit 0
