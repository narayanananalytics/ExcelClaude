@echo off
REM VSTO Add-in Cleanup Script
REM Run this as Administrator to completely remove VSTO add-in installations

echo ========================================
echo VSTO Add-in Cleanup Script
echo ========================================
echo.
echo This script will:
echo 1. Close Excel
echo 2. Clear ClickOnce cache
echo 3. Remove registry entries
echo 4. Clean deployment manifests
echo.
echo Press Ctrl+C to cancel, or
pause

echo.
echo Step 1: Closing Excel...
taskkill /F /IM EXCEL.EXE 2>nul
if %errorlevel% == 0 (
    echo Excel closed successfully
) else (
    echo Excel was not running
)

echo.
echo Step 2: Clearing ClickOnce cache...
echo This may take a minute...

REM Clear ClickOnce application cache
rundll32 dfshim CleanOnlineAppCache

REM Alternative method - delete cache folders directly
if exist "%LocalAppData%\Apps\2.0" (
    echo Removing ClickOnce cache folders...
    rd /s /q "%LocalAppData%\Apps\2.0" 2>nul
)

echo ClickOnce cache cleared

echo.
echo Step 3: Removing registry entries...

REM Remove Excel add-in registry entries
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\ExcelPerplexityVSTO" /f 2>nul
if %errorlevel% == 0 (
    echo Registry entries removed
) else (
    echo No registry entries found (this is OK)
)

REM Remove VSTO deployment entries
reg delete "HKEY_CURRENT_USER\Software\Microsoft\VSTO\Security\Inclusion" /v "file:///C:/Users/naray/Documents/Projects/ExcelClaude/ExcelPerplexityVSTO/ExcelPerplexityVSTO/bin/Debug/ExcelPerplexityVSTO.vsto" /f 2>nul
reg delete "HKEY_CURRENT_USER\Software\Microsoft\VSTO\Security\Inclusion" /v "file:///C:/Users/naray/Documents/Projects/ExcelClaude/ExcelPerplexityVSTO/ExcelPerplexityVSTO/bin/Release/ExcelPerplexityVSTO.vsto" /f 2>nul

echo.
echo Step 4: Cleaning project output...
if exist "C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO\ExcelPerplexityVSTO\bin\Debug" (
    echo Cleaning Debug folder...
    del /q "C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO\ExcelPerplexityVSTO\bin\Debug\*.*" 2>nul
)

if exist "C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO\ExcelPerplexityVSTO\obj\Debug" (
    echo Cleaning obj\Debug folder...
    rd /s /q "C:\Users\naray\Documents\Projects\ExcelClaude\ExcelPerplexityVSTO\ExcelPerplexityVSTO\obj\Debug" 2>nul
)

echo.
echo ========================================
echo Cleanup Complete!
echo ========================================
echo.
echo Next steps:
echo 1. Open Visual Studio
echo 2. Build -^> Clean Solution
echo 3. Build -^> Rebuild Solution
echo 4. Press F5 to debug
echo.
echo If you still have issues, restart your computer.
echo.
pause
