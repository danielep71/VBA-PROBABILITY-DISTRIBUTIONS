@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Sync-PROBDIST-VBA.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

rem Without a dropped workbook, a fresh development workbook is built from
rem the template in the ignored test-results folder, replaced on every run.
rem A dropped workbook is synchronized as it is and never rebuilt.
set "WORKBOOK=%~1"
set "BUILD="
if not defined WORKBOOK (
    set "WORKBOOK=%CD%\test-results\Probability-Distributions-Development.xlsm"
    set "BUILD=-Mode Development -FromTemplate"
)

echo Probability Distributions VBA synchronization
echo.

if not defined BUILD if not exist "%WORKBOOK%" (
    echo.
    echo ERROR: Workbook not found:
    echo   %WORKBOOK%
    echo.
    pause
    exit /b 1
)

echo.
echo Repository:
echo   %CD%
echo Workbook:
echo   %WORKBOOK%
echo.
if defined BUILD (
    echo Building a development workbook from src\workbook\PROBDIST_Template.xlsx
    echo and synchronizing it, ThisWorkbook code included...
) else (
    echo Using the dropped workbook as it is: its sheets are not rebuilt.
    echo Synchronizing VBA source...
)
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkbookPath "%WORKBOOK%" -RepoRoot "%REPO_ROOT%" %BUILD%
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: Probability Distributions VBA synchronization completed.
    if defined BUILD (
        echo Open the workbook with macros enabled and run Debug - Compile VBAProject:
        echo   %WORKBOOK%
    )
) else if "%RC%"=="2" (
    echo NOT CHANGED: nothing was imported; the workbook was not changed.
) else (
    echo FAILED: Probability Distributions VBA synchronization returned exit code %RC%.
)
echo.
pause
exit /b %RC%
