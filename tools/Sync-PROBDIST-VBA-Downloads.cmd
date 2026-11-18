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

set "WORKBOOK=%~1"

if not defined WORKBOOK (
    echo Probability Distributions VBA synchronization from a folder
    echo.
    echo No workbook given. Drag the .xlsm workbook onto this file, or run:
    echo   %~nx0 ^<path to workbook^>
    echo.
    pause
    exit /b 1
)

set "DEFAULT_SOURCE=%USERPROFILE%\Downloads"

echo Probability Distributions VBA synchronization from a folder
echo.
echo Imports the repository modules found in the folder, for example modules
echo downloaded for a test, and new modules named M_STATS_PROBDIST_*. The
echo repository files are not changed.
echo.

if not exist "%WORKBOOK%" (
    echo ERROR: Workbook not found:
    echo   %WORKBOOK%
    echo.
    pause
    exit /b 1
)

set "SOURCE="
set /p "SOURCE=Folder with the modules [%DEFAULT_SOURCE%]: "
if not defined SOURCE set "SOURCE=%DEFAULT_SOURCE%"
set "SOURCE=%SOURCE:"=%"

if not exist "%SOURCE%\" (
    echo.
    echo ERROR: Folder not found:
    echo   %SOURCE%
    echo.
    pause
    exit /b 1
)

echo.
echo Workbook:
echo   %WORKBOOK%
echo Folder:
echo   %SOURCE%
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkbookPath "%WORKBOOK%" -RepoRoot "%REPO_ROOT%" -SourceDir "%SOURCE%"
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: modules imported from the folder.
) else if "%RC%"=="2" (
    echo CANCELLED: the workbook was not changed.
) else (
    echo FAILED: Probability Distributions VBA synchronization returned exit code %RC%.
)
echo.
pause
exit /b %RC%
