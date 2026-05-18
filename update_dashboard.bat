@echo off
setlocal EnableExtensions EnableDelayedExpansion

:: ─────────────────────────────────────────────────────────────────────────────
:: update_dashboard_FIXED.bat
:: Runs market data refresh + dashboard injection + GitHub Pages publish.
:: Fixes Windows Unicode/charmap failures by forcing UTF-8 at every layer.
:: ─────────────────────────────────────────────────────────────────────────────

set "ROOT=%~dp0"
set "CLEAN_SCRIPT=%ROOT%Scripts\run_2026_market_data.py"
set "P1_SCRIPT=%ROOT%Scripts\generate_dashboard_data.py"
set "P2_SCRIPT=%ROOT%Scripts\generate_competitor_data.py"
set "LOG=%ROOT%market_data_refresh.log"
set "REPO=%ROOT%"
set "DASHBOARD=%ROOT%Market Intelligence Dashboard 2026.html"
set "INDEX=%ROOT%index.html"

:: Force UTF-8 for console + Python. This prevents ✓, → and similar symbols crashing.
chcp 65001 >nul 2>&1
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "LANG=en_US.UTF-8"
set "LC_ALL=en_US.UTF-8"

cd /d "%ROOT%"

echo [%date% %time%] Trigger fired >> "%LOG%"
echo STEP 1: folder = %ROOT%
echo STEP 2: log    = %LOG%

where python
if errorlevel 1 (
    echo ERROR: Python not found in PATH
    echo [%date% %time%] ERROR: Python not in PATH >> "%LOG%"
    pause
    exit /b 1
)

echo STEP 3: Python found. Running P0...
echo [%date% %time%] Running cleaning script... >> "%LOG%"
python -X utf8 "%CLEAN_SCRIPT%"
if errorlevel 1 (
    echo ERROR: P0 cleaning script failed ^(see above^)
    echo [%date% %time%] ERROR: Cleaning script failed. >> "%LOG%"
    pause
    exit /b 1
)
echo [%date% %time%] Data cleaned. >> "%LOG%"

echo [%date% %time%] Running P1/P0 dashboard injection... >> "%LOG%"
python -X utf8 "%P1_SCRIPT%" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: P1/P0 script failed. >> "%LOG%"
    exit /b 1
)
echo [%date% %time%] P1/P0 data injected. >> "%LOG%"

:: Allow OneDrive to release the HTML before P2 writes to it.
timeout /t 10 /nobreak >nul

echo [%date% %time%] Running P2 competitor injection... >> "%LOG%"
python -X utf8 "%P2_SCRIPT%" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: P2 script failed. >> "%LOG%"
    exit /b 1
)
echo [%date% %time%] P2 data injected. >> "%LOG%"

if not exist "%DASHBOARD%" (
    echo [%date% %time%] ERROR: Dashboard HTML not found: %DASHBOARD% >> "%LOG%"
    exit /b 1
)

copy /Y "%DASHBOARD%" "%INDEX%" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: Failed to copy dashboard to index.html. >> "%LOG%"
    exit /b 1
)

:: Confirm this folder is a Git repo before trying to publish.
git -C "%REPO%" rev-parse --is-inside-work-tree >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [%date% %time%] WARNING: Not a Git repo. Dashboard updated locally only. >> "%LOG%"
    exit /b 0
)

git -C "%REPO%" add "Market Intelligence Dashboard 2026.html" "index.html" >> "%LOG%" 2>&1

git -C "%REPO%" diff --cached --quiet
if not errorlevel 1 (
    echo [%date% %time%] No dashboard changes to commit. >> "%LOG%"
    echo [%date% %time%] Done >> "%LOG%"
    exit /b 0
)

git -C "%REPO%" commit -m "Dashboard auto-update %date% %time%" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [%date% %time%] WARNING: git commit failed. >> "%LOG%"
    exit /b 1
)

git -C "%REPO%" push >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [%date% %time%] WARNING: git push failed. Local dashboard updated but GitHub not published. >> "%LOG%"
    exit /b 1
)

echo [%date% %time%] GitHub Pages published successfully. >> "%LOG%"
echo [%date% %time%] Done >> "%LOG%"
exit /b 0
