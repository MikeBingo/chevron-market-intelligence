@echo off
:: ─────────────────────────────────────────────────────────────────────────────
:: update_dashboard.bat
:: Triggered by Task Scheduler daily at 2:30 PM.
:: Runs the full market data refresh + dashboard injection pipeline,
:: then publishes the updated dashboard to GitHub Pages.
:: ─────────────────────────────────────────────────────────────────────────────

set ROOT=%~dp0
set CLEAN_SCRIPT=%ROOT%Scripts\run_2026_market_data.py
set P1_SCRIPT=%ROOT%Scripts\generate_dashboard_data.py
set P2_SCRIPT=%ROOT%Scripts\generate_competitor_data.py
set LOG=%ROOT%market_data_refresh.log
set REPO=%ROOT%

:: Force UTF-8 so Python Unicode characters don't crash on Windows
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

echo [%date% %time%] Power Automate trigger fired >> "%LOG%"

:: ── 0. Clean raw daily xlsx → Cleaned Tobacco Market Data 2026.xlsx ──────────
python "%CLEAN_SCRIPT%"
if %errorlevel% neq 0 (
    echo [%date% %time%] ERROR: Cleaning script failed ^(exit %errorlevel%^) >> "%LOG%"
    exit /b %errorlevel%
)
echo [%date% %time%] Data cleaned >> "%LOG%"

:: ── 1a. Run P1 dashboard data (Season Overview + Regional Competitiveness) ───
python "%P1_SCRIPT%"
if %errorlevel% neq 0 (
    echo [%date% %time%] ERROR: P1 script failed ^(exit %errorlevel%^) >> "%LOG%"
    exit /b %errorlevel%
)
echo [%date% %time%] P1 data injected >> "%LOG%"

:: ── 1b. Run P2 dashboard data (Competitor Intelligence) ──────────────────────
python "%P2_SCRIPT%"
if %errorlevel% neq 0 (
    echo [%date% %time%] ERROR: P2 script failed ^(exit %errorlevel%^) >> "%LOG%"
    exit /b %errorlevel%
)
echo [%date% %time%] P2 data injected >> "%LOG%"

:: ── 2. Copy updated dashboard as index.html for GitHub Pages ─────────────────
copy /Y "%ROOT%Market Intelligence Dashboard 2026.html" "%ROOT%index.html" >> "%LOG%" 2>&1

:: ── 3. Push to GitHub Pages ──────────────────────────────────────────────────
git -C "%REPO%" add "Market Intelligence Dashboard 2026.html" index.html >> "%LOG%" 2>&1
git -C "%REPO%" commit -m "Dashboard auto-update — %date% %time%" >> "%LOG%" 2>&1
git -C "%REPO%" push >> "%LOG%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] WARNING: git push failed ^(exit %errorlevel%^) >> "%LOG%"
) else (
    echo [%date% %time%] GitHub Pages published successfully >> "%LOG%"
)

echo [%date% %time%] Done >> "%LOG%"
