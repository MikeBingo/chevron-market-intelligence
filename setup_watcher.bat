@echo off
:: ─────────────────────────────────────────────────────────────────────────────
:: setup_watcher.bat
:: Run ONCE (as Administrator) to register the file watcher as a startup task.
:: After this the watcher starts automatically every time Windows boots.
:: ─────────────────────────────────────────────────────────────────────────────

set TASK_NAME=ChevronDashboardWatcher
set PS_SCRIPT=C:\Users\MikeBingo\Chevron Leaf Tobacco\Power BI Uploads - 2026 TIMB FCV Daily Files\watch_and_update.ps1

echo Registering watcher task: %TASK_NAME%

schtasks /create ^
  /tn "%TASK_NAME%" ^
  /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File \"%PS_SCRIPT%\"" ^
  /sc onlogon ^
  /ru "%USERNAME%" ^
  /rl HIGHEST ^
  /f

if %errorlevel% equ 0 (
    echo.
    echo SUCCESS — watcher registered.
    echo It will start automatically every time you log in to Windows.
    echo Starting it now...
    schtasks /run /tn "%TASK_NAME%"
    echo Watcher is running in the background.
) else (
    echo.
    echo ERROR — run this bat as Administrator and try again.
)

pause
