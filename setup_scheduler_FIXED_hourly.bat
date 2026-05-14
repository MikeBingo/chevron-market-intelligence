@echo off
setlocal EnableExtensions

:: ─────────────────────────────────────────────────────────────────────────────
:: setup_scheduler_FIXED_hourly.bat
:: Registers a resilient polling task.
:: It runs update_dashboard.bat every 30 minutes while the user is logged in.
:: The update script itself decides whether there is new data.
:: This is safer than depending only on a permanently running watcher.
:: ─────────────────────────────────────────────────────────────────────────────

set "TASK_NAME=ChevronDashboardUpdatePolling"
set "BAT_FILE=%~dp0update_dashboard.bat"

chcp 65001 >nul 2>&1

echo Registering scheduled task: %TASK_NAME%
echo Bat file: %BAT_FILE%
echo Frequency: every 30 minutes

schtasks /create ^
  /tn "%TASK_NAME%" ^
  /tr "cmd /c \"%BAT_FILE%\"" ^
  /sc minute ^
  /mo 30 ^
  /ru "%USERNAME%" ^
  /rl HIGHEST ^
  /f

if %errorlevel% equ 0 (
    echo.
    echo SUCCESS - polling task registered.
    echo It will run every 30 minutes and update/push only when there is new data.
    echo.
    echo Starting it now...
    schtasks /run /tn "%TASK_NAME%"
) else (
    echo.
    echo ERROR - could not register task. Try running this file as Administrator.
)

pause
