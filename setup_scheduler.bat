@echo off
:: ─────────────────────────────────────────────────────────────────────────────
:: setup_scheduler.bat
:: Run this ONCE to register the daily auto-update with Windows Task Scheduler.
:: After this you never need to touch anything — it runs at 2:30 PM every day.
:: ─────────────────────────────────────────────────────────────────────────────

set TASK_NAME=ChevronDashboardUpdate
set BAT_FILE=%~dp0update_dashboard.bat
set RUN_TIME=14:30

echo Registering scheduled task: %TASK_NAME%
echo Bat file : %BAT_FILE%
echo Run time : %RUN_TIME% daily

schtasks /create ^
  /tn "%TASK_NAME%" ^
  /tr "cmd /c \"%BAT_FILE%\"" ^
  /sc daily ^
  /st %RUN_TIME% ^
  /ru "%USERNAME%" ^
  /rl HIGHEST ^
  /f

if %errorlevel% equ 0 (
    echo.
    echo SUCCESS — task registered. Dashboard will auto-update at %RUN_TIME% every day.
    echo To change the time: open Task Scheduler, find "%TASK_NAME%", edit the trigger.
    echo To remove it: run  schtasks /delete /tn "%TASK_NAME%" /f
) else (
    echo.
    echo ERROR — could not register task. Try running this bat as Administrator.
)

pause
