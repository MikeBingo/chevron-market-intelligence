# ─────────────────────────────────────────────────────────────────────────────
# watch_and_update.ps1 — FIXED v2
# Watches for new TIMB daily files and runs update_dashboard.bat.
# Fixes:
#   1. Forces UTF-8 output to avoid Windows charmap issues.
#   2. Prevents overlapping runs.
#   3. Updates .watcher_last_day.txt ONLY when update_dashboard.bat succeeds.
#   4. Logs failed runs clearly so the same day can retry later.
# ─────────────────────────────────────────────────────────────────────────────

$FOLDER       = "C:\Users\MikeBingo\Chevron Leaf Tobacco\Power BI Uploads - 2026 TIMB FCV Daily Files"
$BAT          = Join-Path $FOLDER "update_dashboard.bat"
$LOG          = Join-Path $FOLDER "market_data_refresh.log"
$STATE_FILE   = Join-Path $FOLDER ".watcher_last_day.txt"
$LOCK_FILE    = Join-Path $FOLDER ".watcher_running.lock"
$POLL_SECONDS = 300

# Force UTF-8 for PowerShell output and child processes.
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

function Write-Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LOG -Value "[$ts] [WATCHER] $msg" -Encoding UTF8
}

function Get-LatestDay {
    $files = Get-ChildItem -Path $FOLDER -Filter "2026 daily fcv summary day *.xlsx" -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notlike "~`$*" -and $_.Name -notlike "*.tmp" }

    if (-not $files) { return 0 }

    $days = @()
    foreach ($file in $files) {
        if ($file.BaseName -match 'day\s*(\d+)$') {
            $days += [int]$Matches[1]
        }
    }

    if ($days.Count -eq 0) { return 0 }
    return ($days | Measure-Object -Maximum).Maximum
}

function Get-LastProcessedDay {
    if (Test-Path $STATE_FILE) {
        $val = (Get-Content $STATE_FILE -ErrorAction SilentlyContinue | Select-Object -First 1).Trim()
        if ($val -match '^\d+$') { return [int]$val }
    }
    return 0
}

function Set-LastProcessedDay($day) {
    Set-Content -Path $STATE_FILE -Value $day -Encoding UTF8
}

function Run-Update($reason) {
    if (Test-Path $LOCK_FILE) {
        Write-Log "Skipped: another update is already running. Reason was: $reason"
        return
    }

    New-Item $LOCK_FILE -ItemType File -Force | Out-Null

    try {
        $latestBefore = Get-LatestDay
        Write-Log "$reason — running update_dashboard.bat for Day $latestBefore..."

        # Allow OneDrive/SharePoint to finish writing the incoming file.
        Start-Sleep -Seconds 30

        $process = Start-Process -FilePath "cmd.exe" `
                                 -ArgumentList "/c `"$BAT`"" `
                                 -WorkingDirectory $FOLDER `
                                 -Wait `
                                 -PassThru

        $exitCode = $process.ExitCode
        $latestAfter = Get-LatestDay

        if ($exitCode -eq 0) {
            Set-LastProcessedDay $latestAfter
            Write-Log "SUCCESS. Last processed day set to: $latestAfter"
        } else {
            Write-Log "FAILED. update_dashboard.bat exit code: $exitCode. Last processed day NOT changed. It will retry on next poll."
        }
    }
    catch {
        Write-Log "FAILED with exception: $($_.Exception.Message). Last processed day NOT changed."
    }
    finally {
        Remove-Item $LOCK_FILE -Force -ErrorAction SilentlyContinue
    }
}

Write-Log "FileSystemWatcher started. Poll interval: $POLL_SECONDS seconds."

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $FOLDER
$watcher.Filter = "2026 daily fcv summary day *.xlsx"
$watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::LastWrite
$watcher.EnableRaisingEvents = $true

$action = {
    $name = $Event.SourceEventArgs.Name
    if ($name -like "~`$*" -or $name -like "*.tmp") { return }

    $latestDay = Get-LatestDay
    $lastDay = Get-LastProcessedDay

    if ($latestDay -gt $lastDay) {
        Run-Update "FileSystemWatcher: Day $latestDay detected, previous processed day $lastDay"
    }
}

Register-ObjectEvent $watcher "Created" -Action $action | Out-Null
Register-ObjectEvent $watcher "Changed" -Action $action | Out-Null

Write-Log "Watcher active."

while ($true) {
    Start-Sleep -Seconds $POLL_SECONDS

    $latestDay = Get-LatestDay
    $lastDay = Get-LastProcessedDay

    if ($latestDay -gt $lastDay) {
        Run-Update "Polling: Day $latestDay detected, previous processed day $lastDay"
    }
}
