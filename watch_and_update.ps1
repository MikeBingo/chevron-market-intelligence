# ─────────────────────────────────────────────────────────────────────────────
# watch_and_update.ps1
# Runs silently in the background as a Windows startup task.
# Watches the folder for any new daily FCV xlsx file.
# When one appears, waits 30s (OneDrive settle), then fires update_dashboard.bat
# ─────────────────────────────────────────────────────────────────────────────

$FOLDER  = "C:\Users\MikeBingo\Chevron Leaf Tobacco\Power BI Uploads - 2026 TIMB FCV Daily Files"
$BAT     = "$FOLDER\update_dashboard.bat"
$LOG     = "$FOLDER\market_data_refresh.log"
$PATTERN = "2026 daily fcv summary day *.xlsx"

function Write-Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LOG -Value "[$ts] [WATCHER] $msg"
}

Write-Log "FileSystemWatcher started. Watching: $FOLDER"

$watcher                     = New-Object System.IO.FileSystemWatcher
$watcher.Path                = $FOLDER
$watcher.Filter              = "2026 daily fcv summary day *.xlsx"
$watcher.NotifyFilter        = [System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::LastWrite
$watcher.EnableRaisingEvents = $true

$action = {
    $file = $Event.SourceEventArgs.FullPath
    $name = $Event.SourceEventArgs.Name

    # Ignore temp/lock files
    if ($name -like "~$*" -or $name -like "*.tmp") { return }

    Add-Content -Path $LOG -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [WATCHER] New file detected: $name"
    Add-Content -Path $LOG -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [WATCHER] Waiting 30s for OneDrive to settle..."

    Start-Sleep -Seconds 30

    Add-Content -Path $LOG -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [WATCHER] Firing update_dashboard.bat..."
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$BAT`"" -WorkingDirectory $FOLDER -Wait
    Add-Content -Path $LOG -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [WATCHER] Done."
}

Register-ObjectEvent $watcher "Created" -Action $action | Out-Null
Register-ObjectEvent $watcher "Changed" -Action $action | Out-Null

Write-Log "Watcher active. Waiting for new files..."

# Keep the script alive indefinitely
while ($true) { Start-Sleep -Seconds 60 }
