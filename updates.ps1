<# 
.SYNOPSIS
    Searches, downloads, and installs available Windows driver updates via Windows Update.

.DESCRIPTION
    Uses Microsoft.Update COM objects to:
        1) Search for non-installed driver updates
        2) Download them
        3) Install them
    Prints per-update results, summarizes outcome, and indicates if a reboot is required.
    Supports optional transcript logging and a dry run (search only).

.NOTES
    Requires: PowerShell 5+, Windows Update service enabled, Internet.
    Run as Administrator for best results.

.PARAMETER DryRun
    Performs search only. No download or install.

.PARAMETER TranscriptPath
    If provided, starts a transcript (log) to this file.

.EXAMPLE
    .\Update-Drivers.ps1

.EXAMPLE
    .\Update-Drivers.ps1 -DryRun

.EXAMPLE
    .\Update-Drivers.ps1 -TranscriptPath "C:\Logs\DriverUpdate_$(Get-Date -f yyyyMMdd_HHmmss).log"
#>


[CmdletBinding()]
param(
    [switch]$DryRun,
    [string]$TranscriptPath
)

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-Section($text) {
    Write-Host ("`n=== {0} ===" -f $text) -ForegroundColor Cyan
}

function Get-ResultText([int]$code) {
    switch ($code) {
        0 { "NotStarted" }
        1 { "InProgress" }
        2 { "Succeeded" }
        3 { "SucceededWithErrors" }
        4 { "Failed" }
        5 { "Aborted" }
        default { "Unknown($code)" }
    }
}

# -------------------- Start --------------------
if ($TranscriptPath) {
    try {
        $parent = Split-Path -Parent $TranscriptPath
        if ($parent -and -not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent | Out-Null }
        Start-Transcript -Path $TranscriptPath -Append -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "Could not start transcript: $($_.Exception.Message)"
    }
}

try {
    if (-not (Test-Admin)) {
        Write-Warning "It is recommended to run this script as Administrator for driver installs."
    }

    Write-Section "Creating Windows Update session"
    $UpdateSession  = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

    Write-Section "Searching for available driver updates"
    # Filter: only drivers not installed yet
    $criteria = "IsInstalled=0 and Type='Driver'"
    $SearchResult = $UpdateSearcher.Search($criteria)

    Write-Host ("Found {0} driver update(s)." -f $SearchResult.Updates.Count) -ForegroundColor Green

    if ($SearchResult.Updates.Count -eq 0) {
        Write-Host "No driver updates available. Exiting." -ForegroundColor Yellow
        return
    }

    # Print list of drivers to be installed
    for ($i = 0; $i -lt $SearchResult.Updates.Count; $i++) {
        $u = $SearchResult.Updates.Item($i)
        Write-Host ("[{0}] {1}" -f $i, $u.Title)
    }

    if ($DryRun) {
        Write-Host "`nDryRun specified. Skipping download and install." -ForegroundColor Yellow
        return
    }

    Write-Section "Downloading driver updates"
    $UpdatesToProcess = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($u in $SearchResult.Updates) {
        [void]$UpdatesToProcess.Add($u)
    }

    $Downloader = $UpdateSession.CreateUpdateDownloader()
    $Downloader.Updates = $UpdatesToProcess
    $DownloadResult = $Downloader.Download()

    # Optional: check download result codes per update
    for ($i = 0; $i -lt $UpdatesToProcess.Count; $i++) {
        $u = $UpdatesToProcess.Item($i)
        $dr = $u.IsDownloaded
        Write-Host ("Downloaded: {0} -> {1}" -f $u.Title, $(if ($dr) { "Yes" } else { "No" }))
    }

    Write-Section "Installing driver updates"
    $Installer = $UpdateSession.CreateUpdateInstaller()
    $Installer.Updates = $UpdatesToProcess
    $InstallationResult = $Installer.Install()

    $overall = Get-ResultText -code $InstallationResult.ResultCode
    Write-Host ("Overall installation result: {0}" -f $overall) -ForegroundColor Green

    # Per-update installation results
    Write-Section "Per-update results"
    for ($i = 0; $i -lt $UpdatesToProcess.Count; $i++) {
        $u  = $UpdatesToProcess.Item($i)
        $ur = $InstallationResult.GetUpdateResult($i)
        $txt = Get-ResultText -code $ur.ResultCode
        Write-Host ("[{0}] {1} -> {2}" -f $i, $u.Title, $txt)
        if ($ur.HResult -ne 0) {
            Write-Warning ("    HResult: {0}" -f ("0x{0:X8}" -f $ur.HResult))
        }
        if ($ur.RebootRequired) {
            Write-Host "    RebootRequired: True" -ForegroundColor Yellow
        }
    }

    # Reboot summary
    if ($InstallationResult.RebootRequired) {
        Write-Host "`nA system restart is required to complete some driver installations." -ForegroundColor Yellow
    } else {
        Write-Host "`nNo restart required." -ForegroundColor Green
    }

} catch {
    Write-Error ("Unhandled error: {0}" -f $_.Exception.Message)
} finally {
    # Clean up COM objects
    foreach ($obj in @($InstallationResult, $Installer, $Downloader, $UpdatesToProcess, $UpdateSearcher, $UpdateSession)) {
        if ($obj -and ($obj -is [__ComObject])) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
        }
    }
    if ($TranscriptPath) {
        try { Stop-Transcript | Out-Null } catch {}
    }
}
