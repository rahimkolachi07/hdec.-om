param(
    [string]$Time = '09:00'
)

$ErrorActionPreference = 'Stop'

$taskName = 'HDEC Daily Activity Digest'
$runScript = Join-Path $PSScriptRoot 'run_daily_activity_digest.ps1'
$taskCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$runScript`""

schtasks /Create /SC DAILY /TN $taskName /TR $taskCommand /ST $Time /F | Out-Host
Write-Host "Scheduled task '$taskName' created for $Time."
