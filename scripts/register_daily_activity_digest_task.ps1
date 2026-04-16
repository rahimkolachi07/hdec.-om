param(
    [string]$Time = '08:00'
)

$ErrorActionPreference = 'Stop'

$taskName = 'HDEC Daily Activity Digest'
$runScript = Join-Path $PSScriptRoot 'run_daily_activity_digest.ps1'
$wrapperPath = Join-Path $env:TEMP 'hdec_daily_activity_digest.cmd'
$wrapperContent = @"
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$runScript"
"@
Set-Content -Path $wrapperPath -Value $wrapperContent -Encoding Ascii

$taskCommand = ('cmd.exe /c "{0}"' -f $wrapperPath)

& schtasks.exe /Create /SC DAILY /TN $taskName /TR $taskCommand /ST $Time /F | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "Failed to create scheduled task '$taskName'."
}
Write-Host "Scheduled task '$taskName' created for $Time."
