$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

if (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 manage.py send_daily_activity_digest
    exit $LASTEXITCODE
}

if (Get-Command python -ErrorAction SilentlyContinue) {
    & python manage.py send_daily_activity_digest
    exit $LASTEXITCODE
}

throw 'Python launcher not found. Install Python or update scripts/run_daily_activity_digest.ps1.'
