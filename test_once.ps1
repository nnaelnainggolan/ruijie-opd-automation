param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$ImagePath
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Error "Virtual environment belum tersedia. Jalankan .\install.ps1 terlebih dahulu."
}

& ".\.venv\Scripts\python.exe" ".\main.py" --once "$ImagePath"

