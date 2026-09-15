$ErrorActionPreference = "Stop"

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Error "Python tidak ditemukan. Instal Python 3.11 atau lebih baru terlebih dahulu."
}

if (-not (Test-Path ".venv")) {
    python -m venv .venv
}

& ".\.venv\Scripts\python.exe" -m pip install --upgrade pip
& ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt

if (-not (Test-Path ".env")) {
    Copy-Item ".env.example" ".env"
    Write-Host "File .env dibuat. Silakan periksa SCREENSHOT_FOLDER sebelum menjalankan program."
}

New-Item -ItemType Directory -Force -Path "reports", "logs", "data" | Out-Null
Write-Host "Instalasi selesai. Jalankan .\run.ps1"

