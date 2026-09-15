$ErrorActionPreference = "Stop"

if (Test-Path "C:\Program Files\Tesseract-OCR\tesseract.exe") {
    Write-Host "Tesseract sudah terpasang."
    exit 0
}

if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Error "winget tidak tersedia. Instal Tesseract OCR secara manual, lalu atur TESSERACT_COMMAND di .env."
}

winget install --id UB-Mannheim.TesseractOCR --exact --accept-package-agreements --accept-source-agreements

if (Test-Path "C:\Program Files\Tesseract-OCR\tesseract.exe") {
    Write-Host "Tesseract berhasil dipasang. Tutup lalu buka kembali PowerShell."
} else {
    Write-Warning "Tesseract belum ditemukan pada lokasi standar. Periksa TESSERACT_COMMAND di .env."
}

