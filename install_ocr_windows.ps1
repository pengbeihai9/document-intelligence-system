$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Checker = Join-Path $ProjectRoot "tools\check_ocr.py"
$DefaultTesseract = "C:\Program Files\Tesseract-OCR\tesseract.exe"
$PortableTesseract = Join-Path $ProjectRoot "tools\Tesseract-OCR\tesseract.exe"

function Test-Ocr {
    if (Test-Path $Checker) {
        python $Checker
    } else {
        Write-Host "OCR checker not found: $Checker"
    }
}

Write-Host "Checking existing OCR environment..."
Test-Ocr
if ($LASTEXITCODE -eq 0) {
    Write-Host "OCR is already ready."
    exit 0
}

if (Test-Path $PortableTesseract) {
    Write-Host "Portable Tesseract found: $PortableTesseract"
    Test-Ocr
    exit $LASTEXITCODE
}

if (Get-Command winget -ErrorAction SilentlyContinue) {
    Write-Host "Trying winget install: UB-Mannheim.TesseractOCR"
    try {
        winget install --id UB-Mannheim.TesseractOCR -e --accept-package-agreements --accept-source-agreements
    } catch {
        Write-Host "winget install failed: $($_.Exception.Message)"
    }
} else {
    Write-Host "winget was not found."
}

if (Test-Path $DefaultTesseract) {
    Write-Host "Installed: $DefaultTesseract"
    Test-Ocr
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "Tesseract is still not ready."
Write-Host "Offline options:"
Write-Host "1. Install UB-Mannheim Tesseract OCR manually."
Write-Host "2. Or copy a portable Tesseract folder into:"
Write-Host "   tools\Tesseract-OCR\"
Write-Host ""
Write-Host "Required files:"
Write-Host "   tools\Tesseract-OCR\tesseract.exe"
Write-Host "   tools\Tesseract-OCR\tessdata\chi_sim.traineddata"
Write-Host "   tools\Tesseract-OCR\tessdata\eng.traineddata"
Write-Host ""
Write-Host "After installing or copying files, run:"
Write-Host "   python tools\check_ocr.py"
exit 1
