param([string]$Python = "python")
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$localPython = Join-Path $projectRoot ".venv/Scripts/python.exe"
if (!(Test-Path -LiteralPath $localPython)) {
    & $Python -m venv (Join-Path $projectRoot ".venv")
    if ($LASTEXITCODE -ne 0) { throw "Creating the Python environment failed. Use Python 3.11 or 3.12." }
}
& $localPython -m pip install -r (Join-Path $PSScriptRoot "requirements-lock.txt")
if ($LASTEXITCODE -ne 0) { throw "Installing OCR dependencies failed." }
Push-Location $projectRoot
try {
    & $localPython -m pipeline.setup_models
    if ($LASTEXITCODE -ne 0) { throw "Downloading the English OCR model failed." }
    & $localPython -c "from pipeline.ocr import RapidOcrProvider; print(RapidOcrProvider().metadata)"
    if ($LASTEXITCODE -ne 0) { throw "OCR initialization failed." }
} finally {
    Pop-Location
}
