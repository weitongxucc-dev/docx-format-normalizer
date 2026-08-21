# setup.ps1 - Windows environment setup for docx-format-normalizer
# Usage (PowerShell):  powershell -ExecutionPolicy Bypass -File scripts\setup.ps1
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SkillDir  = Split-Path -Parent $ScriptDir
$VenvDir   = Join-Path $SkillDir ".venv"
$PyExe     = Join-Path $VenvDir "Scripts\python.exe"

Write-Host "=== DOCX Format Normalizer - Windows Setup ==="

# Find a suitable Python
$python = $null
foreach ($cmd in @("python3.12", "python3.11", "python3", "python", "py")) {
    if (Get-Command $cmd -ErrorAction SilentlyContinue) { $python = $cmd; break }
}
if (-not $python) {
    Write-Host "ERROR: No Python found. Install Python 3.8+ from https://python.org"
    exit 1
}
Write-Host "[INFO] Using Python: $(& $python --version 2>&1)"

# Reuse existing venv if it works
if (Test-Path $PyExe) {
    & $PyExe -c "import docx" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[INFO] Existing venv is working."
        Write-Host "SUCCESS: Python path = $PyExe"
        exit 0
    } else {
        Write-Host "[INFO] Removing broken venv..."
        Remove-Item -Recurse -Force $VenvDir
    }
}

# Create virtual environment
Write-Host "[INFO] Creating venv..."
if ($python -eq "py") { & py -3 -m venv $VenvDir } else { & $python -m venv $VenvDir }
& $PyExe -m pip install --quiet --upgrade pip
& $PyExe -m pip install --quiet "python-docx>=0.8.11" "lxml>=4.9,<6"

# Verify
& $PyExe -c "from docx import Document; from docx.shared import Pt, Mm; print('OK: python-docx ready')"
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: python-docx installation failed."
    Write-Host "Try manually: $PyExe -m pip install 'python-docx>=0.8.11' 'lxml>=4.9,<6'"
    exit 1
}

Write-Host ""
Write-Host "SUCCESS: Environment ready."
Write-Host "  Python path: $PyExe"
Write-Host "  Use this path to run docx_formatter.py"
