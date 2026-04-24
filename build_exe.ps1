$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$python = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    throw "Missing virtual environment: .venv\Scripts\python.exe"
}

& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onedir `
    --name QA_Manager `
    main.py

$distDir = Join-Path $root "dist\QA_Manager"
if (Test-Path -LiteralPath $distDir) {
    $readme = Get-ChildItem -LiteralPath $root -Filter "README_*.md" -File | Select-Object -First 1
    if ($readme) {
        Copy-Item -LiteralPath $readme.FullName -Destination (Join-Path $distDir $readme.Name) -Force
    }
}

Write-Host ""
Write-Host "Build complete: dist\QA_Manager\QA_Manager.exe"
