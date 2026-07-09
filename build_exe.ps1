# Build the DS XML Lineage Extractor into a single standalone Windows .exe.
#
# Run from the project root:
#     .\build_exe.ps1
#
# Output: dist\DS_Lineage_Extractor.exe  (share this single file with users)

$ErrorActionPreference = "Stop"

$python = ".\.venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    Write-Host "Virtual env not found at $python — falling back to system python." -ForegroundColor Yellow
    $python = "python"
}

Write-Host "Ensuring PyInstaller is installed..." -ForegroundColor Cyan
& $python -m pip install --quiet pyinstaller

Write-Host "Building DS_Lineage_Extractor.exe (this takes a few minutes)..." -ForegroundColor Cyan
& $python -m PyInstaller DS_Lineage_Extractor.spec --noconfirm --clean

$exe = "dist\DS_Lineage_Extractor.exe"
if (Test-Path $exe) {
    $sizeMb = [math]::Round((Get-Item $exe).Length / 1MB, 1)
    Write-Host "`nBuild succeeded: $exe ($sizeMb MB)" -ForegroundColor Green
    Write-Host "Share this single file with users — they just double-click it." -ForegroundColor Green
} else {
    Write-Host "`nBuild finished but $exe was not found. Check the PyInstaller output above." -ForegroundColor Red
    exit 1
}