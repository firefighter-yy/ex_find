$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

python -m pip install -e ".[dev]"
python -m pytest -q
python -m PyInstaller --noconfirm --clean excel-search.spec
python -m pip freeze | Set-Content -Encoding utf8 "dist\dependencies.txt"
Get-FileHash "dist\ExcelInformationSearch.exe" -Algorithm SHA256 |
    Format-List | Out-File -Encoding utf8 "dist\SHA256.txt"

Write-Host "Build complete: dist\ExcelInformationSearch.exe"
