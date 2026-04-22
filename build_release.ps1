param(
    [string]$AppName = "CaptionedLiveCapture",
    [string]$TesseractPath = "C:\Program Files\Tesseract-OCR"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$distDir = Join-Path $root "dist"
$buildDir = Join-Path $root "build"
$releaseDir = Join-Path $root "release"
$appDistDir = Join-Path $distDir $AppName
$zipPath = Join-Path $releaseDir "$AppName-portable.zip"

Set-Location $root

Write-Host "Installing build requirements..."
$env:PYTHONIOENCODING = "utf-8"
py -m pip install -r requirements.txt pyinstaller

Write-Host "Cleaning old build outputs..."
if (Test-Path $buildDir) { Remove-Item -LiteralPath $buildDir -Recurse -Force }
if (Test-Path $distDir) { Remove-Item -LiteralPath $distDir -Recurse -Force }
if (!(Test-Path $releaseDir)) { New-Item -ItemType Directory -Path $releaseDir | Out-Null }
if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }

Write-Host "Building executable..."
py -m PyInstaller --noconfirm --clean --windowed --name $AppName capture_text_app.py

if (!(Test-Path $appDistDir)) {
    throw "Build output was not created: $appDistDir"
}

if (!(Test-Path (Join-Path $TesseractPath "tesseract.exe"))) {
    throw "Tesseract was not found at: $TesseractPath"
}

Write-Host "Bundling Tesseract OCR..."
$bundledTesseract = Join-Path $appDistDir "tesseract"
if (Test-Path $bundledTesseract) { Remove-Item -LiteralPath $bundledTesseract -Recurse -Force }
Copy-Item -LiteralPath $TesseractPath -Destination $bundledTesseract -Recurse -Force

Write-Host "Creating portable zip..."
if (!(Test-Path $releaseDir)) { New-Item -ItemType Directory -Path $releaseDir | Out-Null }
if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path (Join-Path $appDistDir "*") -DestinationPath $zipPath -Force

Write-Host "Release created:"
Write-Host $zipPath
