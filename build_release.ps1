param(
    [string]$AppName = "CaptionedLiveCapture"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$distDir = Join-Path $root "dist"
$buildDir = Join-Path $root "build"
$releaseDir = Join-Path $root "release"
$appDistDir = Join-Path $distDir $AppName
$zipPath = Join-Path $releaseDir "$AppName-portable.zip"

Set-Location $root

function Get-PythonCommand {
    foreach ($candidate in @("py", "python3", "python")) {
        if (Get-Command $candidate -ErrorAction SilentlyContinue) {
            return $candidate
        }
    }
    throw "Python was not found in PATH. Install Python and try again."
}

$pythonCmd = Get-PythonCommand

Write-Host "Installing build requirements..."
$env:PYTHONIOENCODING = "utf-8"
& $pythonCmd -m pip install -r requirements.txt pyinstaller

Write-Host "Cleaning old build outputs..."
if (Test-Path $buildDir) { Remove-Item -LiteralPath $buildDir -Recurse -Force }
if (Test-Path $distDir) { Remove-Item -LiteralPath $distDir -Recurse -Force }
if (!(Test-Path $releaseDir)) { New-Item -ItemType Directory -Path $releaseDir | Out-Null }
if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }

Write-Host "Building executable..."
& $pythonCmd -m PyInstaller --noconfirm --clean --windowed --name $AppName capture_text_app.py

if (!(Test-Path $appDistDir)) {
    throw "Build output was not created: $appDistDir"
}

Write-Host "Creating portable zip..."
if (!(Test-Path $releaseDir)) { New-Item -ItemType Directory -Path $releaseDir | Out-Null }
if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path (Join-Path $appDistDir "*") -DestinationPath $zipPath -Force

Write-Host "Release created:"
Write-Host $zipPath
