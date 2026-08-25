# Builds a self-contained, single-file Windows release for Invigoration.
#
# Usage:
#   .\build-windows.ps1 [-Rid win-x64]
#
# Single-file (not PublishSingleFile=false, unlike build-macos.sh/build-linux.sh — those
# need loose files for .app-bundle/tarball packaging; Windows users expect one .exe).
# IncludeNativeLibrariesForSelfExtract bundles the SC2 native library (stimpak.dll) into the
# exe too, extracted to a per-user temp cache at first run — confirmed working via a live
# smoke test, not just reading the docs.

param(
    [string]$Rid = "win-x64"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

$VersionMatch = Select-String -Path "src\Invigoration.Core\AppVersion.cs" -Pattern '"([0-9][^"]*)"'
$Version = $VersionMatch.Matches[0].Groups[1].Value

$DistDir = Join-Path $ScriptDir "dist\$Rid"
$PackageDir = Join-Path $DistDir "Invigoration-v$Version-$Rid"
$ZipPath = Join-Path $DistDir "Invigoration-v$Version-$Rid.zip"

Write-Host "==> Publishing self-contained single-file $Rid build (version $Version)"
if (Test-Path $DistDir) { Remove-Item -Recurse -Force $DistDir }
dotnet publish src\Invigoration.App\Invigoration.App.csproj `
    -c Release -r $Rid --self-contained true `
    -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
    -o $PackageDir
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed" }

Write-Host "==> Stripping debug symbols (not needed for an end-user distribution, and roughly double the download otherwise)"
Get-ChildItem -Path $PackageDir -Filter "*.pdb" | Remove-Item -Force

Write-Host "==> Zipping"
Compress-Archive -Path $PackageDir -DestinationPath $ZipPath -Force

Write-Host "==> Done: $ZipPath"
