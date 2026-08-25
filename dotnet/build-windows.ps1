# Builds a self-contained, single-file Windows release for Invigoration.
#
# Usage:
#   .\build-windows.ps1 [-Rid win-x64]
#
#   # Once you have a real code-signing certificate (see the "Code signing" section of the
#   # Setup Guide wiki page for why this can't just be automated/free): sign with a cert already
#   # imported into the Windows certificate store (Personal/My), identified by its thumbprint —
#   # this is the usual flow after a CA issues you a cert via their own enrollment tool:
#   .\build-windows.ps1 -CertThumbprint "AB12CD34..."
#
#   # ...or sign with a standalone .pfx file instead:
#   .\build-windows.ps1 -CertPfxPath "C:\path\to\cert.pfx" -CertPfxPassword (Read-Host -AsSecureString)
#
# Single-file (not PublishSingleFile=false, unlike build-macos.sh/build-linux.sh — those
# need loose files for .app-bundle/tarball packaging; Windows users expect one .exe).
# IncludeNativeLibrariesForSelfExtract bundles the SC2 native library (stimpak.dll) into the
# exe too, extracted to a per-user temp cache at first run — confirmed working via a live
# smoke test, not just reading the docs.

param(
    [string]$Rid = "win-x64",

    # Thumbprint of a code-signing cert already imported into the current user's certificate
    # store (Cert:\CurrentUser\My) — mutually exclusive with CertPfxPath/CertPfxPassword below.
    [string]$CertThumbprint = "",

    # Alternative to CertThumbprint: sign using a standalone .pfx file instead of the cert store.
    [string]$CertPfxPath = "",
    [SecureString]$CertPfxPassword = $null,

    # RFC3161 timestamp server — keeps the signature valid after the certificate itself expires.
    # DigiCert's is free and doesn't require owning a DigiCert certificate to use.
    [string]$TimestampUrl = "http://timestamp.digicert.com"
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

$ExePath = Join-Path $PackageDir "Invigoration.App.exe"
if ($CertThumbprint -or $CertPfxPath) {
    $SignTool = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if (-not $SignTool) {
        # Not on PATH by default even with the Windows SDK installed — it lives under a
        # version-specific folder, so search for it rather than hardcoding one.
        $Candidate = Get-ChildItem -Path "${env:ProgramFiles(x86)}\Windows Kits\10\bin" -Recurse -Filter "signtool.exe" -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match "\\x64\\" } | Select-Object -First 1
        if ($Candidate) { $SignTool = $Candidate.FullName } else { throw "signtool.exe not found - install the Windows SDK, or run this from a Developer PowerShell." }
    } else {
        $SignTool = $SignTool.Source
    }

    Write-Host "==> Signing $ExePath"
    $SignArgs = @("sign", "/fd", "SHA256", "/tr", $TimestampUrl, "/td", "SHA256")
    if ($CertThumbprint) {
        $SignArgs += @("/sha1", $CertThumbprint)
    } else {
        if (-not $CertPfxPassword) { throw "CertPfxPath was given without CertPfxPassword." }
        $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($CertPfxPassword))
        $SignArgs += @("/f", $CertPfxPath, "/p", $PlainPassword)
    }

    & $SignTool @SignArgs $ExePath
    if ($LASTEXITCODE -ne 0) { throw "signtool.exe failed" }
    Write-Host "==> Verifying signature"
    & $SignTool verify /pa $ExePath
    if ($LASTEXITCODE -ne 0) { throw "Signature verification failed" }
} else {
    Write-Host "==> Skipping code signing (no -CertThumbprint or -CertPfxPath given) - Windows SmartScreen will show an Unknown Publisher warning until this build is signed."
}

Write-Host "==> Zipping"
Compress-Archive -Path $PackageDir -DestinationPath $ZipPath -Force

Write-Host "==> Done: $ZipPath"
