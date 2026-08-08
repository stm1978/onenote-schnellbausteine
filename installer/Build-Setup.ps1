param(
    [ValidatePattern("^\d+\.\d+\.\d+([.-][0-9A-Za-z.-]+)?$")]
    [string]$Version = "0.1.0",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$toolRoot = Join-Path $root "tools\inno-7.0.2"
$compiler = Join-Path $toolRoot "ISCC.exe"
$download = Join-Path $root "tools\innosetup-7.0.2-x64.exe"
$downloadUrl = "https://github.com/jrsoftware/issrc/releases/download/is-7_0_2/innosetup-7.0.2-x64.exe"
$expectedHash = "5AD54CA3DEF786F8F4212552E54CC6D8D61329E2D24A1CFEE0571D42C2684FF1"

if (-not $SkipBuild) {
    & "$root\test.ps1" -Configuration Release -Platform x86
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    & "$root\test.ps1" -Configuration Release -Platform x64
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if (-not (Test-Path $compiler)) {
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $download) | Out-Null
    if (-not (Test-Path $download)) {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $download
    }
    $actualHash = (Get-FileHash $download -Algorithm SHA256).Hash
    if ($actualHash -ne $expectedHash) {
        throw "Die Prüfsumme des Inno-Setup-Installers ist ungültig. Erwartet: $expectedHash; erhalten: $actualHash"
    }
    $install = Start-Process $download -ArgumentList @(
        "/VERYSILENT",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
        "/SP-",
        "/CURRENTUSER",
        "/DIR=$toolRoot"
    ) -Wait -PassThru
    if ($install.ExitCode -ne 0) {
        throw "Inno Setup konnte nicht lokal bereitgestellt werden (Exitcode $($install.ExitCode))."
    }
    for ($attempt = 0; $attempt -lt 20 -and -not (Test-Path $compiler); $attempt++) {
        Start-Sleep -Milliseconds 250
    }
    if (-not (Test-Path $compiler)) { throw "Inno Setup wurde nicht im erwarteten Verzeichnis installiert: $toolRoot" }
}

$versionParts = ($Version -split "[-.]")
$fileVersion = "{0}.{1}.{2}.0" -f $versionParts[0],$versionParts[1],$versionParts[2]
& $compiler "/DAppVersion=$Version" "/DFileVersion=$fileVersion" (Join-Path $root "installer\Schnellbausteine.iss")
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$setup = Join-Path $root "artifacts\setup\Schnellbausteine-Setup-$Version.exe"
if (-not (Test-Path $setup)) { throw "Setup wurde nicht erzeugt: $setup" }
Write-Host "Installer erfolgreich: $setup"
