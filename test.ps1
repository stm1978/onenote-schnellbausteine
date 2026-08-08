param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [ValidateSet("x86", "x64")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$root\build.ps1" -Configuration $Configuration -Platform $Platform
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& "$root\artifacts\$Platform\$Configuration\Schnellbausteine.Tests.exe"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& "$root\artifacts\$Platform\$Configuration\Schnellbausteine.OneNoteClientHarness.exe"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& "$root\artifacts\$Platform\$Configuration\Schnellbausteine.AddInSmoke.exe" `
    "$root\artifacts\$Platform\$Configuration\Schnellbausteine.AddIn.dll"
exit $LASTEXITCODE
