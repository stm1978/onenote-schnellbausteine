param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [ValidateSet("x86", "x64")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$output = Join-Path $root "artifacts\$Platform\$Configuration"
$framework = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319"
$csc = Join-Path $framework "csc.exe"

function Find-GacAssembly([string]$name, [string[]]$versions) {
    $candidate = $null
    $gacRoots = @(
        (Join-Path $env:WINDIR "assembly\GAC_MSIL\$name"),
        (Join-Path $env:WINDIR "Microsoft.NET\assembly\GAC_MSIL\$name")
    )
    foreach ($version in $versions) {
        $candidate = Get-ChildItem $gacRoots -Recurse -Filter "$name.dll" -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match "\\$version\.0\.0\.0__" } |
            Select-Object -First 1
        if ($candidate) { break }
    }
    if (-not $candidate) { throw "$name wurde nicht im GAC gefunden. Installieren Sie OneNote oder das offizielle Microsoft Office PIA Redistributable." }
    return $candidate.FullName
}

if (-not (Test-Path $csc)) { throw ".NET Framework Compiler nicht gefunden: $csc" }
$oneNote = if ($env:ONENOTE_PIA_PATH -and (Test-Path $env:ONENOTE_PIA_PATH)) {
    $env:ONENOTE_PIA_PATH
} else {
    Find-GacAssembly "Microsoft.Office.Interop.OneNote" @("15", "14")
}
$office = if ($env:OFFICE_PIA_PATH -and (Test-Path $env:OFFICE_PIA_PATH)) {
    $env:OFFICE_PIA_PATH
} else {
    Find-GacAssembly "office" @("15", "14")
}
$extensibility = if ($env:EXTENSIBILITY_PIA_PATH -and (Test-Path $env:EXTENSIBILITY_PIA_PATH)) {
    $env:EXTENSIBILITY_PIA_PATH
} else {
    @(
        (Join-Path ${env:ProgramFiles(x86)} "Common Files\Microsoft Shared\MSEnv\PublicAssemblies\extensibility.dll"),
        (Join-Path $env:ProgramFiles "Common Files\Microsoft Shared\MSEnv\PublicAssemblies\extensibility.dll")
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $extensibility) {
    throw "Extensibility PIA wurde nicht gefunden. Installieren Sie das offizielle Microsoft Office PIA Redistributable."
}

New-Item -ItemType Directory -Force -Path $output | Out-Null
$common = @("/nologo", "/warnaserror+", "/langversion:5", "/optimize+", "/debug:pdbonly")
if ($Configuration -eq "Debug") { $common = @("/nologo", "/warnaserror+", "/langversion:5", "/debug:full", "/define:DEBUG;TRACE") }

$coreSources = Get-ChildItem "$root\src\Schnellbausteine.Core" -Recurse -Filter *.cs | ForEach-Object FullName
& $csc @common /target:library /platform:anycpu /out:"$output\Schnellbausteine.Core.dll" /r:System.Runtime.Serialization.dll /r:System.Xml.Linq.dll $coreSources
if ($LASTEXITCODE -ne 0) { throw "Core-Kompilierung fehlgeschlagen." }

$addinSources = Get-ChildItem "$root\src\Schnellbausteine.AddIn" -Recurse -Filter *.cs | ForEach-Object FullName
& $csc @common /target:library /platform:$Platform /out:"$output\Schnellbausteine.AddIn.dll" /r:"$output\Schnellbausteine.Core.dll" /r:System.Xml.Linq.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:"$oneNote" /link:"$office" /link:"$extensibility" /resource:"$root\src\Schnellbausteine.AddIn\Ribbon.xml,Schnellbausteine.AddIn.Ribbon.xml" $addinSources
if ($LASTEXITCODE -ne 0) { throw "Add-in-Kompilierung fehlgeschlagen." }
Copy-Item $oneNote "$output\Microsoft.Office.Interop.OneNote.dll" -Force

$testSources = Get-ChildItem "$root\tests\Schnellbausteine.Tests" -Recurse -Filter *.cs | ForEach-Object FullName
& $csc @common /target:exe /platform:anycpu /out:"$output\Schnellbausteine.Tests.exe" /r:"$output\Schnellbausteine.Core.dll" /r:System.Xml.Linq.dll $testSources
if ($LASTEXITCODE -ne 0) { throw "Test-Kompilierung fehlgeschlagen." }

$surrogateSource = Join-Path $root "tests\SurrogateSmoke.cs"
& $csc @common /target:exe /platform:$Platform /out:"$output\Schnellbausteine.SurrogateSmoke.exe" /link:"$extensibility" $surrogateSource
if ($LASTEXITCODE -ne 0) { throw "Surrogate-Test-Kompilierung fehlgeschlagen." }

$addInSmokeSource = Join-Path $root "tests\AddInSmoke.cs"
& $csc @common /target:exe /platform:$Platform /out:"$output\Schnellbausteine.AddInSmoke.exe" /r:System.Xml.Linq.dll $addInSmokeSource
if ($LASTEXITCODE -ne 0) { throw "Add-in-Smoke-Test-Kompilierung fehlgeschlagen." }

Write-Host "Build erfolgreich: $output"
