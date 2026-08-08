param(
    [Parameter(Mandatory = $true)]
    [string]$SetupPath,
    [ValidateSet("x86", "x64")]
    [string]$Platform
)

$ErrorActionPreference = "Stop"
$setup = (Resolve-Path $SetupPath).Path
$appRoot = Join-Path $env:LOCALAPPDATA "Schnellbausteine\AddIn"
$dataRoot = Join-Path $env:LOCALAPPDATA "Schnellbausteine\Data"
$registryView = if ($Platform -eq "x86") {
    [Microsoft.Win32.RegistryView]::Registry32
} else {
    [Microsoft.Win32.RegistryView]::Registry64
}
$cscript = if ($Platform -eq "x86") {
    "$env:WINDIR\SysWOW64\cscript.exe"
} else {
    "$env:WINDIR\System32\cscript.exe"
}

New-Item -ItemType Directory -Force -Path $dataRoot | Out-Null
$marker = Join-Path $dataRoot ("installer-test-" + [Guid]::NewGuid().ToString("N") + ".txt")
Set-Content -Path $marker -Value "must survive uninstall"

try {
    $install = Start-Process $setup -ArgumentList @(
        "/VERYSILENT",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
        "/ONENOTEBITNESS=$Platform"
    ) -Wait -PassThru
    if ($install.ExitCode -ne 0) { throw "Setup wurde mit Exitcode $($install.ExitCode) beendet." }

    $assembly = Join-Path $appRoot "$Platform\Schnellbausteine.AddIn.dll"
    if (-not (Test-Path $assembly)) { throw "Installierte Assembly fehlt: $assembly" }

    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser, $registryView)
    try {
        $addIn = $base.OpenSubKey(
            "Software\Microsoft\Office\OneNote\AddIns\Schnellbausteine.OneNoteAddIn")
        if (-not $addIn) { throw "OneNote-Add-in-Registrierung fehlt." }
        if ($addIn.GetValue("LoadBehavior") -ne 3) { throw "LoadBehavior ist nicht 3." }
        $addIn.Close()

        $com = $base.OpenSubKey(
            "Software\Classes\CLSID\{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}\InprocServer32")
        if (-not $com) { throw "COM-Registrierung fehlt." }
        $codeBase = [string]$com.GetValue("CodeBase")
        $com.Close()
        if ($codeBase -notmatch "/$Platform/Schnellbausteine\.AddIn\.dll$") {
            throw "CodeBase verweist auf die falsche Bitness: $codeBase"
        }
    }
    finally {
        $base.Close()
    }

    & $cscript //nologo "$PSScriptRoot\..\tests\ComSmoke.vbs"
    if ($LASTEXITCODE -ne 0) { throw "COM-/Ribbon-Smoke-Test fehlgeschlagen." }

    $upgrade = Start-Process $setup -ArgumentList @(
        "/VERYSILENT",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
        "/ONENOTEBITNESS=$Platform"
    ) -Wait -PassThru
    if ($upgrade.ExitCode -ne 0) { throw "Upgrade wurde mit Exitcode $($upgrade.ExitCode) beendet." }
}
finally {
    $uninstaller = Join-Path $appRoot "unins000.exe"
    if (Test-Path $uninstaller) {
        $uninstall = Start-Process $uninstaller -ArgumentList @(
            "/VERYSILENT",
            "/SUPPRESSMSGBOXES",
            "/NORESTART"
        ) -Wait -PassThru
        if ($uninstall.ExitCode -ne 0) {
            throw "Deinstallation wurde mit Exitcode $($uninstall.ExitCode) beendet."
        }
    }

    if (Test-Path (Join-Path $appRoot "$Platform\Schnellbausteine.AddIn.dll")) {
        throw "Add-in-Dateien wurden bei der Deinstallation nicht entfernt."
    }
    if (-not (Test-Path $marker)) {
        throw "Die Deinstallation hat Benutzerdaten entfernt."
    }
    Remove-Item $marker -Force
}

Write-Host "Setup-Lebenszyklus fur $Platform erfolgreich."
