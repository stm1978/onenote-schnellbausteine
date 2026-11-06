param(
    [Parameter(Mandatory = $true)]
    [string]$SetupPath,
    [ValidateSet("x86", "x64")]
    [string]$Platform,
    [switch]$SkipHostActivation
)

$ErrorActionPreference = "Stop"
$setup = (Resolve-Path $SetupPath).Path
$addInAppId = "CCF3B64A-9580-4D44-BDCE-49BFE95AC510"

function Get-AddInSurrogate {
    @(Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" |
        Where-Object { $_.CommandLine -match $addInAppId })
}

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
$surrogateSmoke = Join-Path $PSScriptRoot (
    "..\artifacts\$Platform\Release\Schnellbausteine.SurrogateSmoke.exe")
if (-not (Test-Path $surrogateSmoke)) {
    throw "Surrogate-Lifecycle-Test fehlt: $surrogateSmoke"
}
$addInSmoke = Join-Path $PSScriptRoot (
    "..\artifacts\$Platform\Release\Schnellbausteine.AddInSmoke.exe")
if (-not (Test-Path $addInSmoke)) {
    throw "Managed-Add-in-Test fehlt: $addInSmoke"
}
if ((Get-AddInSurrogate).Count -ne 0) {
    throw "Ein alter Schnellbausteine-COM-Surrogatprozess läuft noch."
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
    $assemblyInfo = [Diagnostics.FileVersionInfo]::GetVersionInfo($assembly)
    if ($assemblyInfo.ProductName -ne "OneNote Schnellbausteine") {
        throw "Assembly-Produktname ist inkonsistent: $($assemblyInfo.ProductName)"
    }
    if ($assemblyInfo.CompanyName -ne "Stefan Malter | Malter365.de") {
        throw "Assembly-Publisher ist inkonsistent: $($assemblyInfo.CompanyName)"
    }
    if ($assemblyInfo.FileVersion -ne "0.2.4.0") {
        throw "Assembly-Dateiversion ist inkonsistent: $($assemblyInfo.FileVersion)"
    }
    if ($assemblyInfo.ProductVersion -ne "Beta 0.2.4") {
        throw "Assembly-Produktversion ist inkonsistent: $($assemblyInfo.ProductVersion)"
    }
    $oneNotePia = Join-Path $appRoot "$Platform\Microsoft.Office.Interop.OneNote.dll"
    if (-not (Test-Path $oneNotePia)) {
        throw "Installierte OneNote-PIA fehlt: $oneNotePia"
    }

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

        $managedCategory = $base.OpenSubKey(
            "Software\Classes\CLSID\{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}" +
            "\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}")
        if (-not $managedCategory) { throw "Managed-COM-Kategorie fehlt." }
        $managedCategory.Close()

        $versionedServer = $base.OpenSubKey(
            "Software\Classes\CLSID\{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}" +
            "\InprocServer32\1.0.0.0")
        if (-not $versionedServer) { throw "Versionierte COM-Registrierung fehlt." }
        $versionedServer.Close()

        $appId = $base.OpenSubKey(
            "Software\Classes\AppID\{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}")
        if (-not $appId) { throw "COM-AppID-Registrierung fehlt." }
        if ($null -eq $appId.GetValue("DllSurrogate")) {
            throw "DllSurrogate-Registrierung fehlt."
        }
        $appId.Close()
    }
    finally {
        $base.Close()
    }

    $uninstallBase = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        [Microsoft.Win32.RegistryView]::Registry32)
    try {
        $uninstall = $uninstallBase.OpenSubKey(
            "Software\Microsoft\Windows\CurrentVersion\Uninstall" +
            "\{5B1518D1-0620-49BC-A423-0684C3E21917}_is1")
        if (-not $uninstall) { throw "Add/Remove-Programs-Registrierung fehlt." }
        try {
            if ($uninstall.GetValue("DisplayName") -ne "OneNote Schnellbausteine") {
                throw "ARP-Produktname ist inkonsistent."
            }
            if ($uninstall.GetValue("DisplayVersion") -ne "Beta 0.2.4") {
                throw "ARP-DisplayVersion ist inkonsistent."
            }
            if ($uninstall.GetValue("Publisher") -ne "Stefan Malter | Malter365.de") {
                throw "ARP-Publisher ist inkonsistent."
            }
        }
        finally {
            $uninstall.Close()
        }
    }
    finally {
        $uninstallBase.Close()
    }

    & $addInSmoke $assembly
    if ($LASTEXITCODE -ne 0) { throw "Managed-Add-in-/Ribbon-Test fehlgeschlagen." }

    if (-not $SkipHostActivation) {
        $comOutput = @(& $cscript //nologo "$PSScriptRoot\..\tests\ComSmoke.vbs" 2>&1)
        $comOutput | ForEach-Object { Write-Host $_ }
        if ($LASTEXITCODE -ne 0 -or
            ($comOutput -join [Environment]::NewLine) -notmatch
                "COM-Aktivierung und Ribbon-Dispatch erfolgreich") {
            throw "COM-/Ribbon-Smoke-Test fehlgeschlagen."
        }

        & $surrogateSmoke
        if ($LASTEXITCODE -ne 0) {
            throw "COM-Surrogate-Lifecycle-Test fehlgeschlagen."
        }

        for ($second = 1;
            $second -le 15 -and (Get-AddInSurrogate).Count -ne 0;
            $second++) {
            Start-Sleep -Seconds 1
        }
        if ((Get-AddInSurrogate).Count -ne 0) {
            throw "Der COM-Surrogatprozess wurde nach dem Lifecycle-Test nicht beendet."
        }
    }

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

Write-Host "Setup-Lebenszyklus für $Platform erfolgreich."
