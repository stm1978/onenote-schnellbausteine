param(
    [switch]$KeepProgramFiles
)

$ErrorActionPreference = "Stop"
$progId = "Schnellbausteine.OneNoteAddIn"
$clsid = "{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}"
$appRoot = Join-Path $env:LOCALAPPDATA "Schnellbausteine\AddIn"
$dataRoot = Join-Path $env:LOCALAPPDATA "Schnellbausteine\Data"

function Remove-Registration([Microsoft.Win32.RegistryView]$view) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        foreach ($path in @(
            "Software\Microsoft\Office\OneNote\AddIns\$progId",
            "Software\Classes\CLSID\$clsid",
            "Software\Classes\AppID\$clsid",
            "Software\Classes\$progId"
        )) {
            $base.DeleteSubKeyTree($path, $false)
        }
    }
    finally {
        $base.Close()
    }
}

Remove-Registration ([Microsoft.Win32.RegistryView]::Registry32)
if ([Environment]::Is64BitOperatingSystem) {
    Remove-Registration ([Microsoft.Win32.RegistryView]::Registry64)
}

Write-Host "Schnellbausteine wurde in OneNote deaktiviert." -ForegroundColor Green

$oneNote = Get-Process ONENOTE -ErrorAction SilentlyContinue
if ($oneNote) {
    Write-Warning "OneNote läuft noch. Beenden Sie den hängenden Prozess im Task-Manager und starten Sie dieses Skript danach erneut, um die Programmdateien zu entfernen."
}
elseif (-not $KeepProgramFiles) {
    $uninstaller = Join-Path $appRoot "unins000.exe"
    if (Test-Path $uninstaller) {
        $process = Start-Process $uninstaller -ArgumentList @(
            "/VERYSILENT",
            "/SUPPRESSMSGBOXES",
            "/NORESTART"
        ) -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            throw "Die Deinstallation wurde mit Exitcode $($process.ExitCode) beendet."
        }
    }
    elseif (Test-Path $appRoot) {
        Remove-Item $appRoot -Recurse -Force
    }
}

Write-Host "Ihre Schnellbausteine bleiben erhalten: $dataRoot"
Write-Host "Starten Sie OneNote jetzt neu."
