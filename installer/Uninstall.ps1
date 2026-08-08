param(
    [ValidateSet("x86", "x64")]
    [string]$Platform = "x64",
    [switch]$DeleteUserData
)

$ErrorActionPreference = "Stop"
$view = if ($Platform -eq "x86") { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }
$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
$clsid = "{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}"
$progId = "Schnellbausteine.OneNoteAddIn"
try {
    $base.DeleteSubKeyTree("Software\Microsoft\Office\OneNote\AddIns\$progId", $false)
    $base.DeleteSubKeyTree("Software\Classes\CLSID\$clsid", $false)
    $base.DeleteSubKeyTree("Software\Classes\AppID\$clsid", $false)
    $base.DeleteSubKeyTree("Software\Classes\$progId", $false)
}
finally {
    $base.Close()
}

$addInPath = Join-Path $env:LOCALAPPDATA "Schnellbausteine\AddIn\$Platform"
if (Test-Path $addInPath) { Remove-Item $addInPath -Recurse -Force }
if ($DeleteUserData) {
    $dataPath = Join-Path $env:LOCALAPPDATA "Schnellbausteine\Data"
    if (Test-Path $dataPath) { Remove-Item $dataPath -Recurse -Force }
}
Write-Host "Schnellbausteine ($Platform) deinstalliert."
