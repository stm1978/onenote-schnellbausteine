param(
    [ValidateSet("x86", "x64")]
    [string]$Platform = "x64",
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$source = Join-Path $root "artifacts\$Platform\$Configuration"
if (-not (Test-Path "$source\Schnellbausteine.AddIn.dll")) {
    throw "Build nicht gefunden. Zuerst .\build.ps1 -Configuration $Configuration -Platform $Platform ausfuhren."
}

$target = Join-Path $env:LOCALAPPDATA "Schnellbausteine\AddIn\$Platform"
New-Item -ItemType Directory -Force -Path $target | Out-Null
Copy-Item "$source\Schnellbausteine.AddIn.dll","$source\Schnellbausteine.Core.dll" -Destination $target -Force

$view = if ($Platform -eq "x86") { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }
$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
$clsid = "{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}"
$progId = "Schnellbausteine.OneNoteAddIn"
$assemblyPath = Join-Path $target "Schnellbausteine.AddIn.dll"
$assemblyName = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath).FullName
$codeBase = ([Uri]$assemblyPath).AbsoluteUri

try {
    $class = $base.CreateSubKey("Software\Classes\CLSID\$clsid")
    $class.SetValue("", "Schnellbausteine fur OneNote")
    $class.SetValue("AppID", $clsid)
    $class.CreateSubKey("ProgID").SetValue("", $progId)
    $server = $class.CreateSubKey("InprocServer32")
    $server.SetValue("", "mscoree.dll")
    $server.SetValue("ThreadingModel", "Both")
    $server.SetValue("Class", "Schnellbausteine.AddIn.AddIn")
    $server.SetValue("Assembly", $assemblyName)
    $server.SetValue("RuntimeVersion", "v4.0.30319")
    $server.SetValue("CodeBase", $codeBase)
    $server.Close()
    $class.Close()

    $prog = $base.CreateSubKey("Software\Classes\$progId")
    $prog.SetValue("", "Schnellbausteine fur OneNote")
    $prog.CreateSubKey("CLSID").SetValue("", $clsid)
    $prog.Close()

    $appId = $base.CreateSubKey("Software\Classes\AppID\$clsid")
    $appId.SetValue("", "Schnellbausteine fur OneNote")
    $appId.SetValue("DllSurrogate", "")
    $appId.Close()

    $addIn = $base.CreateSubKey("Software\Microsoft\Office\OneNote\AddIns\$progId")
    $addIn.SetValue("FriendlyName", "Schnellbausteine")
    $addIn.SetValue("Description", "Lokale Schnellbausteine fur OneNote Desktop")
    $addIn.SetValue("LoadBehavior", 3, [Microsoft.Win32.RegistryValueKind]::DWord)
    $addIn.Close()
}
finally {
    $base.Close()
}

Write-Host "Schnellbausteine ($Platform) fur den aktuellen Benutzer installiert. OneNote jetzt starten."
