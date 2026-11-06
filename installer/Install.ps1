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
    throw "Build nicht gefunden. Zuerst .\build.ps1 -Configuration $Configuration -Platform $Platform ausführen."
}

$target = Join-Path $env:LOCALAPPDATA "Schnellbausteine\AddIn\$Platform"
New-Item -ItemType Directory -Force -Path $target | Out-Null
Copy-Item "$source\Schnellbausteine.AddIn.dll","$source\Schnellbausteine.Core.dll","$source\Microsoft.Office.Interop.OneNote.dll" -Destination $target -Force

$view = if ($Platform -eq "x86") { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }
$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
$clsid = "{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}"
$progId = "Schnellbausteine.OneNoteAddIn"
$assemblyPath = Join-Path $target "Schnellbausteine.AddIn.dll"
$assembly = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath)
$assemblyName = $assembly.FullName
$assemblyVersion = $assembly.Version.ToString()
$codeBase = ([Uri]$assemblyPath).AbsoluteUri

try {
    $class = $base.CreateSubKey("Software\Classes\CLSID\$clsid")
    $class.SetValue("", "OneNote Schnellbausteine")
    $class.SetValue("AppID", $clsid)
    $class.CreateSubKey("ProgID").SetValue("", $progId)
    $class.CreateSubKey("VersionIndependentProgID").SetValue("", $progId)
    $class.CreateSubKey("Programmable").SetValue("", "")
    $class.CreateSubKey(
        "Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}").Close()
    $server = $class.CreateSubKey("InprocServer32")
    $server.SetValue("", "mscoree.dll")
    $server.SetValue("ThreadingModel", "Both")
    $server.SetValue("Class", "Schnellbausteine.AddIn.AddIn")
    $server.SetValue("Assembly", $assemblyName)
    $server.SetValue("RuntimeVersion", "v4.0.30319")
    $server.SetValue("CodeBase", $codeBase)
    $versionedServer = $server.CreateSubKey($assemblyVersion)
    $versionedServer.SetValue("Class", "Schnellbausteine.AddIn.AddIn")
    $versionedServer.SetValue("Assembly", $assemblyName)
    $versionedServer.SetValue("RuntimeVersion", "v4.0.30319")
    $versionedServer.SetValue("CodeBase", $codeBase)
    $versionedServer.Close()
    $server.Close()
    $class.Close()

    $prog = $base.CreateSubKey("Software\Classes\$progId")
    $prog.SetValue("", "OneNote Schnellbausteine")
    $prog.CreateSubKey("CLSID").SetValue("", $clsid)
    $prog.Close()

    $appId = $base.CreateSubKey("Software\Classes\AppID\$clsid")
    $appId.SetValue("", "OneNote Schnellbausteine")
    $appId.SetValue("DllSurrogate", "")
    $appId.Close()

    $addIn = $base.CreateSubKey("Software\Microsoft\Office\OneNote\AddIns\$progId")
    $addIn.SetValue("FriendlyName", "OneNote Schnellbausteine")
    $addIn.SetValue("Description", "OneNote Schnellbausteine – Beta 0.2.4")
    $addIn.SetValue("LoadBehavior", 3, [Microsoft.Win32.RegistryValueKind]::DWord)
    $addIn.Close()
}
finally {
    $base.Close()
}

Write-Host "OneNote Schnellbausteine ($Platform) für den aktuellen Benutzer installiert. OneNote jetzt starten."
