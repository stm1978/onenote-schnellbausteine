param()

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

function Assert-Contains([string]$Path, [string]$Expected) {
    $content = Get-Content (Join-Path $root $Path) -Raw
    if (-not $content.Contains($Expected)) {
        throw "$Path enthält nicht den erwarteten Wert: $Expected"
    }
}

function Assert-Equal([string]$Description, [object]$Actual, [object]$Expected) {
    if ($Actual -ne $Expected) {
        throw "$Description ist inkonsistent. Erwartet: $Expected; erhalten: $Actual"
    }
}

$product = "OneNote Schnellbausteine"
$publisher = "Stefan Malter | Malter365.de"
$displayVersion = "Beta 0.2"
$technicalVersion = "0.2.0-beta.1"
$setupName = "OneNote-Schnellbausteine-Setup-Beta-0.2.exe"

Assert-Contains "src\Schnellbausteine.AddIn\Properties\AssemblyInfo.cs" "AssemblyProduct(`"$product`")"
Assert-Contains "src\Schnellbausteine.AddIn\Properties\AssemblyInfo.cs" "AssemblyCompany(`"$publisher`")"
Assert-Contains "src\Schnellbausteine.AddIn\Properties\AssemblyInfo.cs" "AssemblyVersion(`"1.0.0.0`")"
Assert-Contains "src\Schnellbausteine.AddIn\Properties\AssemblyInfo.cs" "AssemblyFileVersion(`"0.2.0.0`")"
Assert-Contains "src\Schnellbausteine.AddIn\Properties\AssemblyInfo.cs" "AssemblyInformationalVersion(`"$displayVersion`")"
Assert-Contains "src\Schnellbausteine.Core\Properties\AssemblyInfo.cs" "AssemblyProduct(`"$product`")"
Assert-Contains "src\Schnellbausteine.AddIn\AddIn.cs" 'Guid("CCF3B64A-9580-4D44-BDCE-49BFE95AC510")'
Assert-Contains "src\Schnellbausteine.AddIn\AddIn.cs" 'ProgId("Schnellbausteine.OneNoteAddIn")'

Assert-Contains "installer\Schnellbausteine.iss" "#define AppName `"$product`""
Assert-Contains "installer\Schnellbausteine.iss" "#define AppPublisher `"$publisher`""
Assert-Contains "installer\Schnellbausteine.iss" "#define AppVersion `"$displayVersion`""
Assert-Contains "installer\Schnellbausteine.iss" "VersionInfoProductTextVersion={#AppVersion}"
Assert-Contains "installer\Schnellbausteine.iss" '#define AppIdValue "{{5B1518D1-0620-49BC-A423-0684C3E21917}"'
Assert-Contains "installer\Schnellbausteine.iss" '#define AddInClsid "{{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}"'
Assert-Contains "installer\Schnellbausteine.iss" '#define AddInProgId "Schnellbausteine.OneNoteAddIn"'
Assert-Contains "installer\Schnellbausteine.iss" "OutputBaseFilename=$($setupName.Substring(0, $setupName.Length - 4))"
Assert-Contains "installer\Schnellbausteine.iss" "DefaultDirName={localappdata}\Schnellbausteine\AddIn"
Assert-Contains "installer\Test-Setup.ps1" 'Join-Path $env:LOCALAPPDATA "Schnellbausteine\Data"'
Assert-Contains "installer\Test-Setup.ps1" 'if ($uninstall.GetValue("DisplayVersion") -ne "Beta 0.2")'
Assert-Contains "installer\Schnellbausteine.iss" "ArchitecturesAllowed=x86compatible x64compatible"
Assert-Contains "installer\Schnellbausteine.iss" "OneNotePlatform := ReadPlatformFromRegistry();"
Assert-Contains "installer\Build-Setup.ps1" "[string]`$Version = `"$technicalVersion`""

$ribbonPath = Join-Path $root "src\Schnellbausteine.AddIn\Ribbon.xml"
$ribbon = New-Object Xml.XmlDocument
$ribbon.PreserveWhitespace = $true
$ribbon.Load($ribbonPath)
$namespaces = New-Object Xml.XmlNamespaceManager($ribbon.NameTable)
$namespaces.AddNamespace("r", "http://schemas.microsoft.com/office/2009/07/customui")
$insertTab = $ribbon.SelectSingleNode("//r:tab[@idMso='TabInsert']", $namespaces)
if (-not $insertTab) { throw "Ribbon.xml enthält keine TabInsert-Integration." }
if ($ribbon.SelectSingleNode("//r:tab[@idMso='TabHome']", $namespaces)) {
    throw "Ribbon.xml darf keine TabHome-Integration enthalten."
}
$group = $insertTab.SelectSingleNode("r:group[@id='sqGroup']", $namespaces)
if (-not $group) { throw "Ribbon.xml enthält die erwartete Ribbon-Gruppe nicht." }
Assert-Equal "Ribbon-Gruppenbeschriftung" $group.label "Malter365.de"
$mainMenu = $group.SelectSingleNode("r:dynamicMenu[@id='sqMain']", $namespaces)
if (-not $mainMenu) { throw "Ribbon.xml enthält das erwartete dynamicMenu nicht." }
Assert-Equal "Ribbon-Control-Beschriftung" $mainMenu.label "Schnellbausteine"
$contextMenus = @($ribbon.SelectNodes(
    "/r:customUI/r:contextMenus/r:contextMenu/r:dynamicMenu", $namespaces))
Assert-Equal "Anzahl der Kontext-Untermenüs" $contextMenus.Count 4
foreach ($contextMenu in $contextMenus) {
    Assert-Equal "Kontext-Untermenü-Beschriftung" $contextMenu.label "Schnellbausteine"
}
$saveButtons = @($ribbon.SelectNodes(
    "/r:customUI/r:contextMenus/r:contextMenu/r:button", $namespaces))
Assert-Equal "Anzahl der Kontext-Speicherbefehle" $saveButtons.Count 4
foreach ($saveButton in $saveButtons) {
    Assert-Equal "Kontext-Speicherbefehl" $saveButton.label "Als Schnellbaustein speichern..."
}
if ($ribbon.OuterXml.Contains("label=`"$product`"")) {
    throw "Ribbon.xml darf den Produktnamen nicht als sichtbare Ribbon-Beschriftung verwenden."
}
Assert-Contains "README.md" $displayVersion
Assert-Contains ".github\workflows\build-release.yml" $setupName

Write-Host "Branding- und Upgrade-Kompatibilitätsprüfung erfolgreich."
