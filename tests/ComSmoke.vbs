Option Explicit

Dim addIn
Dim ribbonXml

Set addIn = CreateObject("Schnellbausteine.OneNoteAddIn")
ribbonXml = addIn.GetCustomUI("Microsoft.OneNote")

If InStr(ribbonXml, "<tab idMso=""TabInsert"">") = 0 Or _
   InStr(ribbonXml, "<group id=""sqGroup"" label=""Malter365.de"">") = 0 Or _
   InStr(ribbonXml, "<dynamicMenu id=""sqMain"" label=""Schnellbausteine""") = 0 Or _
   InStr(ribbonXml, "<dynamicMenu id=""sqInsertText"" label=""Schnellbausteine""") = 0 Or _
   InStr(ribbonXml, "label=""Als Schnellbaustein speichern...""") = 0 Or _
   InStr(ribbonXml, "<dynamicMenu") = 0 Or _
   InStr(ribbonXml, "getContent=""GetSnippetMenuContent""") = 0 Then
    WScript.Echo "Ribbon XML fehlt oder ist unvollstaendig."
    WScript.Quit 1
End If

If InStr(ribbonXml, "<gallery") <> 0 Or _
   InStr(ribbonXml, "<menu id=""sqMain""") <> 0 Or _
   InStr(ribbonXml, "label=""OneNote Schnellbausteine""") <> 0 Or _
   InStr(ribbonXml, "<tab idMso=""TabHome"">") <> 0 Or _
   InStr(ribbonXml, "OpenSnippetMenu") <> 0 Or _
   InStr(ribbonXml, "InsertGallerySnippet") <> 0 Or _
   InStr(ribbonXml, "id=""sqInsert""") <> 0 Or _
   InStr(ribbonXml, "id=""sqManage""") <> 0 Then
    WScript.Echo "Veraltete Ribbon-Menuestruktur ist noch aktiv."
    WScript.Quit 1
End If

WScript.Echo "COM-Aktivierung und Ribbon-Dispatch erfolgreich."
