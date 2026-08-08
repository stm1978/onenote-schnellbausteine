Option Explicit

Dim addIn
Dim ribbonXml

Set addIn = CreateObject("Schnellbausteine.OneNoteAddIn")
ribbonXml = addIn.GetCustomUI("Microsoft.OneNote")

If InStr(ribbonXml, "id=""sqMain""") = 0 Or _
   InStr(ribbonXml, "label=""Schnellbausteine""") = 0 Or _
   InStr(ribbonXml, "onAction=""OpenSnippetMenu""") = 0 Then
    WScript.Echo "Ribbon XML fehlt oder ist unvollstaendig."
    WScript.Quit 1
End If

If InStr(ribbonXml, "dynamicMenu") <> 0 Or _
   InStr(ribbonXml, "id=""sqInsert""") <> 0 Or _
   InStr(ribbonXml, "id=""sqManage""") <> 0 Then
    WScript.Echo "Unsicheres dynamisches Ribbon-Menue ist noch aktiv."
    WScript.Quit 1
End If

WScript.Echo "COM-Aktivierung und Ribbon-Dispatch erfolgreich."
