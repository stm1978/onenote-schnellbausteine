Option Explicit

Dim addIn
Dim ribbonXml
Dim menuXml

Set addIn = CreateObject("Schnellbausteine.OneNoteAddIn")
ribbonXml = addIn.GetCustomUI("Microsoft.OneNote")

If InStr(ribbonXml, "sqInsertMenu") = 0 Then
    WScript.Echo "Ribbon XML fehlt oder ist unvollstaendig."
    WScript.Quit 1
End If

menuXml = addIn.GetSnippetMenuContent(Nothing)
If InStr(menuXml, "sqEmpty") = 0 Then
    WScript.Echo "Dynamischer Ribbon-Callback ist nicht erreichbar."
    WScript.Quit 1
End If

WScript.Echo "COM-Aktivierung und Ribbon-Dispatch erfolgreich."
