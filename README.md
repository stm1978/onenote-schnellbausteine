# Schnellbausteine fur OneNote

Schnellbausteine ist ein bewusst kleines, lokales COM-Add-in fur **OneNote Desktop
unter Windows** (Microsoft 365 bzw. OneNote 2016). Markierte OneNote-Inhalte
lassen sich als benannte Bausteine speichern und spater uber das Start-Menuband
an der Cursorposition wieder einfugen.

Das Projekt ist eine eigenstandige Implementierung. Die Architektur orientiert
sich an den offentlich dokumentierten Mustern von
[OneMore](https://github.com/stevencohn/OneMore) (MPL-2.0), enthalt aber keinen
kopierten OneMore-Quellcode.

## Voraussetzungen

- Windows 10/11
- OneNote Desktop aus Microsoft 365 oder OneNote 2016
- .NET Framework 4.8
- Zum Bauen: Visual Studio 2022 Build Tools mit **.NET-Desktopentwicklung** und
  installiertem OneNote/Office PIA
- PowerShell 5.1 oder neuer

OneNote und Add-in mussen dieselbe Bitness haben. Unter
`Datei > Konto > Info zu OneNote` steht, ob OneNote 32 oder 64 Bit verwendet.
x86 und x64 werden unterstutzt. ARM64 ist im MVP nicht enthalten.

## Bauen und testen

In einer Developer PowerShell:

```powershell
.\build.ps1 -Configuration Release -Platform x64
.\test.ps1 -Configuration Release -Platform x64
```

Fur 32-Bit-OneNote `-Platform x86` verwenden. `build.ps1` nutzt den mit .NET
Framework gelieferten C#-Compiler und die lokal installierten Office-PIAs. Die
Solution kann alternativ direkt mit Visual Studio/MSBuild gebaut werden.

Die Ausgaben liegen in `artifacts\<Plattform>\<Konfiguration>`.

## Installation

OneNote vor Installation und Deinstallation beenden. Danach **ohne
Administratorrechte**:

```powershell
.\installer\Install.ps1 -Platform x64 -Configuration Release
```

Das Skript kopiert die Dateien nach
`%LOCALAPPDATA%\Schnellbausteine\AddIn\<Plattform>` und registriert den
COM-Server sowie das OneNote-Add-in ausschliesslich fur den aktuellen Benutzer.
Es verwendet kein `regasm`. Registriert werden konsistent:

- ProgID `Schnellbausteine.OneNoteAddIn`
- CLSID `{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}`
- `HKCU\Software\Microsoft\Office\OneNote\AddIns\Schnellbausteine.OneNoteAddIn`
- `LoadBehavior=3`

Deinstallation:

```powershell
.\installer\Uninstall.ps1 -Platform x64
```

## Bedienung

1. Inhalt auf einer OneNote-Seite markieren.
2. Im Text-Kontextmenu **Als Schnellbaustein speichern...** wahlen und einen
   eindeutigen Namen eingeben. Bei Namenskonflikten wird explizit nachgefragt,
   ob uberschrieben oder ein anderer Name verwendet werden soll.
3. Im Register **Start** in der Gruppe **Schnellbausteine** das Menu
   **Einfugen** offnen. Eintrage sind alphabetisch sortiert.
4. **Verwalten...** benennt Eintrage um oder loscht sie nach Bestatigung.

Beim Einfugen ersetzt ein Baustein markierten Text wie in Word. Reine
Textfragmente werden direkt in den Lauf eingefugt. Blockinhalte (Listen,
Tabellen, Bilder oder mehrere Absatze) teilen den aktuellen Absatz sicher und
werden dazwischen eingesetzt. Bei einem technisch nicht eindeutig teilbaren
Cursor in Hyperlink- oder MathML-Inhalt wird der Baustein hinter dem betroffenen
Absatz eingefugt; der Spezialinhalt bleibt unverandert. Kann OneNote gar keine
Selection-Markierung liefern, wird am Ende des ersten Outlines eingefugt.

## Gespeicherte Daten

Alle Daten bleiben lokal unter:

```text
%LOCALAPPDATA%\Schnellbausteine\Data
```

`catalog.json` enthalt nur versionierte Metadaten. Jeder Baustein liegt unter
einer zufalligen ID als eigene XML-Datei in `snippets`; Namen werden nie als
Dateinamen verwendet. Schreibvorgange erfolgen uber eine temporare Datei im
gleichen Verzeichnis und anschliessendes atomisches Ersetzen. Bilder werden als
Base64-Daten im Fragment gespeichert. OneNote-`TagDef`-Abhangigkeiten werden
mitgespeichert und beim Einfugen auf die Zielseite abgebildet.

## Einschrankungen

- Nur OneNote Desktop fur Windows; nicht OneNote im Web, macOS oder die alte
  Windows-10-Store-App.
- Keine Cloud-Synchronisierung, Kategorien, KI, Vorlagen oder Platzhalter.
- Eingebettete Dateien sind nicht Ziel des MVP. Formatierter Text,
  Uberschriften, Listen, Tabellen, OneNote-Tags/Checkboxen und Bilder werden
  unterstutzt.
- Hyperlink-/MathML-Spezialcursor werden absichtlich nur an einer sicheren
  Absatzgrenze behandelt, damit kein Inhalt verloren geht.

## Troubleshooting

- **Add-in fehlt:** OneNote wirklich beenden, korrekte Bitness installieren und
  unter `Datei > Optionen > Add-Ins` nachsehen. Der Registrierungspfad muss
  `LoadBehavior=3` enthalten.
- **LoadBehavior wird 2:** Meist stimmt die Bitness nicht oder eine Abhangigkeit
  fehlt. Neu bauen/installieren und `%LOCALAPPDATA%\Schnellbausteine\addin.log`
  prufen.
- **Bilder fehlen:** Den Baustein erneut speichern. Der Speichervorgang bricht
  sichtbar ab, wenn OneNote fur ein Callback-Bild keine Binardaten liefert.
- **Defekter Katalog:** `catalog.json` wird nicht still verworfen. Die
  Fehlermeldung und `addin.log` nennen den betroffenen Pfad; die einzelnen
  Fragmentdateien bleiben erhalten.

## Architektur

- `Schnellbausteine.Core`: Repository, atomare Persistenz und OneNote-XML-Logik
- `Schnellbausteine.AddIn`: COM-Lebenszyklus, Ribbon, OneNote-PIA und WinForms
- `Schnellbausteine.Tests`: unabhangig ausfuhrbare Unit-Tests

Das Add-in nutzt `IDTExtensibility2`, `Microsoft.Office.Core.IRibbonExtensibility`,
Ribbon Custom UI `2009/07` und OneNote XML `XMLSchema.xs2013`. COM-Objekte werden
pro Aktion erzeugt und explizit freigegeben; das von `OnConnection` gelieferte
Application-Objekt wird nicht dauerhaft gehalten.
