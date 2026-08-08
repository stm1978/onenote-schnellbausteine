# Schnellbausteine für OneNote

Schnellbausteine ist ein bewusst kleines, lokales COM-Add-in für **OneNote
Desktop unter Windows** (Microsoft 365 bzw. OneNote 2016). Markierte
OneNote-Inhalte lassen sich als benannte Bausteine speichern und später über das
Start-Menüband an der Cursorposition wieder einfügen.

Das Projekt ist eine eigenständige Implementierung. Die Architektur orientiert
sich an den öffentlich dokumentierten Mustern von
[OneMore](https://github.com/stevencohn/OneMore) (MPL-2.0), enthält aber keinen
kopierten OneMore-Quellcode.

> [!CAUTION]
> **`v0.1.0-preview.1` ist zurückgezogen und darf nicht installiert werden.**
> Diese Version konnte OneNote beim Start blockieren. Falls sie noch installiert
> ist, OneNote schließen und das datenbewahrende
> [`Recover-Schnellbausteine.ps1`](installer/Recover-Schnellbausteine.ps1)
> ausführen oder die App nach einem Windows-Neustart normal deinstallieren.
> Gespeicherte Bausteine unter `%LOCALAPPDATA%\Schnellbausteine\Data` werden
> dabei nicht gelöscht.

## Installation für Endnutzer

1. Auf der
   [Release-Seite](https://github.com/stm1978/onenote-schnellbausteine/releases)
   die neueste, nicht zurückgezogene Datei
   **`Schnellbausteine-Setup-....exe`** herunterladen.
2. OneNote vollständig schließen.
3. Die heruntergeladene Setup-Datei doppelklicken und **Installieren** wählen.
4. OneNote neu starten.

Setup erkennt automatisch, ob OneNote 32 oder 64 Bit verwendet, und installiert
die passende Add-in-Version ausschließlich für das aktuelle Windows-Konto.
Administratorrechte und PowerShell-Befehle sind nicht erforderlich.

**Voraussetzungen:** Windows 10/11, .NET Framework 4.8 und OneNote Desktop aus
Microsoft 365 oder OneNote 2016. ARM64 ist im MVP nicht enthalten.

Die Deinstallation erfolgt wie bei anderen Windows-Programmen unter
**Einstellungen > Apps > Installierte Apps > Schnellbausteine für OneNote**.
Gespeicherte Schnellbausteine bleiben dabei erhalten.

> Die Vorabversion ist noch nicht digital signiert. Windows SmartScreen kann
> deshalb beim ersten Start **Weitere Informationen > Trotzdem ausführen**
> verlangen. Setup selbst benötigt keine erhöhten Rechte.

## Bedienung

1. Inhalt auf einer OneNote-Seite markieren.
2. Im Text-Kontextmenü **Als Schnellbaustein speichern...** wählen und einen
   eindeutigen Namen eingeben. Bei Namenskonflikten wird explizit nachgefragt,
   ob überschrieben oder ein anderer Name verwendet werden soll.
3. Im Register **Start** auf **Schnellbausteine** klicken und den gewünschten,
   alphabetisch sortierten direkten nativen Office-Menüeintrag wählen. Unter der
   nativen Trennlinie öffnet **Schnellbausteine verwalten** die bestehende
   Verwaltung zum Umbenennen oder Löschen nach Bestätigung.

Beim Einfügen ersetzt ein Baustein markierten Text wie in Word. Reine
Textfragmente werden direkt in den Lauf eingefügt. Blockinhalte (Listen,
Tabellen, Bilder oder mehrere Absätze) teilen den aktuellen Absatz sicher und
werden dazwischen eingesetzt. Bei einem technisch nicht eindeutig teilbaren
Cursor in Hyperlink- oder MathML-Inhalt wird der Baustein hinter dem betroffenen
Absatz eingefügt; der Spezialinhalt bleibt unverändert. Nach dem Klick auf einen
Menüeintrag liest der verzögerte STA-Einfügepfad die Seite einschließlich
`piBinaryDataSelection`. Fehlt die Cursor- oder Auswahlmarkierung oder ist sie
mehrdeutig, bricht das Add-in sichtbar ab, statt an einer geratenen
Outline-Position einzufügen.

Menü, Tastatursteuerung, Escape und Click-away werden vollständig von Office
bereitgestellt; das Add-in öffnet keine eigene Popup-Oberfläche.

> **Experimentell:** Offline lässt sich nicht beweisen, ob OneNote die
> `piBinaryDataSelection`-Markierung beim Wechsel vom Seiteneditor in die native
> Ribbon-Menü bis zum direkten `button.onAction` erhält. Der Insertionspfad protokolliert
> deshalb ausschließlich Scope, Markertypen, Markeranzahl und XML-Pfade, niemals
> Nutztext, und fügt ohne eindeutige Markierung nichts ein.

Ist der Ziel-Textcontainer schmaler als der gespeicherte Baustein, verbreitert das
Add-in nur diesen Container auf die gespeicherte Mindestbreite. Breitere Container
werden niemals verkleinert; Höhe, Position und andere Outlines bleiben unverändert.

## Gespeicherte Daten

Alle Daten bleiben lokal unter:

```text
%LOCALAPPDATA%\Schnellbausteine\Data
```

`catalog.json` enthält nur versionierte Metadaten. Jeder Baustein liegt unter
einer zufälligen ID als eigene XML-Datei in `snippets`; Namen werden nie als
Dateinamen verwendet. Schreibvorgänge erfolgen über eine temporäre Datei im
gleichen Verzeichnis und anschließendes atomisches Ersetzen. Bilder werden als
Base64-Daten im Fragment gespeichert. OneNote-`TagDef`- und
`QuickStyleDef`-Abhängigkeiten werden mitgespeichert und beim Einfügen auf die
Zielseite abgebildet. Neue Fragmente speichern außerdem die aus Outline, Tabelle
oder Bild abgeleitete Mindestbreite kulturunabhängig. Bestehende Fragmente ohne
diese optionale Angabe bleiben lesbar und verändern die Zielbreite nicht.

## Einschränkungen

- Nur OneNote Desktop für Windows; nicht OneNote im Web, macOS oder die alte
  Windows-10-Store-App.
- Keine Cloud-Synchronisierung, Kategorien, KI, Vorlagen oder Platzhalter.
- Eingebettete Dateien sind nicht Ziel des MVP. Formatierter Text,
  Überschriften, Listen, Tabellen, OneNote-Tags/Checkboxen und Bilder werden
  unterstützt.
- Hyperlink-/MathML-Spezialcursor werden absichtlich nur an einer sicheren
  Absatzgrenze behandelt, damit kein Inhalt verloren geht.

## Troubleshooting

- **Setup findet OneNote nicht:** Nur die klassische Desktop-Ausgabe wird
  unterstützt. OneNote einmal starten, wieder vollständig schließen und Setup
  erneut ausführen.
- **Setup meldet, OneNote sei geöffnet:** Auch minimierte OneNote-Fenster
  schließen. Im Task-Manager darf kein `ONENOTE.EXE` mehr laufen.
- **Add-in fehlt:** OneNote neu starten und unter
  `Datei > Optionen > Add-Ins` nachsehen. Bei deaktivierten Elementen
  **Schnellbausteine** wieder aktivieren.
- **LoadBehavior wird 2:** Setup erneut ausführen und
  `%LOCALAPPDATA%\Schnellbausteine\addin.log` prüfen.
- **OneNote reagiert nach einem Befehl nicht:** `addin.log` enthält für Callback,
  STA-Dispatcher, OneNote-COM und Dialoge jeweils `before`/`after`-Phasen sowie
  Prozess, Thread und Apartment. Selection-Diagnosen nennen zusätzlich Scope,
  Anzahl, Knotentypen, strukturelle XML-Pfade und den gewählten Anker. Die letzte
  `before`-Zeile ohne passende `after`-Zeile grenzt den blockierenden Aufruf ein.
  Seiteninhalte, Namen und Binärdaten werden nicht protokolliert.
- **Bilder fehlen:** Den Baustein erneut speichern. Der Speichervorgang bricht
  sichtbar ab, wenn OneNote für ein Callback-Bild keine Binärdaten liefert.
- **Defekter Katalog:** `catalog.json` wird nicht still verworfen. Die
  Fehlermeldung und `addin.log` nennen den betroffenen Pfad; die einzelnen
  Fragmentdateien bleiben erhalten.

## Entwicklung

Zum Bauen werden Visual Studio 2022 Build Tools mit
**.NET-Desktopentwicklung**, .NET Framework 4.8, PowerShell 5.1+ und die mit
OneNote installierten Office-PIAs benötigt.

Setup installiert die offizielle Microsoft-OneNote-PIA lokal neben dem Add-in.
Dadurch verwendet der COM-Zugriff konkrete PIA-Interfaces statt dynamischer
Dispatch-Bindung.

```powershell
.\test.ps1 -Configuration Release -Platform x86
.\test.ps1 -Configuration Release -Platform x64
.\installer\Build-Setup.ps1 -Version 0.1.0
.\tests\Test-OneNoteStartup.ps1 `
    -SetupPath .\artifacts\setup\Schnellbausteine-Setup-0.1.0.exe
```

Die Build-Ausgaben liegen unter `artifacts\<Plattform>\<Konfiguration>`, Setup
unter `artifacts\setup`. `Build-Setup.ps1` lädt den offiziell veröffentlichten,
per SHA-256 fixierten Inno-Setup-Compiler 7.0.2 und baut reproduzierbar beide
Bitnesses in eine EXE. Die GitHub-Actions-Pipeline führt beide Testsuiten sowie
stille Installations-, Managed-Add-in-/Ribbon-, Registry-, Upgrade- und
Deinstallationstests aus. Der saubere GitHub-Windows-Runner enthält kein Office
und kann deshalb keinen echten OneNote- oder Office-COM-Hosttest ersetzen. Der
vollständige DllHost-Lifecycle läuft lokal standardmäßig in `Test-Setup.ps1`;
`Test-OneNoteStartup.ps1` ist der zusätzliche Regressionstest mit einem
tatsächlich installierten OneNote Desktop. OneNote muss davor vollständig
geschlossen sein. Bei einem Tag `v*` wird Setup automatisch an ein GitHub
Release angehängt.

Die PowerShell-Skripte bleiben als Entwickler-Fallback verfügbar:

```powershell
.\installer\Install.ps1 -Platform x64 -Configuration Release
.\installer\Uninstall.ps1 -Platform x64
```

## Architektur

- `Schnellbausteine.Core`: Repository, atomare Persistenz und OneNote-XML-Logik
- `Schnellbausteine.AddIn`: COM-Lebenszyklus, Ribbon, OneNote-PIA und WinForms
- `Schnellbausteine.Tests`: unabhängig ausführbare Unit-Tests
- `installer/Schnellbausteine.iss`: universelles deutsches Inno Setup

Das Add-in nutzt `IDTExtensibility2`, `Microsoft.Office.Core.IRibbonExtensibility`,
Ribbon Custom UI `2009/07` und OneNote XML `XMLSchema.xs2013`. COM-Objekte werden
pro Aktion erzeugt und explizit freigegeben; das von `OnConnection` gelieferte
Application-Objekt wird nicht dauerhaft gehalten. Ribbon-Aktionen geben sofort
an OneNote zurück. OneNote-COM-Aufrufe und WinForms-Dialoge laufen anschließend
auf einem eigenen STA-Thread mit Windows-Message-Pump; Dialoge gehören dabei zum
aktiven OneNote-Hauptfenster. Das einzelne Ribbon-Control ist ein natives
Office-`dynamicMenu`, dessen Inhalt direkt aus alphabetisch sortierten nativen
Buttons, einer nativen Trennlinie und dem abschließenden Manage-Button besteht.
Der synchrone `getContent`-Callback greift nur auf das lokale Repository-Snapshot
zu und erzeugt kein OneNote-COM-Objekt. Erst `button.onAction` stellt den Insert
nach Callback-Rückkehr auf den STA-Dispatcher und liest dort die aktuelle Seite
mit `piBinaryDataSelection`.
Der dabei erfasste `lastModifiedTime`-Wert verhindert, dass zwischenzeitliche
Änderungen aus einem anderen Fenster oder durch Synchronisierung überschrieben
werden; in diesem Fall bricht OneNote das Update sichtbar ab. Das aktualisierte
XML enthält einen validierten leeren `one:T selected="all"`-Cursoranker direkt
nach dem eingefügten Inhalt. Nach einem erfolgreichen Update aktiviert das
Add-in nur noch best-effort das zuvor erfasste OneNote-Hauptfenster. Es verwendet
bewusst keine `NavigateTo`-Editornavigation, da OneNote dabei nach einem Insert
abstürzen kann. Solange OneNote keine sichere Editorfokus-API bereitstellt, kann
zum sofortigen Weitertippen nach einem Ribbon-Insert weiterhin ein Klick auf die
Seite erforderlich sein. Die Containerverbreiterung ändert ausschließlich die
Breite und stellt keinen Editorfokus her. Ein Aktivierungsfehler wird ausschließlich
protokolliert und blockiert den erfolgreichen Insert nicht durch einen modalen
Hinweis.
Damit folgt der Einfügepfad wie OneMores öffentlich beschriebenes
`SelectionRange`-/`PageEditor.InsertOrReplace`-Muster einer expliziten
Selection-Range statt einer globalen Outline-Suche, ohne OneMore-Code zu
übernehmen.

### Experimenteller OneNote-Realtest

1. Cursor an eine klar erkennbare Stelle innerhalb eines Text-Outlines setzen,
   Menü öffnen, einen Baustein anklicken und Position sowie Selection-Diagnose
   prüfen.
2. Den Test in einem zweiten Outline wiederholen und bestätigen, dass niemals in
   das erste Outline oder eine andere Position eingefügt wird.
3. **Schnellbausteine verwalten** öffnen, Bausteine umbenennen und löschen sowie
   einen neuen Baustein speichern; nach jedem CRUD-Vorgang die aktualisierte,
   alphabetischen Menüeinträge und den deaktivierten Leerzustand prüfen.
