# OneNote Schnellbausteine

**Beta 0.2.5** von **Stefan Malter | Malter365.de**

OneNote Schnellbausteine ist ein bewusst kleines, lokales COM-Add-in für **OneNote
Desktop unter Windows** (Microsoft 365 bzw. OneNote 2016). Markierte
OneNote-Inhalte lassen sich als benannte Bausteine speichern und später über das
Menüband auf der Registerkarte **Einfügen** an der Cursorposition wieder einfügen.

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

## Beta 0.2.5 herunterladen und installieren

> [!CAUTION]
> **Beta 0.2.5 vorerst nicht zum Einfügen verwenden.** Ein Tester meldet, dass
> nach dem Einfügen auch vorhandene Originalzeichnungen ihre Rainbow-Farbe
> verlieren. Der bisherige XML-Pfad schreibt die gesamte Seite zurück.
> Der noch unveröffentlichte Stand schreibt nur den eindeutig identifizierten
> Ziel-Textcontainer und neue Formatdefinitionen zurück. Andere Seitenobjekte
> werden nicht zurückgesendet. Freihand-Bausteine und Zielcontainer mit
> Freihanddaten bleiben gesperrt.
> Diese Änderungen sind im veröffentlichten Installer nicht enthalten und
> stellt keine Reparatur der Ink-Darstellung dar.
> Zudem blockiert Microsoft Defender den veröffentlichten Installer als
> `Program:Win32/Contebrew.A!ml`. Die Erkennung ist nicht geklärt;
> Schutzmaßnahmen nicht umgehen und den Installer nicht weitergeben.

> [!IMPORTANT]
> Dies ist eine öffentliche Testversion. Bitte wichtige OneNote-Notizbücher wie
> gewohnt synchronisieren bzw. sichern.

1. Im Release
   [**OneNote Schnellbausteine – Beta 0.2.5**](https://github.com/stm1978/onenote-schnellbausteine/releases/tag/v0.2.5-beta.1)
   die Datei **`OneNote-Schnellbausteine-Setup-Beta-0.2.5.exe`** herunterladen.
2. OneNote vollständig schließen.
3. Die heruntergeladene Setup-Datei doppelklicken und **Installieren** wählen.
4. OneNote neu starten.

Das universelle Setup erkennt automatisch, ob OneNote 32 oder 64 Bit verwendet, und installiert
die passende Add-in-Version ausschließlich für das aktuelle Windows-Konto.
Administratorrechte und PowerShell-Befehle sind nicht erforderlich.

**Voraussetzungen:** Windows 10/11, .NET Framework 4.8 und OneNote Desktop aus
Microsoft 365 oder OneNote 2016. ARM64 ist im MVP nicht enthalten.

Die Deinstallation erfolgt wie bei anderen Windows-Programmen unter
**Einstellungen > Apps > Installierte Apps > OneNote Schnellbausteine**.
Gespeicherte Schnellbausteine unter
`%LOCALAPPDATA%\Schnellbausteine\Data` bleiben bei der Deinstallation
standardmäßig erhalten.

Diese Version hat kein Ablaufdatum und sperrt ihre Funktionen nicht zeitgesteuert.
Lokale Schnellbausteine werden durch Upgrade oder Deinstallation nicht gelöscht.

> Die Vorabversion ist noch nicht digital signiert. Die fehlende Signatur ist
> keine Erklärung oder Entwarnung für die aktuelle Defender-Erkennung.

## Bedienung

Im noch unveröffentlichten Speicherpfad wird vor dem Namensdialog nur die
Seitenstruktur einschließlich Auswahlmarkierungen (`piSelection`) aufgenommen.
Erst nach Bestätigung des Namens lädt das Add-in die Binärdaten der ausgewählten
Bilder und vollständigen Freihandobjekte nach. Abbrechen lädt keine Binärdaten.
Die Auswahl wird nach dem Dialog nicht erneut gelesen.

1. Inhalt auf einer OneNote-Seite markieren.
2. Im Kontextmenü von Text, Outline-Handle, Tabelle oder Freihandauswahl
   **Als Schnellbaustein speichern...** wählen und einen eindeutigen Namen
   eingeben. Bei Namenskonflikten wird explizit nachgefragt, ob überschrieben
   oder ein anderer Name verwendet werden soll.
3. Im Register **Einfügen** in der Gruppe **Malter365.de** auf
   **Schnellbausteine** klicken und den gewünschten,
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

Der unveröffentlichte Einfügepfad arbeitet auf einer Kopie der gelesenen Seite.
An OneNote werden ausschließlich der vollständige geänderte Ziel-Outline mit
seiner bestehenden `objectID` und neue `TagDef`-/`QuickStyleDef`-Definitionen
gesendet. Unveränderte Originalzeichnungen, Bilder, Titel, andere Outlines und
unbekannte Seitendaten bleiben außerhalb dieses Updates. Änderungen außerhalb
des Zielcontainers, fehlende oder doppelte Objekt-IDs sowie Änderungen an
bestehenden Formatdefinitionen führen zum sichtbaren Abbruch. OneNotes
Änderungszeitpunkt-Prüfung bleibt aktiv. Dieser Pfad ist automatisiert am
COM-Adapter geprüft, aber noch nicht im realen OneNote validiert.

Menü, Tastatursteuerung, Escape und Click-away werden vollständig von Office
bereitgestellt; das Add-in öffnet keine eigene Popup-Oberfläche.

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
Zielseite abgebildet. Neue Fragmente speichern außerdem die aus Outline, Tabelle,
Bild oder Freihandzeichnung abgeleitete Mindestbreite kulturunabhängig.
Vollständig markierte Freihandzeichnungen werden mit `Size`, `ShapeInfo`,
`inkOrigin` und opaken Base64-Daten gespeichert. Eine `callbackID` wird vor dem
Speichern über OneNote materialisiert; externe `File`-Pfade werden nie
persistiert. Bestehende Fragmente ohne
diese optionale Angabe bleiben lesbar und verändern die Zielbreite nicht.

## Einschränkungen

- Nach dem Einfügen kann ein Klick auf die OneNote-Seite erforderlich sein, bevor
  direkt weitergetippt werden kann.
- Nur OneNote Desktop für Windows; nicht OneNote im Web, macOS oder die alte
  Windows-10-Store-App.
- Keine Cloud-Synchronisierung, Kategorien, KI, Vorlagen oder Platzhalter.
- Eingebettete Dateien sind nicht Ziel des MVP. Formatierter Text,
  Überschriften, Listen, OneNote-Tags/Checkboxen, Bilder, ganze Tabellen,
  vollständige Tabellenzeilen/-spalten, rechteckige Zellbereiche und vollständige
  Freihandzeichnungen werden unterstützt. Eindeutig getrennte Auswahlwurzeln aus
  Text, Bild oder genau einer Tabelle können gemeinsam mit einer oder mehreren
  vollständigen Freihandzeichnungen dokumentgeordnet gespeichert werden.
- Mehrere Tabellen, nicht rechteckige Zellbereiche, Tabellen mit extern
  markierten Inhalten, überlappende Auswahlwurzeln sowie partielle oder durch
  vollständig markierte Seitenvorfahren mehrdeutige Freihandauswahlen werden
  sicher abgelehnt.
- Eine Lasso-Auswahl wird nicht offline auf einzelne Striche zugeschnitten.
  Gespeichert wird ausschließlich ein vollständiger `InkDrawing`-Blob, den
  OneNote selbst als eigenständige Binärdaten
  liefert. Fehlen diese Daten, bricht der Speichervorgang sichtbar ab.
- Tester melden für Beta 0.2.5 weiterhin falsche Anordnung zusammengesetzter
  Zeichnungen und einfarbige Rainbow-Striche. Die XML-Roundtrip-Tests belegen
  keine originalgetreue Darstellung in OneNote; diese Fehler sind noch offen.
- Als Schutz vor Farbverlust verweigert der unveröffentlichte Add-in-Stand
  Einfügen, sobald der Ziel-Textcontainer oder Baustein `InkDrawing` enthält.
  Freihand in anderen Containern oder direkt auf der Seite wird nicht erneut
  geschrieben und sperrt einen getrennten Textcontainer nicht mehr.
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
- **OneNote reagiert nach einem Befehl nicht:** OneNote vollständig schließen und
  neu starten. Bleibt das Problem bestehen, bitte ein
  [GitHub-Issue](https://github.com/stm1978/onenote-schnellbausteine/issues) erstellen.
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
.\installer\Build-Setup.ps1 -Version 0.2.5-beta.1
.\tests\Test-OneNoteStartup.ps1 `
    -SetupPath .\artifacts\setup\OneNote-Schnellbausteine-Setup-Beta-0.2.5.exe
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

Die automatisierten XML-Roundtrip-Tests prüfen die Bewahrung der Ink-Daten und
Rendering-Metadaten sowie die relative Position mehrerer Striche. Ein echter
OneNote-Desktop-Test ist davon nicht ersetzt: Regenbogenstift-Darstellung und
räumliche Positionierung müssen in OneNote manuell validiert werden.

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
aktiven OneNote-Hauptfenster. Das Control **Schnellbausteine** in der Gruppe
**Malter365.de** auf der eingebauten Registerkarte **Einfügen** und die gleichnamigen
Untermenüs in den Text-, Outline-,
Freihand- und Lasso-Kontextmenüs sind native Office-`dynamicMenu`-Controls. Ihr
Inhalt besteht direkt aus alphabetisch sortierten nativen Buttons, einer nativen
Trennlinie und dem abschließenden Manage-Button. Jedes geöffnete Menü besitzt
einen eigenen Snapshot; ein Refresh einer anderen Oberfläche macht seine zufälligen
Item-IDs nicht ungültig. Save, Rename und Delete invalidieren alle Snapshots.
Der synchrone `getContent`-Callback greift nur auf das lokale Repository zu und
erzeugt kein OneNote-COM-Objekt. Erst `button.onAction` stellt den Insert
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

Beim Erfassen von Freihand werden die seitenbezogenen Ink-Positionen in ein
gemeinsames relatives Koordinatensystem überführt und beim Einfügen gemeinsam
am Ziel-Outline verankert. Die Ink-Binärdaten und Rendering-XML bleiben dabei
erhalten. Die tatsächliche Darstellung des Regenbogenstifts und die Positionen
der Striche müssen zusätzlich in OneNote manuell geprüft werden.

### Experimenteller OneNote-Realtest

1. Cursor an eine klar erkennbare Stelle innerhalb eines Text-Outlines setzen,
   Menü öffnen, einen Baustein anklicken und Position sowie Selection-Diagnose
   prüfen.
2. Den Test in einem zweiten Outline wiederholen und bestätigen, dass niemals in
   das erste Outline oder eine andere Position eingefügt wird.
3. **Schnellbausteine verwalten** öffnen, Bausteine umbenennen und löschen sowie
   einen neuen Baustein speichern; nach jedem CRUD-Vorgang die aktualisierte,
   alphabetischen Menüeinträge und den deaktivierten Leerzustand prüfen.
4. Ganze Tabelle, vollständige Zeilen, vollständige Spalten und einen
   zusammenhängenden rechteckigen Zellbereich jeweils speichern und an einem
   Text- sowie Zellanker einfügen; Struktur, Zellformatierung, Inhalte, Bilder,
   Tags und Spaltenbreiten vergleichen.
5. Je eine vollständige Stift-, Maus- und Lasso-Auswahl speichern und wieder
   einfügen. Für jede Auswahl Typ und Anzahl der gelieferten `InkDrawing`-Objekte
   sowie Binärlänge und Hash vor und nach dem Roundtrip vergleichen.
6. Bei einer Lasso-Auswahl über Teilstriche ausdrücklich prüfen, ob OneNote
   eigenständige vollständige Binärdaten liefert. Ohne solchen Blob muss das
   Add-in ablehnen; ein Cropping einzelner Striche wird nicht behauptet.
