# OneNote Schnellbausteine – Beta 0.2.5

## Unveröffentlicht

- Der Speicherpfad liest vor dem Namensdialog nur die Seitenstruktur mit Auswahlmarkierungen (`piSelection`). Erst nach Bestätigung des Namens werden Binärdaten der ausgewählten Bilder und vollständigen Ink-Objekte geladen; ein Abbruch lädt keine Binärdaten. Die Seiten-/Auswahlaufnahme wird nach dem Dialog nicht wiederholt.
- Ink-Binärdatenreferenzen über `objectID` werden vor dem Entfernen seitenbezogener IDs aufgelöst. Regressionstests prüfen die Beschränkung auf ausgewählte Objekte.
- Die reale Dialogwartezeit wurde noch nicht in OneNote gemessen. Die Tester melden weiterhin falsche Anordnung zusammengesetzter Zeichnungen und einfarbige Rainbow-Striche; diese Fehler sind nicht als behoben anzusehen.

## v0.2.5-beta.1

- Ermöglicht das getrennte Speichern ausgewählter Zeichnungen und Texte, auch wenn sie in derselben OneNote-OE-Struktur liegen.
- Erfasst vollständig ausgewählte Ink-Gruppen unter einer OE-Wurzel und bewahrt die relativen Positionen beim gemeinsamen Einfügen.
- Ergänzt Laufzeitmessungen für Seitendatenabruf, XML-Verarbeitung, Fragmentladen und OneNote-Update, um langsame Einfügevorgänge gezielt einzugrenzen.

## v0.2.4-beta.1

- Behebt den Fehler beim Wiederherstellen der Listenauswahl nach einer Umbenennung; die Liste wird nicht mehr während einer Enumeration geändert.
- Ergänzt einen Regressionstest für aufeinanderfolgende Umbenennungen mit wechselnder Sortierposition.

## v0.2.3-beta.1

- Trennt das Speichern einer Umbenennung von späteren Listen-/Ribbon-Refreshes, damit nachgelagerte Fehler nicht mehr als fehlgeschlagene Umbenennung gemeldet werden; mehrere aufeinanderfolgende Renames sind abgesichert.
- Ink-Positionen werden beim Speichern als relative Abstände der gemeinsamen Auswahl bewahrt und beim Einfügen gemeinsam am Ziel-Outline verankert.
- Ink-Binärdaten sowie vorhandene Darstellungsattribute und -elemente werden beim XML-Roundtrip erhalten.
- Entfernt Ablaufdatum, Warnhinweise, Download-Dialog und zeitgesteuerte Funktionssperren.
- Automatisierte XML-Tests decken Ink-Binärpayloads, Rendering-XML und relative Ink-Positionen ab. Die reale Regenbogenstift-Darstellung und Positionierung müssen noch manuell in OneNote geprüft werden.
