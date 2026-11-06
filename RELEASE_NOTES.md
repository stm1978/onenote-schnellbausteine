# OneNote Schnellbausteine – Beta 0.2.3

## v0.2.3-beta.1

- Trennt das Speichern einer Umbenennung von späteren Listen-/Ribbon-Refreshes, damit nachgelagerte Fehler nicht mehr als fehlgeschlagene Umbenennung gemeldet werden; mehrere aufeinanderfolgende Renames sind abgesichert.
- Ink-Positionen werden beim Speichern als relative Abstände der gemeinsamen Auswahl bewahrt und beim Einfügen gemeinsam am Ziel-Outline verankert.
- Ink-Binärdaten sowie vorhandene Darstellungsattribute und -elemente werden beim XML-Roundtrip erhalten.
- Entfernt Ablaufdatum, Warnhinweise, Download-Dialog und zeitgesteuerte Funktionssperren.
- Automatisierte XML-Tests decken Ink-Binärpayloads, Rendering-XML und relative Ink-Positionen ab. Die reale Regenbogenstift-Darstellung und Positionierung müssen noch manuell in OneNote geprüft werden.
