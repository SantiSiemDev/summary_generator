# Benutzerhandbuch – Summary of Regulatory Changes

Sprache: [English](user_manual_en.md) | Deutsch

## 1. Was der Generator macht

Der Generator erstellt das quartalsweise „Summary of regulatory changes":
ein Deckblatt mit Übersicht (Circulars und Standards, jeweils unterteilt
in New, Upgrade, Archivings und Update), danach Detailtabellen mit
englischen und deutschen Titeln, Links und Zusammenfassungen.

Er liest nur die Arbeitsmappe `bus_data.xlsx`, vor allem das Blatt
`Summary`, und verändert sie nie. Das Ergebnis ist eine HTML-Datei, die
du im Browser als PDF druckst.

## 2. Was du brauchst

- Die aktuelle `bus_data.xlsx` (Format `.xlsx`; Dateien mit Makros, `.xlsm`, können nicht gelesen werden).
- Entweder den SiemensGPT-Bot (`generate_summary.py` liegt in seinem Knowledge) oder Python 3 mit `openpyxl` für die lokale Nutzung.
- Chrome oder Edge zum Speichern als PDF.

## 3. Arbeitsmappe vorbereiten

Die Arbeitsmappe hat drei Blätter:

| Blatt | Zweck |
|---|---|
| Roadmap | Geplante Änderungen je Circular inklusive Planned Quarter. Wird separat gepflegt, der Generator liest nur. |
| Summary | Das, was im Dokument erscheint. Eine Zeile pro Regelung. |
| Masterliste | Referenzliste aller Dokumente (Nummer, Version, Org-Einheit, Ansprechpartner, Titel, Links). |

### Spalten im Blatt Summary

| Spalte | Inhalt |
|---|---|
| Type | `Circular` oder `Standard`. |
| Change | `New`, `Upgrade`, `Update` oder `Archiving`. |
| Governance responsibility | Verantwortliche Einheit. |
| Document Number | Zum Beispiel `1_C_126` oder `1_ST_18`. Die Überschrift `Number` funktioniert ebenfalls. |
| Current Version / Planned Version | Zwei getrennte Spalten. |
| Title DE, Link DE, Summary DE | Optional. Bleiben sie leer, wird nur die englische Zeile angezeigt. |
| Title EN, Link EN, Summary EN | Englischer Titel, Link und Zusammenfassung. |

Die Spaltenüberschriften werden über den Namen erkannt, nicht über die
Position. Groß- und Kleinschreibung und zusätzliche Leerzeichen spielen
keine Rolle. Benenne die Spalten nicht um.

### Automatisches Ausfüllen

Wenn die Formeln vorhanden sind, tippst du nur die Document Number ein.
Type, Change, Governance responsibility, Versionen, Titel und Links
werden aus Roadmap und Masterliste ergänzt.

- Ein Circular, das in der Roadmap steht, übernimmt seinen Change von dort.
- Ein Circular, das nicht in der Roadmap steht, wird zu `Archiving`.
- Ein Standard steht nicht in der Roadmap, sein Change bleibt deshalb leer. Trag ihn von Hand ein (`New`, `Upgrade`, `Update` oder `Archiving`).
- Summary DE und Summary EN schreibst du immer selbst. Auch die Planned Version musst du eventuell von Hand eintragen, zum Beispiel bei einem Standard.
- Speichere die Datei nach dem Bearbeiten in Excel. Die Formeln berechnet Excel, der Generator liest die gespeicherten Ergebnisse. Eine Datei aus einem anderen Programm enthält unter Umständen keine Ergebnisse.

### So werden Versionen angezeigt

| Change | Anzeige im Dokument |
|---|---|
| Update, Upgrade | `2.3 → 3.0` (Current → Planned) |
| New | Nur die Planned Version (die Current Version, falls Planned leer ist) |
| Archiving | Nur die archivierte Version |

Versionen haben immer einen Dezimalpunkt (aus `3` wird `3.0`).

### Regeln

- `Update/Upgrade` ist nicht zulässig. Ist die Entscheidung noch offen, kann die Zeile nicht veröffentlicht werden. Der Generator bricht ab und nennt die Zeile.
- Eine leere Zelle bedeutet „noch nicht bekannt" und bleibt im Dokument leer. Trag keine Platzhalter wie `TBD` oder `-` ein.
- Archivings ohne Zusammenfassung bekommen automatisch einen Standardsatz. Willst du ein Nachfolgedokument nennen, schreib deinen eigenen Text in Summary EN und Summary DE.
- Summary soll nur das Quartal enthalten, das du gerade vorbereitest. Leere die alten Zeilen, wenn du ein neues Quartal beginnst.

## 4. Welches Quartal

Der Generator wählt das Quartal automatisch: das früheste Quartal in der
Roadmap, Spalte Planned Quarter, das noch nicht vorbei ist. Um ein
bestimmtes Quartal zu erzwingen, nenne es in deiner Anfrage an den Bot,
übergib es als Kommandozeilenargument oder trag den Schlüssel `Quarter`
in einem Blatt `Meta` ein.

Akzeptierte Formate: `Q4/2026`, `Q4-2026`, `Q4 2026` (zuerst das Quartal, dann ein vierstelliges Jahr).
Nicht akzeptiert: `2026 Q4`, `Q4 FY26`.

Das Quartal bestimmt nur den Titel des Dokuments („Changes to Circulars in Q4/2026"). Alle Zeilen aus Summary werden übernommen.

## 5. Dokument erzeugen

### Mit dem Bot

1. Hänge die aktuelle `bus_data.xlsx` an den Chat an.
2. Frag nach der Summary, zum Beispiel „Erstelle die Summary für dieses Quartal". Nenne ein Quartal nur, wenn du ein bestimmtes willst.
3. Der Bot antwortet auf Deutsch mit der HTML-Datei und der Anzahl der Einträge je Kategorie.

### Lokal

```bash
pip install openpyxl
python src/generate_summary.py bus_data.xlsx [out_dir] [quarter]
```

`out_dir` ist standardmäßig der aktuelle Ordner. Beispiel mit Quartal: `python src/generate_summary.py bus_data.xlsx . Q1/2027`.

## 6. Als PDF speichern

1. Öffne die HTML-Datei in Chrome oder Edge.
2. Drücke Strg+P (Windows) bzw. Cmd+P (Mac).
3. Ziel: Als PDF speichern. Papierformat: A4. Ränder: Keine.
4. Schalte „Hintergrundgrafiken" ein. Ohne diese Option gehen der dunkelblaue Hintergrund und die Farben verloren.
5. Speichern.

## 7. Das Dokument lesen

Das Deckblatt zeigt je ein 2x2-Raster für Circulars und für Standards:
New (links oben), Upgrade (rechts oben), Archivings (links unten),
Update (rechts unten). Ein Kasten ohne Einträge zeigt „Nothing to report
this quarter". Ab Seite 2 hat jede Regelung eine Tabellenzeile mit
Governance responsibility, Nummer, Version, Titel, Link und
Zusammenfassung, zuerst Englisch, darunter Deutsch.

| Farbe | Bedeutung |
|---|---|
| Grün | New |
| Violett | Upgrade |
| Gold | Update |
| Grau | Archiving |
| Petrol-Rahmen | Circulars |
| Blauer Rahmen | Standards |

Eine Legende wiederholt das ab Seite 2 auf jeder Seite.

## 8. Fehlerbehebung

| Meldung oder Symptom | Ursache und Lösung |
|---|---|
| `<Nummer>: Change is ''` | Die Change-Zelle ist leer. Typisch bei Standards. Trag den Change von Hand ein. |
| `Change is 'Update/Upgrade'` | Entscheidung noch offen. Trag `Update` oder `Upgrade` ein. |
| `has no column recognizable as the regulation number` | Die Überschrift der Nummernspalte ist weder `Document Number` noch `Number`. Prüf die Überschriften. |
| `No quarter could be determined` | Planned Quarter ist leer oder im falschen Format. Nutze `Q4/2026` oder nenne das Quartal ausdrücklich. |
| `No regulations found for <Quartal>` | Summary enthält keine Zeilen. |
| Leere Zellen im Dokument, obwohl Excel Werte zeigt | Die Datei wurde nach der Formeländerung nicht in Excel gespeichert. Öffnen, speichern und neu anhängen. |
| Weiße Seite ohne Farben im PDF | „Hintergrundgrafiken" war beim Drucken ausgeschaltet. |
| Bot findet das Skript nicht | `generate_summary.py` fehlt im Knowledge. Hänge es an den Chat an. |

## 9. Quartalswechsel

1. Speichere eine Kopie der fertigen Arbeitsmappe als Ablage.
2. Leere die alten Zeilen in Summary (die Formeln bleiben).
3. Aktualisiere die Roadmap für das nächste Quartal.
4. Trag die neuen Document Numbers ein und ergänze Change, Versionen und Zusammenfassungen.

## 10. Für Maintainer

- `src/generate_summary.py` enthält Design, Layout und Logik. Die Seitenaufteilung wird in Python anhand der Textlänge berechnet. Prüf nach Änderungen an Schriftarten oder Spaltenbreiten ein gedrucktes Ergebnis.
- Der Bot muss das Skript mit `exec(open("generate_summary.py", encoding="utf-8").read())` laden, nie mit `import`.
- Eine PDF-Erzeugung ist bewusst nicht eingebaut. Siehe „Design notes" in `README.md`.
- Bot-Anweisungen: `docs/bot_instructions.md`.
