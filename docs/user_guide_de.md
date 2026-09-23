# Anleitung: Corporate Circular Roadmap – Chatbot

Diese Anleitung erklärt die aktuelle Arbeitsweise mit dem SiemensGPT-Bot
"Corporate Circular Roadmap Generator".

**Wichtigste Änderung:** Der Bot bearbeitet die Daten nicht mehr selbst.
Alle Änderungen erfolgen direkt in der gemeinsamen Excel-Datei. Der Bot
liest diese Datei nur noch und erzeugt daraus die Roadmap – er ist ein
reines Erzeugungswerkzeug, kein Bearbeitungswerkzeug.

---

## 1. Die zwei Dateien

1. **`generate_roadmap.py`** – das Layout/Design-Skript. Liegt im
   Wissensspeicher ("Knowledge") des Bots und muss normalerweise
   **nicht** jedes Mal neu angehängt werden – nur wenn der Bot
   ausdrücklich danach fragt (z. B. falls dort noch keine oder mehrere
   Versionen liegen).
2. **`roadmap_data.xlsx`** – die eigentlichen Daten, eine Zeile pro
   Rundschreiben. Liegt im gemeinsamen Ablageort (SharePoint/Teams) und
   wird dort von allen im Team direkt gepflegt. **Muss zu Beginn jeder
   neuen Unterhaltung erneut angehängt werden** – das ändert sich
   ständig, deshalb gibt es dafür keinen dauerhaften Speicherplatz beim
   Bot.

---

## 2. Änderungen vornehmen

**Alle Änderungen geschehen direkt in `roadmap_data.xlsx`** – neue
Zeile für ein neues Rundschreiben, Zelle bearbeiten für eine Änderung.
Der Bot wird dafür nicht mehr benötigt. Sobald die Datei aktualisiert
ist, hängen Sie einfach die neueste Version in der nächsten
Bot-Unterhaltung an, um die Roadmap zu erzeugen.

Spalten in der Excel-Datei:

| Spalte | Bedeutung |
|---|---|
| Circular Number | z. B. 1_C_126 |
| Governance Responsibility | z. B. P&O TO. Mehrere Einheiten für ein Rundschreiben mit Semikolon trennen (z. B. "MAP; ST") – jede Einheit erscheint dann auf einer eigenen Zeile |
| Circular Name | Titel des Rundschreibens |
| Description of Changes | Beschreibung – leer lassen, wenn noch nicht bekannt |
| Impact on Existing Circular | Upgrade / Update / Update/Upgrade / New |
| Current Version | Bisherige Version – bei einem neuen Rundschreiben ohne Vorgänger einfach leer lassen |
| Planned Version | Geplante neue Version |
| Contact Circular | Ansprechpartner Rundschreiben |
| Contact Siemens Governance Office | Ansprechpartner Governance |
| Planned Quarter | z. B. Q4/2026 |
| Document URL | Link zum Dokument, optional |
| Internal Comment | interne Notiz, wird nie veröffentlicht |

**Leere Zelle = noch nicht bekannt.** Nichts anderes eintragen (kein
"-", kein "TBD").

---

## 3. Eine Unterhaltung mit dem Bot starten

1. Chat öffnen. Der Bot meldet sich sofort von selbst, ohne dass Sie
   etwas schreiben müssen:

   > Hallo! Bitte lade die aktuelle Roadmap-Excel-Datei
   > (roadmap_data.xlsx) hoch, damit ich loslegen kann.

2. `roadmap_data.xlsx` anhängen (falls noch nicht geschehen).
3. Danach erscheint das Menü:

   > Was möchtest du tun?
   > 1. Preliminary Roadmap erstellen und herunterladen
   > 2. Final Roadmap erstellen und herunterladen
   > 3. Aktuellen Status anzeigen

Sie können auch direkt schreiben, was Sie möchten, z. B. "gib mir die
aktuelle Vorschau" – der Bot erkennt das auch ohne Zahlenauswahl.

Der Bot antwortet immer auf Deutsch, unabhängig davon, in welcher
Sprache Sie selbst schreiben.

---

## 4. Die drei Optionen

**1. Preliminary Roadmap erstellen und herunterladen**
Erzeugt die aktuelle Vorschau-Version, ohne jede Prüfung – einfach
"wie sieht es gerade aus".

**2. Final Roadmap erstellen und herunterladen**
Prüft zuerst, ob alle Rundschreiben vollständig sind (siehe Abschnitt
5), und fragt bei offenen Punkten nach, ob trotzdem veröffentlicht
werden soll. Erst nach Bestätigung wird die finale Version (ohne den
roten "PRELIMINARY"-Hinweis) erzeugt.

**3. Aktuellen Status anzeigen**
Liste aller Rundschreiben, gruppiert nach Quartal, mit dem Status
"vollständig" oder "unvollständig: [betroffene Felder]".

---

## 5. Status verstehen: "vollständig" vs. "unvollständig"

Zwei Punkte, die auf den ersten Blick überraschen können:

- **Eine generische Beschreibung zählt als unvollständig**, auch wenn
  im Dokument ein Satz steht. Ist "Description of Changes" leer, aber
  "Impact" bekannt, erzeugt das System automatisch einen allgemeinen
  Satz wie "This Circular will be upgraded." – sieht im fertigen
  Dokument gut aus, gilt aber weiterhin als nicht vollständig, weil es
  noch keine echte, spezifische Beschreibung ist.
- **"Update/Upgrade" bei Impact zählt als unvollständig.** Das ist eine
  offene Entweder-oder-Angabe. Erst "Upgrade", "Update" oder "New"
  allein gilt als abgeschlossene Entscheidung.

Beide Punkte werden unabhängig voneinander geprüft.

---

## 6. Layout-Hinweis: leichte Quartale

Alle Seiten sind immer im Hochformat, unabhängig davon, wie viele
Rundschreiben ein Quartal hat. Hat ein Quartal nur ein oder zwei
Rundschreiben, wird die Zeilenhöhe trotzdem angemessen gewählt und die
Tabelle auf der Seite zentriert – keine einzelne Zeile, die über die
ganze Seite gezogen wird, und keine winzige Tabelle, die oben verloren
wirkt. Das passiert automatisch, dafür ist nichts zu tun.

---

## 7. Praktische Hinweise

- **Immer die neueste Excel-Datei anhängen.** Nach jeder Änderung in
  der Excel-Datei muss die aktuelle Version erneut an den Bot
  angehängt werden, auch innerhalb derselben Unterhaltung – der Bot
  verwendet immer nur die zuletzt angehängte Datei.
- **Kein PDF direkt aus dem Chat.** Der Bot liefert HTML. Für ein PDF:
  Datei im Browser öffnen, dort "Drucken" → "Als PDF speichern".
- **Download-Link erscheint nicht anklickbar?** Schreiben Sie: "Gib
  mir den echten Link als klickbaren Link, z. B. [Dateiname](URL)."
- **Neue Unterhaltung = neuer Start für die Excel-Datei.** Die
  `roadmap_data.xlsx` muss jedes Mal erneut angehängt werden; das
  Skript normalerweise nicht (liegt im Knowledge-Speicher, siehe
  Abschnitt 1).
- **Der Bot bearbeitet keine Daten mehr.** Für jede inhaltliche
  Änderung: direkt in `roadmap_data.xlsx` eintragen, dann die
  aktualisierte Datei beim nächsten Mal anhängen.
