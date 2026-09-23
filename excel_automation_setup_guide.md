# Excel Automation Setup Guide / Handbuch zur Einrichtung der Excel-Automatisierung

This manual explains how to restore the automated formulas in the `bus_data.xlsx` workbook. It contains both an English section (using commas and English functions like `XLOOKUP`) and a German section (using semicolons and German functions like `XVERWEIS`).

---

# 🇬🇧 ENGLISH VERSION (For English Excel)

**Note:** The formulas below use the `XLOOKUP` function, which requires Microsoft Excel 365 or Excel 2021 (or newer). 

## 1. Summary Sheet Setup

The automation in the Summary Sheet relies on the **Document Number** being entered in **Column D**. Copy the formulas below and paste them into the corresponding columns in **Row 2**.

| Column Name | Target Cell | Formula to Paste |
| :--- | :--- | :--- |
| **Type** | A2 | `=IF(LEFT(D2, 4)="1_C_", "Circular", "Standard")` |
| **Governance Responsibility** | C2 | `=XLOOKUP(D2, RegList!A:A, RegList!C:C, "")` |
| **Current Version** | E2 | `=XLOOKUP(D2, RegList!A:A, RegList!B:B, "")` |
| **Planned Version** | F2 | `=XLOOKUP(D2, Roadmap!A:A, Roadmap!G:G, "")` |
| **Title DE** | G2 | `=XLOOKUP(D2, RegList!A:A, RegList!E:E, "")` |
| **Link DE** | H2 | `=XLOOKUP(D2, RegList!A:A, RegList!F:F, "")` |
| **Title EN** | J2 | `=XLOOKUP(D2, RegList!A:A, RegList!G:G, "")` |
| **Link EN** | K2 | `=XLOOKUP(D2, RegList!A:A, RegList!H:H, "")` |


## 2. Roadmap Sheet Setup

The automation in the Roadmap Sheet relies on the **Circular Number** being entered in **Column A**. Copy the formulas below and paste them into the corresponding columns in **Row 2**.

| Column Name | Target Cell | Formula to Paste |
| :--- | :--- | :--- |
| **Governance Responsibility** | B2 | `=XLOOKUP(A2, RegList!A:A, RegList!C:C, "")` |
| **Circular Name** | C2 | `=XLOOKUP(A2, RegList!A:A, RegList!G:G, "")` |
| **Current Version** | F2 | `=XLOOKUP(A2, RegList!A:A, RegList!B:B, "")` |
| **Contact Circular** | H2 | `=XLOOKUP(A2, RegList!A:A, RegList!D:D, "")` |
| **Document URL** | K2 | `=XLOOKUP(A2, RegList!A:A, RegList!H:H, "")` |

---

# 🇩🇪 DEUTSCHE VERSION (Für deutsches Excel)

**Hinweis:** Die untenstehenden Formeln verwenden Semikolons (`;`) als Trennzeichen und die Funktion `XVERWEIS`, die Microsoft Excel 365 oder Excel 2021 (oder neuer) erfordert.

## 1. Einrichtung des Tabellenblatts "Summary Sheet"

Die Automatisierung im "Summary Sheet" basiert darauf, dass die **Document Number** in **Spalte D** eingegeben wird. Kopieren Sie die untenstehenden Formeln und fügen Sie diese in die entsprechenden Spalten in **Zeile 2** ein.

| Spaltenname | Zielzelle | Einzufügende Formel |
| :--- | :--- | :--- |
| **Type** | A2 | `=WENN(LINKS(D2; 4)="1_C_"; "Circular"; "Standard")` |
| **Governance Responsibility** | C2 | `=XVERWEIS(D2; RegList!A:A; RegList!C:C; "")` |
| **Current Version** | E2 | `=XVERWEIS(D2; RegList!A:A; RegList!B:B; "")` |
| **Planned Version** | F2 | `=XVERWEIS(D2; Roadmap!A:A; Roadmap!G:G; "")` |
| **Title DE** | G2 | `=XVERWEIS(D2; RegList!A:A; RegList!E:E; "")` |
| **Link DE** | H2 | `=XVERWEIS(D2; RegList!A:A; RegList!F:F; "")` |
| **Title EN** | J2 | `=XVERWEIS(D2; RegList!A:A; RegList!G:G; "")` |
| **Link EN** | K2 | `=XVERWEIS(D2; RegList!A:A; RegList!H:H; "")` |


## 2. Einrichtung des Tabellenblatts "Roadmap Sheet"

Die Automatisierung im "Roadmap Sheet" basiert darauf, dass die **Circular Number** in **Spalte A** eingegeben wird. Kopieren Sie die untenstehenden Formeln und fügen Sie diese in die entsprechenden Spalten in **Zeile 2** ein.

| Spaltenname | Zielzelle | Einzufügende Formel |
| :--- | :--- | :--- |
| **Governance Responsibility** | B2 | `=XVERWEIS(A2; RegList!A:A; RegList!C:C; "")` |
| **Circular Name** | C2 | `=XVERWEIS(A2; RegList!A:A; RegList!G:G; "")` |
| **Current Version** | F2 | `=XVERWEIS(A2; RegList!A:A; RegList!B:B; "")` |
| **Contact Circular** | H2 | `=XVERWEIS(A2; RegList!A:A; RegList!D:D; "")` |
| **Document URL** | K2 | `=XVERWEIS(A2; RegList!A:A; RegList!H:H; "")` |

---

# 🔄 How to apply formulas / Formeln anwenden

1. Highlight all the cells in Row 2 where you just pasted the formulas. / *Markieren Sie alle Zellen in Zeile 2, in die Sie gerade die Formeln eingefügt haben.*
2. Hover over the bottom-right corner of the highlighted cells until your cursor turns into a small black cross (**+**), known as the fill handle. / *Bewegen Sie den Mauszeiger über die untere rechte Ecke der markierten Zellen, bis der Cursor zu einem kleinen, schwarzen Kreuz (**+**) wird (Ausfüllkästchen).*
3. Double-click the fill handle to copy the formulas down automatically, OR click and drag it down as far as you need. / *Doppelklicken Sie auf das Ausfüllkästchen, um die Formeln automatisch nach unten zu kopieren, ODER klicken und ziehen Sie es so weit nach unten, wie Sie es benötigen.*

# 🛠️ Troubleshooting / Fehlerbehebung

* **Cells appearing blank? / Zellen erscheinen leer?** 
  This is intentional. The formulas end with `""` which tells Excel to leave the cell blank if no Document/Circular Number has been entered yet. / *Das ist beabsichtigt. Die Formeln enden mit `""`, was Excel anweist, die Zelle leer zu lassen, solange noch keine Nummer eingegeben wurde.*
* **Formula displaying as text? / Formel wird als Text angezeigt?** 
  Ensure the cells are formatted as "General" and not "Text". / *Stellen Sie sicher, dass die Zellen als "Standard" und nicht als "Text" formatiert sind.*