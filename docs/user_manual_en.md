# User Manual – Summary of Regulatory Changes

Language: English | [Deutsch](user_manual_de.md)

## 1. What it does

The generator produces the quarterly "Summary of regulatory changes": a
cover page with an overview (Circulars and Standards, each split into
New, Upgrade, Archivings and Update) followed by detail tables with
English and German titles, links and summaries.

It reads only the workbook `bus_data.xlsx`, mainly the sheet `Summary`.
It never changes the workbook. The result is an HTML file, which you
print to PDF in your browser.

## 2. What you need

- The current `bus_data.xlsx` (format `.xlsx`; macro-enabled `.xlsm` files cannot be read).
- Either the SiemensGPT bot (`generate_summary.py` is stored in its Knowledge) or Python 3 with `openpyxl` for local use.
- Chrome or Edge to save the PDF.

## 3. Preparing the workbook

The workbook has three sheets:

| Sheet | Purpose |
|---|---|
| Roadmap | Planned changes per Circular, including the Planned Quarter. Maintained separately; the generator only reads it. |
| Summary | What appears in the document. One row per regulation. |
| Masterliste | Reference list of all documents (number, version, org unit, contact, titles, links). |

### Columns of the sheet Summary

| Column | Content |
|---|---|
| Type | `Circular` or `Standard`. |
| Change | `New`, `Upgrade`, `Update` or `Archiving`. |
| Governance responsibility | Responsible unit. |
| Document Number | For example `1_C_126` or `1_ST_18`. A header called `Number` also works. |
| Current Version / Planned Version | Two separate columns. |
| Title DE, Link DE, Summary DE | Optional. If empty, only the English row is shown. |
| Title EN, Link EN, Summary EN | English title, link and summary. |

Column headers are matched by name, not by position. Capitalization and
extra spaces do not matter. Do not rename them.

### Autofill

If the formulas are in place, you only type the Document Number. Type,
Change, Governance responsibility, versions, titles and links are then
filled from Roadmap and Masterliste.

- A Circular that is in Roadmap takes its Change from there.
- A Circular that is not in Roadmap becomes `Archiving`.
- A Standard is not in Roadmap, so its Change stays empty. Type it by hand (`New`, `Upgrade`, `Update` or `Archiving`).
- Summary DE and Summary EN are always typed by hand. Planned Version may also need a manual entry, for example for a Standard.
- After editing, save the file in Excel. The formulas are calculated by Excel, and the generator reads the saved results. A file produced by another program may contain no results.

### How versions are shown

| Change | Shown in the document |
|---|---|
| Update, Upgrade | `2.3 → 3.0` (Current → Planned) |
| New | Planned Version only (Current if Planned is empty) |
| Archiving | The archived version only |

Versions always carry a decimal point (`3` becomes `3.0`).

### Rules

- `Update/Upgrade` is not allowed. If the decision is still open, the row cannot be published. The generator stops and names the row.
- A blank cell means "not known yet". It stays blank in the document. Do not type placeholders such as `TBD` or `-`.
- Archivings with empty summaries get a standard sentence automatically. If you want to name a successor document, type your own text in Summary EN and Summary DE.
- Summary should contain only the quarter you are preparing. Clear the old rows when you start a new quarter.

## 4. Which quarter

The generator chooses the quarter automatically: the earliest quarter in
Roadmap, column Planned Quarter, that is not yet in the past. To force a
quarter, name it in your request to the bot or pass it as a command
argument, or set the key `Quarter` in a sheet called `Meta`.

Accepted formats: `Q4/2026`, `Q4-2026`, `Q4 2026` (quarter first, then a 4-digit year).
Not accepted: `2026 Q4`, `Q4 FY26`.

The quarter only sets the title of the document ("Changes to Circulars in Q4/2026"). Every row in Summary is included.

## 5. Generating the document

### With the bot

1. Attach the current `bus_data.xlsx` to the chat.
2. Ask for the summary, for example "Erstelle die Summary für dieses Quartal". Name a quarter only if you want a specific one.
3. The bot replies in German with the HTML file and the number of entries per category.

### Locally

```bash
pip install openpyxl
python src/generate_summary.py bus_data.xlsx [out_dir] [quarter]
```

`out_dir` defaults to the current folder. Example with a quarter: `python src/generate_summary.py bus_data.xlsx . Q1/2027`.

## 6. Saving as PDF

1. Open the HTML file in Chrome or Edge.
2. Press Ctrl+P (Windows) or Cmd+P (Mac).
3. Destination: Save as PDF. Paper size: A4. Margins: None.
4. Switch on "Background graphics". Without it the navy background and colors are lost.
5. Save.

## 7. Reading the document

The cover page shows a 2x2 grid for Circulars and another for Standards:
New (top left), Upgrade (top right), Archivings (bottom left), Update
(bottom right). A box without entries says "Nothing to report this
quarter". From page 2 on, each regulation has a table row with
Governance responsibility, number, version, title, link and summary,
English first and German below.

| Color | Meaning |
|---|---|
| Green | New |
| Violet | Upgrade |
| Gold | Update |
| Grey | Archiving |
| Petrol frame | Circulars |
| Blue frame | Standards |

A legend repeats this on every page from page 2.

## 8. Troubleshooting

| Message or symptom | Cause and fix |
|---|---|
| `<number>: Change is ''` | The Change cell is empty. Typical for Standards. Type the Change by hand. |
| `Change is 'Update/Upgrade'` | Decision still open. Enter `Update` or `Upgrade`. |
| `has no column recognizable as the regulation number` | The number column header is not `Document Number` or `Number`. Check the headers. |
| `No quarter could be determined` | Planned Quarter is empty or in an unsupported format. Use `Q4/2026`, or name the quarter explicitly. |
| `No regulations found for <quarter>` | Summary has no rows. |
| Empty cells in the document where Excel shows values | The file was not saved by Excel after the formulas changed. Open, save, and attach it again. |
| White page, no colors in the PDF | "Background graphics" was off while printing. |
| Bot cannot find the script | `generate_summary.py` is missing from Knowledge. Attach it to the chat. |

## 9. Quarter handover

1. Save a copy of the finished workbook for your records.
2. Clear the old rows in Summary (keep the formulas).
3. Update Roadmap for the next quarter.
4. Enter the new Document Numbers and complete Change, versions and summaries.

## 10. For maintainers

- `src/generate_summary.py` contains design, layout and logic. Pagination is calculated in Python from text length. After changing fonts or column widths, check a printed result.
- The bot must load the script with `exec(open("generate_summary.py", encoding="utf-8").read())`, never `import`.
- PDF generation is deliberately not built in. See "Design notes" in `README.md`.
- Bot instructions: `docs/bot_instructions.md`.
