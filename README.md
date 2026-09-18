# Summary of Regulatory Changes Generator

Generates the quarterly "Summary of regulatory changes" directly from
the shared `bus_data.xlsx` spreadsheet. There's no preliminary/published
distinction for this document — it's always the current, complete
state of the target quarter's `Summary` sheet. All data lives in the
spreadsheet; this tool only reads and renders it - it never writes to
the workbook at all. A single self-contained script: HTML-only output,
no PDF library, no browser dependency at runtime.

`Summary` is maintained by hand, not synced automatically from
`Circulars` - an earlier version of this had a `sync()` function and a
companion VBA macro for that; both were dropped (see "Design notes")
because the platform this runs in can't read the `.xlsm` format macros
require, and having only one of the two mechanisms left an
inconsistency not worth the complexity. Populating `Summary` from
`Circulars` is now a manual copy, same as adding an archiving or a new
Standard already was.

Originally built to run as a custom GPT/agent ("SiemensGPT") with Code
Interpreter, but `src/generate_summary.py` is a standalone script and
works the same way run locally.

## Repository structure

```
src/generate_summary.py     - the generator (design + logic in one file)
docs/bot_instructions.md    - instructions for the chat-bot deployment
```

No data lives in this repository — the actual working `bus_data.xlsx`
is maintained separately (SharePoint/Teams) and never committed here.
Generated output (`*SummaryRegChanges_*.html`) isn't committed either,
for the same reason: it's real regulation data, regenerated on demand.

## Usage

```bash
pip install openpyxl

python src/generate_summary.py bus_data.xlsx
```

Optional trailing arguments: `[out_dir]` (default `.`), and `[quarter]`
to override auto-detection (e.g. `Q1/2027`) - see "Which quarter" below.

Or import directly (add `src/` to your path, or run from inside it):

```python
from generate_summary import load_from_xlsx, build

data = load_from_xlsx("bus_data.xlsx")
html_path, _ = build(data, out_dir="/mnt/data")
```

Output is HTML only - PDF generation was deliberately not attempted
(see "Design notes"). Open the HTML in any browser and use its own
Print to PDF if a PDF is needed.

### Spreadsheet schema

Two sheets, both in `bus_data.xlsx`:

**`Circulars`** - maintained by a separate process; this script never
writes to it, not even indirectly. Standard columns: `Circular
Number`, `Governance Responsibility`, `Circular Name`, `Description of
Changes`, `Impact on Existing Circular`, `Current Version`, `Planned
Version`, `Planned Quarter`, `Document URL`.

**`Summary`** - what this document actually reads, maintained by
hand. One row per regulation, EN and DE side by side:

| Column | Notes |
|---|---|
| Type | `Circular` / `Standard` - inferred from the number (`_ST_` -> Standard) if left blank |
| Change | `New` / `Upgrade` / `Update` / `Archiving`. `Update/Upgrade` (undecided) is rejected at render time, row and number named in the error - an undecided item cannot be published |
| Governance responsibility | |
| Number | |
| Current Version / Planned Version | Two separate columns, matching `Circulars`' own shape - copy the values across as they are. Combined into the `2.3 → 3.0` display at render time, not stored pre-combined |
| Title EN / Link EN / Summary EN | |
| Title DE / Link DE / Summary DE | Leave blank if the document is English-only - the German row is then simply omitted from the table |

No `Quarter` column and no `Archived into` column - see "Design notes"
for what that means for archivings specifically (the target circular
has to be named in the typed sentence, not a separate field) and for
row scope in general (`Summary` is expected to hold only the quarter
currently being prepared - clear it out and start the next one fresh
rather than letting rows from different quarters accumulate together).

A blank cell means "not known yet" - rendered as a blank cell too.

## Which quarter

Unless `quarter=` is passed explicitly, `load_from_xlsx()` auto-detects
it: the earliest quarter still present in `Circulars`' `Planned
Quarter` column that is not already behind today's date. A published
quarter drops off `Circulars` once it's finalized elsewhere, so this is
normally the one being prepared next - no `Meta` sheet or manual
bookkeeping required. If a `Meta` sheet with a `Quarter` key is present
and non-blank, that value wins instead. This only decides the label
used in the document title ("...in Q4/2026") - it has nothing to do
with which rows in `Summary` get included, since `Summary` no longer
tracks quarter per row at all (every row present is rendered).

## Deploying as a chat bot

`docs/bot_instructions.md` is written to be pasted into a custom
GPT/agent's instructions field. `src/generate_summary.py` lives in the
bot's Knowledge/persistent files (it rarely changes), `bus_data.xlsx`
is attached fresh each conversation (it changes constantly). Load the
script with `exec()`, never `import` (see "Design notes"). The bot
never edits data itself at all; it only reads whatever spreadsheet is
attached and renders from it. Replies in German, informal register.

## Design notes

A few non-obvious decisions and bugs found while building this,
worth knowing before changing the file:

- **There is no automated way to populate `Summary` from `Circulars` -
  this was tried twice and both attempts were reverted.** First a
  `sync()` function plus a VBA macro that mirrored it for live editing
  inside Excel. The VBA macro needs a macro-enabled workbook
  (`.xlsm`), and the platform this runs in can't read that format at
  all - discovered only after `Summary`'s own columns were
  restructured (`Version` split into `Current Version` / `Planned
  Version`; `Quarter` and `Archived into` dropped) and the *old* VBA
  macro, which hardcoded column positions instead of looking them up
  by header text, silently wrote into the wrong cells against the new
  layout - confirmed by finding pairs of rows per circular in the
  live file, one correctly shaped and one with a URL sitting in the
  `Title EN` column. The macro was rewritten to look up columns by
  header name instead (so a reshuffle fails loudly instead of
  corrupting data silently), but with `.xlsm` unreadable by the
  platform either way, keeping only the Python-side `sync()` (which
  didn't itself need `.xlsm`) would have left two supposedly-
  interchangeable mechanisms that weren't actually equivalent anymore.
  Dropped both rather than maintain that inconsistency. If automated
  sync is reintroduced later, rebuild it in a form the platform can
  actually read end to end - not just locally.
- **PDF generation is intentionally not attempted.** Multiple
  approaches were tried: weasyprint, xhtml2pdf, wkhtmltopdf,
  Playwright, and a network pip-install for one of them. Every one
  either failed or hung the hosting sandbox. HTML-only is the
  answer, permanently - pure string-building and a file write,
  nothing that can hang. If a PDF is genuinely needed, open the HTML
  in any browser and use its own Print to PDF.
- **Pagination is decided entirely in Python before any HTML is
  written - there is no `<script>` tag in the output deciding it
  live.** An earlier version of this file used a browser-side
  JavaScript block that measured the real DOM at render time
  (`window.__paginated`, Playwright driving a live Chromium instance
  for the actual PDF export) - that whole approach depended on a
  runtime browser being available, which doesn't reliably work in the
  hosting sandbox. Replaced with a character-count height-estimation
  approach: `_block_height_mm` / `_card_body_height_mm` predict
  wrapped text height directly from content and column width,
  `_chunk_by_height` packs entries onto pages against that prediction.
  Real Chromium (Playwright) is still used to *calibrate* the
  estimation constants below and to verify real print-to-PDF output
  during development - it is never imported by this file itself.
- **Character-width constants are measured, not guessed** (real
  Chromium, `canvas.measureText`, this table's actual font/size/
  weight) - `_TABLE_CHAR_WIDTH_MM` (7.5pt regular) and
  `_CARD_CHAR_WIDTH_MM` (8pt regular). The calibration sample was
  German text specifically: it measured wider than English at the same
  size in every case checked (umlauts, different letter frequency), so
  using the wider number for both languages is the safe direction -
  predicting slightly more height than needed, never less. Re-measure
  the same way before changing font sizes, weights, or column widths
  in the CSS.
- **The `Governance responsibility` table header needs a 23mm column,
  not 21mm.** Measured directly (`canvas.measureText`): the single
  word "responsibility" alone at bold 7.5pt is 65.0px wide, and 21mm
  of column only gives 64.25px of usable width after padding - 0.75px
  short, just enough to force an orphaned "y" onto a third line by
  itself. Confirmed via an actual rendered page, not predicted.
  Re-measure this specific word before narrowing this column again.
- **A blank `Meta` sheet cell must not become a stored empty string.**
  Found via an actual rendered footer: the copyright year came out
  blank even though the code had `meta.setdefault("Copyright year",
  ...)`. Root cause - if the `Meta` sheet lists a key with nothing
  filled in yet (which it normally does, since the sheet is often
  created with all keys pre-listed), the old loading code inserted
  that key with an empty-string value; `dict.setdefault` and
  `dict.get(key, fallback)` only apply their fallback when the key is
  *absent*, not when it's present-but-falsy, so every default silently
  became a no-op. Fixed at the source: a blank Value cell is now
  treated as if the row weren't there at all, so `setdefault`/`.get(...,
  fallback)` downstream work correctly again. If adding a new
  `meta.setdefault(...)` call, this is already handled - the bug can
  only reappear by reading the `Meta` sheet a second way that skips
  this filter.
- **The Siemens wordmark is inlined as SVG with `fill="currentColor"`,
  not referenced as an `<img src="data:...">`.** An `<img>` reference
  with a CSS `filter` (`invert(...) hue-rotate(...) brightness(...)`)
  to recolor it for light pages was tried first and produced a
  visibly wrong purple-ish tint, confirmed on an actual rendered page.
  Inlining the same path data as a real `<svg fill="currentColor">`
  sidesteps the problem entirely: it inherits `color` from its
  containing page (navy by default, white on `.page.cover`) the same
  way any other text on the page does.
- **Only `.band` claims a cover page's leftover vertical space
  (`flex-grow`), `.foot` has none of its own.** Two elements each
  trying to auto-consume leftover space compete for it and split it
  unevenly instead of one of them cleanly taking all of it. Confirmed
  by an actual rendered page after `margin-top: auto` was mistakenly
  added to both `.band` and `.foot` at once - it left secondary
  elements effectively unconsumed with the footer nowhere near the
  bottom of the page. `.band.circulars` is given a slightly larger
  growth share than the Standards band, since it normally holds more
  cards.
- **The cover page has exactly two states: fits on one page, or
  Standards moves to its own second page - no intermediate
  font-shrinking step.** An earlier browser-JS version of this
  document progressively shrank the whole page's font size in a loop,
  re-measuring after each step, before falling back to a second page.
  That loop depended on live DOM measurement (the same runtime-browser
  dependency this rewrite removes) and added real fragility for a case
  that, in practice, essentially never needs the shrink step at
  realistic data volumes (tested up to 40 regulations across all seven
  buckets - one clean split to two pages, nothing in between was ever
  needed). Simpler and more robust: `_band_height_mm` predicts each
  band's real height from its actual card content, and the page plan
  is decided once, in Python, before any HTML is written.
- **`.page:not(:last-child)` for the page break, not every `.page`.**
  Applying `page-break-after: always` to the last page too creates a
  genuine trailing blank page in Chromium's print pipeline.
- **The footer is in normal document flow, not `position: absolute`.**
  An absolutely positioned element doesn't know when the content above
  it grows taller than expected (e.g. a long German title wrapping to
  three lines) and can end up overlapping it.
- **Table row heights are natural, not forced.** Rows are left at
  their natural content height, never artificially stretched to fill
  the page. The Python-side height estimation here exists only to
  decide *where a page breaks*, not to control each row's rendered
  height.
- **Load the script with `exec()`, never `import`, in the chat-bot
  deployment.** A plain `import` (even with `importlib.reload()`) can
  silently serve Python's cached bytecode from an earlier version
  instead of the freshly attached file. `exec()` never touches the
  import cache at all. Only matters for a persistent Code Interpreter
  session; a normal `python src/generate_summary.py ...` run from the
  command line isn't affected.
- **Verify layout changes against a real rendered PDF (print-to-PDF via
  a real browser), not just reading the HTML or an on-screen check.**
  Every non-obvious bug in this list above was found that way, not by
  inspection.
