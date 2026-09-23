# Corporate Circular Roadmap Generator

Generates the Corporate Circular Roadmap directly from a shared Excel
spreadsheet — a preliminary working version, a final published version
(with an automatic completeness check), or a plain-text status summary.
All data lives in the spreadsheet; this tool only reads and renders it,
it never edits anything itself.

Originally built to run as a custom GPT/agent ("SiemensGPT") with Code
Interpreter, but `src/generate_roadmap.py` is a standalone script and
works the same way run locally.

## Repository structure

```
src/generate_roadmap.py     - the generator (design + logic in one file)
docs/bot_instructions.md    - instructions for the chat-bot deployment
docs/user_guide_de.md       - end-user guide (German)
```

No data lives in this repository — the actual working spreadsheet is
maintained separately (SharePoint/Teams) and never committed here. The
few illustrative examples inside the docs/code use generic placeholder
values (`1_C_001`, `Jane Example`, etc.), not real circular numbers or
names.

## Usage

```bash
pip install openpyxl

python src/generate_roadmap.py roadmap_data.xlsx preliminary ./out
python src/generate_roadmap.py roadmap_data.xlsx final ./out
```

Or import directly:

```python
from generate_roadmap import load_from_xlsx, build

data = load_from_xlsx("roadmap_data.xlsx")
html_path, _ = build(data, final=False, out_dir="./out")
```

Output is HTML only — PDF generation was deliberately removed (see
"Design notes" below). Open the HTML in any browser and use its own
Print to PDF if a PDF is needed.

### Spreadsheet schema

The `Circulars` sheet needs these exact column headers:

| Column | Notes |
|---|---|
| Circular Number | e.g. `1_C_045` |
| Governance Responsibility | Multiple units for one circular go in the same cell, semicolon-separated (e.g. `MAP; ST`) — each renders on its own line |
| Circular Name | |
| Description of Changes | Leave blank if unknown — a generic sentence is generated from Impact at render time |
| Impact on Existing Circular | `Upgrade` / `Update` / `Update/Upgrade` / `New` |
| Current Version | May be blank (e.g. a brand-new circular has no prior version) |
| Planned Version | |
| Contact Circular | |
| Contact Siemens Governance Office | |
| Planned Quarter | e.g. `Q4/2026` |
| Document URL | Optional; if set, Circular Name becomes a hyperlink |
| Internal Comment (never published) | Working notes only |

A blank cell means "not known yet" — that's the only convention for
unknown data; never write placeholder text like "TBD" or "-".

## Deploying as a chat bot

`docs/bot_instructions.md` is written to be pasted into a custom
GPT/agent's instructions field. It expects `roadmap_data.xlsx` attached
fresh each conversation (it changes constantly), and
`generate_roadmap.py` available in the bot's Knowledge/persistent files
(it rarely changes, so it doesn't need re-attaching every time — a
fresh chat attachment still overrides Knowledge for that conversation
if you want to test a change). The bot never edits data itself; it
only reads whatever spreadsheet is attached and renders from it. It's
also instructed to always reply in German (informal register)
regardless of what language the user writes in — see the LANGUAGE
section in `docs/bot_instructions.md` if deploying somewhere that
should behave differently.

See `docs/user_guide_de.md` for the end-user-facing walkthrough (German).

## Design notes

A few non-obvious decisions worth knowing before changing this file:

- **PDF generation is intentionally removed.** Every approach tried
  (multiple PDF libraries, a network `pip install` fallback) either
  failed or hung the hosting sandbox for minutes. HTML-only has zero
  external dependencies and nothing that can hang.
- **`table-layout: fixed` + `overflow-wrap: break-word` on every cell.**
  Narrow columns with long single words (e.g. "Governance
  Responsibility") will overflow their cell without this — confirmed
  by direct measurement in a real browser (Chromium via Playwright),
  not just visual inspection. Removing this reintroduces that bug. Be
  careful going the other way too: a column narrowed *too* far can
  cause a single unbreakable token (e.g. a circular number like
  "1_C_126") to wrap mid-string instead — confirmed via actual PDF
  text extraction, not a screen-rendering check, which did not catch
  it. `COL_NAME_PCT`/`COL_DESC_PCT` and the character-width constants
  near them were sized from measured text width plus a safety margin,
  not guessed.
- **Pages are not one-per-quarter.** Circulars in a quarter are chunked
  across as few pages as they need via `_chunk_by_height`: a hard cap
  of `MAX_CIRCULARS_PER_PAGE` (5) per page, with one exception — a
  quarter whose total is exactly 6 stays on one page if it fits, rather
  than splitting into an awkward 5-plus-1. On top of the count cap,
  `_estimate_row_height_mm` predicts each circular's actual rendered
  row height directly from its content (Circular Name and Description
  of Changes text, wrapped at the real column width using measured
  character widths and line heights — see `_AVG_CHAR_WIDTH_MM`/
  `_LINE_HEIGHT_MM`), and a page splits further if the next circular's
  predicted height would overflow into the footer. This replaced two
  earlier, unmeasured heuristics (a flat per-row "weight" score, then a
  character-count threshold) that both went stale as soon as the column
  widths or row styling changed elsewhere and were never re-validated —
  predicting the real rendered height directly from actual content and
  geometry doesn't need recalibrating just because some other part of
  the layout changes. Validate any change to column widths, font
  sizes/weights, or padding the same way this was built: render real
  and synthetic rows in an actual browser and compare predicted vs.
  measured height, don't just eyeball it.
  **Order is never changed to pack pages more evenly.** A page can end
  up smaller than the count cap even with plenty of later circulars
  that would have fit individually — that's expected when an early
  heavy circular forces the page to close before those later ones are
  ever reached. This was tried once (a first-fit version that pulled a
  later, lighter circular ahead of an earlier, heavier one) and
  reverted — preserving strict ascending order across the whole quarter
  turned out to be a hard requirement, not a style preference.
- **Every page is portrait, always — landscape for light pages was
  tried and explicitly reverted per feedback.** A separate, second
  layer handles making a light page still look intentional:
  `_display_row_heights_mm` runs *after* pagination (never influences
  it) and gives each row on a page a display height slightly taller
  than its own real predicted need, using up whatever space is
  genuinely left over on that specific page. This is deliberately safe
  by construction rather than a fixed floor: it starts from the sum of
  the already-validated real heights for that page, computes exactly
  how much room is left over, and distributes only that (capped per
  row) as a bonus on top of each row's own real height. A flat per-page
  floor (e.g. "every row on this page is at least 38mm") was tried
  first and reverted — it doesn't account for a page where one row
  already needs more than that floor, and every other row getting the
  floor on top can push the page's total height past budget even
  though pagination correctly sized the page. Confirmed as a genuine
  bug via an actual print-to-PDF (6 physical pages instead of the
  correct 5), not just suspected.
  **Set the height via CSS `height`, not `min-height`, on `td`.** These
  are not interchangeable here: `min-height` on a table cell inside
  `table-layout: fixed` was confirmed (via `getComputedStyle` showing
  the correct value, while the actually-rendered row was still far
  shorter) to not reliably control row height at all in this context,
  even though the computed value was correct. Plain `height` has been
  validated repeatedly throughout this project to behave as an
  effective *minimum* here regardless (content still grows past it when
  it needs to) — do not "fix" this back to `min-height`, it looks more
  correct and isn't.
- **The footer is in normal document flow (`margin-top: auto` on
  `.table-wrap`, not on `.footer` itself), not `position: absolute`.**
  An absolutely positioned footer doesn't know when the content above
  it grows taller than expected (e.g. a heavy row), so it can end up
  overlapping the table instead of moving out of the way — confirmed
  via an actual print-to-PDF, not on-screen rendering, which did not
  show the bug. `.page` uses `min-height`, not a fixed `height`, for
  the same reason: it needs to be able to grow. Don't add
  `margin-top: auto` back onto `.footer` itself even though that seems
  intuitive for "push the footer to the bottom" — `.table-wrap` already
  has its own top *and* bottom auto margins for centering, and a third
  auto margin on `.footer` competes with those for the same leftover
  space, producing an uneven split instead of true centering (confirmed
  by direct measurement: a light page's table ended up 84mm from the
  header but only 5mm from the footer, not roughly centered). Removing
  the footer's own margin doesn't move it off the bottom of the page in
  practice — it's still the last flex child on a fixed-height page —
  it just fixes how the space above it gets divided.
- **`.page:not(:last-child)` for the page break, not every `.page`.**
  Applying `page-break-after: always` to the last page too creates a
  genuine trailing blank page in Chromium's print pipeline — confirmed
  by counting physical PDF pages against logical ones.
- **The header (title + PRELIMINARY badge) always stays pinned to the
  top of every page.** An earlier version centered the whole page
  content vertically to avoid a lone gap below a short table, which
  also dragged the title down with it — confirmed wrong by direct
  measurement. The current approach only gives the *table* auto
  top/bottom margins, so it centers within the space below the header
  without ever moving the header itself.
- **When testing layout with Playwright's `page.pdf()`, pass
  `prefer_css_page_size=True`.** Without it, Playwright silently
  defaults to Letter-sized output regardless of this file's
  `@page { size: A4 portrait; }` rule — an entire verification pass in
  this project's history ran against the wrong physical page size
  without any error, only caught by explicitly checking `pdfinfo`'s
  reported page size, not by anything looking wrong in the rendered
  output itself.
- **Sorting is numeric within quarter, chronological across quarters** —
  both computed in `build()`, not left to spreadsheet row order.
- **"Impact" is a display-only merge.** Impact and the Current→Planned
  version transition are shown together in one column, but stay as
  separate columns in the actual spreadsheet.
- **Load the script with `exec()`, never `import`, if re-attaching it
  mid-session.** A plain `import generate_roadmap` (even with
  `importlib.reload()`) can silently serve Python's cached bytecode
  from an earlier version instead of the freshly attached file —
  confirmed by decompiling a `.pyc` the bot had cached and finding none
  of the current fixes in it. `exec()` never touches the import cache
  at all, which removes the failure mode rather than managing it. This
  only matters for the chat-bot deployment path (Code Interpreter
  sessions that persist across messages) — a normal `python
  generate_roadmap.py ...` run from the command line isn't affected.
- **On-screen rendering checks are not sufficient for verifying layout
  fixes.** Several bugs in this project's history (the footer overlap,
  the wrapped circular number) only showed up in an actual print-to-PDF
  or in extracted PDF text, not in on-screen `scrollWidth`/`scrollHeight`
  checks. Verify layout changes against a real rendered PDF, not just a
  screenshot.
- **`generate_roadmap.py` normally lives in the bot's Knowledge/
  persistent files**, not a fresh chat attachment every conversation —
  see `docs/bot_instructions.md`'s Section 6 for the full logic
  (Knowledge as default, a fresh attachment as an explicit override,
  and what to do if Knowledge is ambiguous).
- **Current Version / Planned Version go through
  `_format_version_value`, not a plain `str()`.** Excel stores a
  version typed as e.g. "3" as the actual number 3, and openpyxl reads
  that back as a Python `int` — a plain `str()` conversion then gives
  "3", not "3.0", inconsistent with every other version in the
  document. Confirmed as a real, reproducible bug (not hypothetical),
  by writing an actual int-typed cell and reading it back — the same
  bug was independently found and fixed in a sibling project reading
  version numbers from the same kind of spreadsheet, which is why this
  one was checked rather than assumed not to apply. The fix normalizes
  any numeric cell to at least one decimal place without rounding away
  real extra precision if a version genuinely has it.
- **The department footnote is "P&O TO OD GOV", not "P&O TO OD RG
  GOV".** Changed for the October 1, 2026 reorganization that dropped
  "RG" from this department's own org code — if you're looking at an
  older generated document and it still says "RG GOV", that's not
  stale output, that's genuinely what the department was called then.
