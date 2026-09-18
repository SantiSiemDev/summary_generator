#!/usr/bin/env python3
"""
generate_summary.py
====================
Generates the Summary of Regulatory Changes as HTML from bus_data.xlsx.
Do not reintroduce a PDF library or a runtime browser dependency here;
multiple approaches were tried (weasyprint, xhtml2pdf, wkhtmltopdf,
Playwright, a network pip-install) and every one either failed or hung
the hosting sandbox. HTML-only is pure string-building and a file
write - nothing that can hang. If a PDF is needed, open the HTML in
any browser and use its own Print to PDF.

There is no preliminary/published distinction for this document - it
is always the current, complete picture of the target quarter as the
Summary sheet stands right now. Don't add a draft/final split back in;
it doesn't correspond to anything in the real publishing process this
generates for.

USAGE (from Code Interpreter):
    python generate_summary.py /path/to/bus_data.xlsx

Or import and call directly:
    from generate_summary import build, load_from_xlsx
    data = load_from_xlsx("/path/to/bus_data.xlsx")
    build(data, out_dir="/mnt/data")

INPUT: the "Summary" sheet of bus_data.xlsx if present (one row per
regulation, English + German side by side); falls back to reading
"Circulars" directly if there is no Summary sheet yet (English only,
no archivings - see load_from_xlsx). Only the CURRENT bus quarter is
rendered - "Circulars" may list several quarters at once, but this
document is scoped to what is being published this cycle. "Current" is
auto-detected (the earliest quarter still present in the data that is
not already behind today's date) unless a quarter is passed explicitly
- see load_from_xlsx's `quarter` argument.

LAYOUT: A4 portrait. Page 1 is a dark overview page with up to seven
cards (New/Upgrade/Update/Archivings for Circulars; New/Upgrade/
Archivings for Standards), each listing the regulations in that
bucket. The remaining pages are light detail tables, one section per
(type, bucket) combination that actually has rows: New or revised
Circulars, New or revised Standards, Archived Circulars, Archived
Standards - each starting on its own page and spanning as many pages
as its row count needs. All of this is decided in Python before any
HTML is written (see "Pagination" below) - there is no script tag in
the output deciding it live.

OUTPUT: writes "SummaryRegChanges_<year>_Q<n>.html" to out_dir (year
before quarter, no date stamp; the footer already carries the
generation date).

PDF GENERATION IS INTENTIONALLY NOT ATTEMPTED - see the README for the
history of why. Do not reintroduce it here without being asked.
"""

import sys
import os
import re
import math
from datetime import date

# ---------------------------------------------------------------------
# The real Siemens wordmark, inlined as SVG with fill="currentColor" so it
# automatically follows each page's own text color (white on the dark cover
# page, navy on light detail pages) - no CSS filter/recolor hack needed. An
# earlier version referenced it as an <img src="data:...svg;base64,...">
# with the fill hardcoded to white, which needed a CSS filter to recolor it
# for light pages; that was tried and produced a visibly wrong purple-ish
# tint (confirmed via an actual rendered page, not just reading the filter
# values) - inlining as real SVG with currentColor sidesteps the problem
# entirely rather than fighting the filter further.
# ---------------------------------------------------------------------
WORDMARK = (

    '<svg viewBox="0 0 1000 159" xmlns="http://www.w3.org/2000/svg" class="wordmark" fill="currentColor">'
    '<path d="M3.086 152.537 3.086 122.46C20.205 127.848 35.353 130.542 48.53 130.542 66.723 130.542 75.821 '
    '125.733 75.821 116.122 75.821 112.539 74.497 109.528 71.843 107.09 69.129 104.504 62.178 100.919 51.008 '
    '96.326 30.966 88.085 17.897 81.057 11.818 75.244 3.939 67.571 0 57.895 0 46.202 0 31.144 5.74 19.667 17.212 '
    '11.78 28.557 3.962 43.33 0.057 61.554 0.057 71.595 0.057 86.128 1.905 105.137 5.606L105.137 34.539C90.993 '
    '28.889 77.864 26.07 65.734 26.07 48.653 26.07 40.113 30.76 40.113 40.161 40.113 43.675 41.833 46.541 45.278 '
    '48.763 48.143 50.561 56.037 54.359 68.943 60.169 87.526 68.422 99.897 75.596 106.061 81.698 113.375 88.936 '
    '117.039 98.302 117.039 109.782 117.039 126.283 109.862 138.87 95.518 147.543 83.897 154.576 68.828 158.078 '
    '50.32 158.078 34.69 158.078 18.942 156.237 3.086 152.537L3.086 152.537Z"/>'
    '<path d="M141.063 2.704 183.603 2.704 183.603 155.001 141.063 155.001Z"/>'
    '<path d="M222.616 155.001 222.616 2.704 331.721 2.704 331.721 30.25 263.616 30.25 263.616 64.639 322.898 '
    '64.639 322.898 89.765 263.616 89.765 263.616 125.906 333.476 125.906 333.476 155.001 222.616 155.001Z"/>'
    '<path d="M355.99 155.001 355.99 2.704 411.213 2.704 450.541 100.05 490.955 2.704 543.371 2.704 543.371 '
    '155.001 502.951 155.001 502.951 47.361 459.234 156.316 432.881 156.316 390.05 47.361 390.05 155.001 '
    '355.99 155.001Z"/>'
    '<path d="M577.554 155.001 577.554 2.704 686.658 2.704 686.658 30.25 618.554 30.25 618.554 64.639 677.836 '
    '64.639 677.836 89.765 618.554 89.765 618.554 125.906 688.414 125.906 688.414 155.001 577.554 155.001Z"/>'
    '<path d="M711.927 155.001 711.927 2.704 764.782 2.704 822.749 104.667 822.749 2.704 853.291 2.704 853.291 '
    '155.001 801.973 155.001 743.127 51.746 743.127 155.001 711.927 155.001Z"/>'
    '<path d="M885.316 152.537 885.316 122.46C902.28 127.848 917.402 130.542 930.687 130.542 948.845 130.542 '
    '957.928 125.733 957.928 116.122 957.928 112.539 956.647 109.528 954.088 107.09 951.518 104.504 944.633 '
    '100.919 933.434 96.326 913.319 88.085 900.212 81.057 894.109 75.244 886.245 67.571 882.313 57.895 882.313 '
    '46.202 882.313 31.144 888.017 19.667 899.426 11.78 910.842 3.962 925.635 0.057 943.803 0.057 953.916 '
    '0.057 966.446 1.664 981.394 4.878L984.288 5.606 984.288 34.539C970.208 28.889 957.116 26.07 945.008 26.07 '
    '927.869 26.07 919.3 30.76 919.3 40.161 919.3 43.675 921.026 46.541 924.478 48.763 927.344 50.561 935.281 '
    '54.359 948.288 60.169 966.79 68.422 979.13 75.596 985.294 81.698 992.608 88.936 996.272 98.302 996.272 '
    '109.782 996.272 126.283 989.058 138.87 974.642 147.543 963.021 154.576 947.988 158.078 939.552 158.078 '
    '916.929 158.078 901.166 156.237 885.316 152.537L885.316 152.537Z"/>'
    "</svg>"
)

NAVY = "#000028"
TEAL = "#009999"
GREY = "#DFDFD9"
LINK = "#00557C"
WHITE = "#FFFFFF"

ARROW = "\u2192"
DE_MONTHS = ["Januar", "Februar", "M\u00e4rz", "April", "Mai", "Juni",
             "Juli", "August", "September", "Oktober", "November", "Dezember"]
EN_MONTHS = ["January", "February", "March", "April", "May", "June",
             "July", "August", "September", "October", "November", "December"]

# Office icons lifted from the source PowerPoint deck (Document, Add,
# Caret Up, Repeat, Filing Box), same as before.
ICON_PATHS = {
    "doc": ('<path d="M23 82 23 14 48 14 48 35 73 35 73 82 23 82ZM54 16.5 66.5 29 54 29 54 16.5ZM54 8 17 8 17 88 '
            '79 88 79 30 54 8Z"/><rect x="29" y="45" width="38" height="4"/><rect x="29" y="37" width="13" '
            'height="4"/><rect x="29" y="53" width="38" height="4"/><rect x="29" y="61" width="38" height="4"/>'
            '<rect x="29" y="69" width="38" height="4"/>'),
    "add": '<path d="M88 42 54 42 54 8 42 8 42 42 8 42 8 54 42 54 42 88 54 88 54 54 88 54Z"/>',
    "caret": ('<path d="M23.993 64.122 19.751 59.878 48.007 31.628 76.257 59.878 72.015 64.121 48.007 40.112 '
              '23.993 64.122Z"/>'),
    "repeat": ('<path d="M20.9 57.6 40.1 38.4 29.4 38.4C32.9 31.6 40 26.9 48.2 26.9 53.5 26.9 58.4 28.9 62.1 32.2'
               'L81.6 32.2C75.7 19.6 62.9 11 48.2 11 31.1 11 16.7 22.6 12.5 38.4L1.7 38.4 20.9 57.6Z"/>'
               '<path d="M94.6 57.6 75.4 38.4 56.2 57.6 67 57.6C63.5 64.4 56.4 69.1 48.2 69.1 42.9 69.1 38 67.1 '
               '34.3 63.8L14.7 63.8C20.7 76.4 33.4 85 48.2 85 65.3 85 79.7 73.4 83.9 57.6L94.6 57.6Z"/>'),
    "archive": ('<rect x="10.006" y="18" width="76" height="12"/><path d="M13.006 34 13.006 78 83.006 78 83.006 34Z'
                'M58.006 50 38.006 50 38.006 42 58.006 42Z"/>'),
}

CIRCULAR_CARDS = [("New Siemens Circular", "New", "add"), ("Upgrade", "Upgrade", "caret"),
                   ("Update", "Update", "repeat"), ("Archivings", "Archiving", "archive")]
STANDARD_CARDS = [("New Siemens Standard", "New", "add"), ("Upgrade", "Upgrade", "caret"),
                   ("Archivings", "Archiving", "archive")]

SECTIONS = [
    ("Circular", ("New", "Upgrade", "Update"), "New or revised", "Circulars"),
    ("Standard", ("New", "Upgrade", "Update"), "New or revised", "Standards"),
    ("Circular", ("Archiving",), "Archived", "Circulars"),
    ("Standard", ("Archiving",), "Archived", "Standards"),
]

CSS = """
@page { size: A4 portrait; margin: 0; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body { font-family: 'Siemens Sans', Arial, Helvetica, sans-serif; color: %(NAVY)s; }
.page { position: relative; width: 210mm; height: 297mm; overflow: hidden;
        padding: 13mm 11mm 14mm 11mm; background: %(WHITE)s; display: flex; flex-direction: column; }
/* page-break-after on every .page, including the last one, forces a genuine
   trailing blank page in Chromium's print pipeline - confirmed by counting
   physical PDF pages against logical ones. :not(:last-child) avoids it. */
.page:not(:last-child) { page-break-after: always; }
.page.cover { background: %(NAVY)s; color: %(WHITE)s; }

h1 { font-size: 21pt; font-weight: 700; color: %(TEAL)s; margin: 0; line-height: 1.1; flex-shrink: 0; }
.subtitle { font-size: 8.5pt; font-weight: 700; margin: 1.5mm 0 0 0; flex-shrink: 0; }
.page.cover .subtitle { color: %(WHITE)s; }
.page:not(.cover) .subtitle { color: %(NAVY)s; }
.subtitle u { text-decoration: underline; }

/* ---------- footer: normal document flow, not position:absolute - an
   absolutely positioned footer doesn't know when the content above it
   grows taller than expected and can end up overlapping it (confirmed
   via actual print-to-PDF measurement, not on-screen rendering). No
   margin-top:auto here either - .band below already grows via
   flex-grow to consume the page's leftover space; a second auto-
   growing element here would compete with it for the same space -
   only ONE element should claim the leftover space. Without its own
   margin, .foot still lands at the bottom in practice: it is simply
   the last flex child on a page whose height is fixed. */
.foot { flex-shrink: 0; padding-top: 4mm; display: flex; align-items: flex-end; gap: 4mm; font-size: 6.5pt; }
.foot .pg { font-weight: 700; white-space: nowrap; }
.foot .meta { flex: 1; }
.wordmark { width: 26mm; display: block; flex-shrink: 0; }
.page.cover .foot { color: %(WHITE)s; }

/* ---------- cover ---------- */
/* Only .band claims the page's leftover vertical space (flex-grow) -
   .foot above deliberately has no auto-growth of its own; only ONE
   element on the page should be the sole growing one. The circulars
   band gets a larger share since it normally holds more content
   (4 cards vs. Standards' 3). */
.band { display: flex; gap: 3mm; margin-top: 6mm; flex: 1 1 auto; min-height: 0; }
.band.circulars { flex-grow: 1.4; }
.band-label { writing-mode: vertical-rl; transform: rotate(180deg); font-size: 8pt; font-weight: 700;
              letter-spacing: .4pt; display: flex; align-items: center; justify-content: center; width: 5mm; flex-shrink: 0; }
.grid { flex: 1; display: grid; grid-template-columns: 1fr 1fr; gap: 3mm; align-content: stretch; grid-auto-rows: 1fr; }
.card { display: flex; flex-direction: column; }
.card-h { background: %(TEAL)s; color: %(WHITE)s; font-size: 10pt; font-weight: 700; padding: 2.4mm 3.6mm 2.4mm 3mm;
          display: flex; align-items: center; justify-content: space-between; gap: 2mm;
          border-radius: 1.2mm 1.2mm 0 0; min-height: 10mm; flex: 0 0 auto; }
.card-b { background: %(GREY)s; color: %(NAVY)s; padding: 2.5mm 3mm 3mm 3mm; border-radius: 0 0 1.2mm 1.2mm; }
.card-b ul { margin: 0; padding-left: 3.4mm; }
.card-b li { font-size: 8pt; line-height: 1.25; margin-bottom: 1.2mm; }
.card-b li:last-child { margin-bottom: 0; }
.card-b li.empty { list-style: none; padding-left: 0; }
.cardicon { position: relative; display: inline-block; width: 7mm; height: 7mm; flex: 0 0 auto; }
.ico { position: absolute; inset: 0; width: 100%%; height: 100%%; }
.ico.badge { width: 42%%; height: 42%%; inset: -0.2mm -0.2mm auto auto; }
.ico.solo { width: 78%%; height: 78%%; inset: auto 0 0 auto; }

.note { margin-top: 6mm; border: 0.5pt solid %(WHITE)s; padding: 2.6mm 3mm; font-size: 7.5pt; line-height: 1.35; flex-shrink: 0; }
.note a { color: %(WHITE)s; }

/* ---------- detail tables ---------- */
table { width: 100%%; border-collapse: collapse; table-layout: fixed; margin-top: 4mm; border: 0.8pt solid %(TEAL)s; }
th, td { background: %(GREY)s; border: 0.9pt solid %(WHITE)s; padding: 1.6mm 2mm; font-size: 7.5pt;
         line-height: 1.3; vertical-align: middle; overflow-wrap: break-word; }
th { font-weight: 700; text-align: center; }
td.c-gov, td.c-no, td.c-ver { text-align: center; }
td.c-title { text-align: center; overflow-wrap: anywhere; word-break: break-word; hyphens: auto; }
td.c-sum { text-align: left; }
/* c-gov is 23mm, not a rounder-looking 21 or 22 - measured directly (real
   Chromium, canvas.measureText): the single word "responsibility" alone at
   bold 7.5pt is 65.0px wide, and 21mm only gives 64.25px of usable width
   after padding - 0.75px short, enough to force an orphaned "y" onto its
   own third line. 23mm gives 71.8px, comfortably clear. Re-measure this
   word specifically before narrowing this column again. */
.c-gov { width: 23mm; } .c-no { width: 19mm; } .c-title { width: 34mm; } .c-ver { width: 16mm; }
a { color: %(LINK)s; }
""" % dict(NAVY=NAVY, TEAL=TEAL, GREY=GREY, WHITE=WHITE, LINK=LINK)


# ---------------------------------------------------------------------
# Height estimation constants. Real character widths measured via
# canvas.measureText in actual Chromium (Playwright, dev-time only -
# never imported by this file), German text used for the calibration
# sample since it measured wider than English at the same size in
# every case checked (umlauts, different letter frequency) - using the
# wider number for both languages is the safe direction: predicting
# slightly MORE height than needed, never less. Re-measure the same
# way before changing font sizes, weights, or column widths in the CSS
# above - these numbers are coupled to that CSS, not independent.
# ---------------------------------------------------------------------
_MM_PER_PX = 25.4 / 96
_PT_TO_MM = 25.4 / 72

_TABLE_CHAR_WIDTH_MM = 4.665 * _MM_PER_PX      # 7.5pt regular, German sample
_TABLE_LINE_HEIGHT_MM = 7.5 * 1.3 * _PT_TO_MM  # matches CSS td{font-size:7.5pt;line-height:1.3}
_TABLE_CELL_H_PADDING_MM = 2 * 2                # td padding: 1.6mm 2mm -> 2mm each side
_TABLE_CELL_V_PADDING_MM = 1.6 * 2              # td padding: 1.6mm top + 1.6mm bottom

_CARD_CHAR_WIDTH_MM = 4.971 * _MM_PER_PX        # 8pt regular, German sample
_CARD_LINE_HEIGHT_MM = 8 * 1.25 * _PT_TO_MM     # matches CSS .card-b li{font-size:8pt;line-height:1.25}
_CARD_BULLET_GAP_MM = 1.2                        # .card-b li{margin-bottom:1.2mm}
_CARD_BODY_V_PADDING_MM = 2.5 + 3                # .card-b{padding:2.5mm ... 3mm}
_CARD_HEADER_H_MM = 10                           # .card-h{min-height:10mm}

TABLE_CONTENT_WIDTH_MM = 188      # page width 210mm - 11mm margins each side
COL_GOV_MM, COL_NO_MM, COL_TITLE_MM, COL_VER_MM = 23, 19, 34, 16   # must match the CSS widths above
COL_SUM_MM = TABLE_CONTENT_WIDTH_MM - COL_GOV_MM - COL_NO_MM - COL_TITLE_MM - COL_VER_MM
CARD_GAP_MM = 3
CARD_WIDTH_MM = (TABLE_CONTENT_WIDTH_MM - CARD_GAP_MM) / 2
CARD_BODY_TEXT_WIDTH_MM = CARD_WIDTH_MM - 6 - 3.4   # card padding (3mm each side) + bullet indent

# Vertical budget for the two card bands together on the cover page,
# validated against a real rendered page (see the README) rather than
# derived purely analytically from the surrounding box model - the
# header, subtitle, note box and page padding all contribute and are
# easier to confirm by rendering than to add up by hand.
COVER_BANDS_BUDGET_MM = 205

# Vertical budget for a detail table page's row content, same
# validate-by-rendering approach: 297mm page - 13mm/14mm padding -
# title+subtitle block - table header row - footer block.
TABLE_BODY_BUDGET_MM = 225
MAX_ROWS_PER_TABLE_PAGE = 14


def _wrapped_lines(text, width_mm, char_width_mm):
    text = text or ""
    if not text:
        return 0
    chars_per_line = max(1, width_mm / char_width_mm)
    return max(1, math.ceil(len(text) / chars_per_line))


def _card_body_height_mm(bullets):
    if not bullets:
        return _CARD_BODY_V_PADDING_MM + _CARD_LINE_HEIGHT_MM  # one empty-state bullet
    lines = sum(_wrapped_lines(b, CARD_BODY_TEXT_WIDTH_MM, _CARD_CHAR_WIDTH_MM) for b in bullets)
    return lines * _CARD_LINE_HEIGHT_MM + max(0, len(bullets) - 1) * _CARD_BULLET_GAP_MM + _CARD_BODY_V_PADDING_MM


def _band_height_mm(card_bodies_mm):
    """card_bodies_mm: one pre-computed body height per card, 2 cards per
    visual row (CSS grid-template-columns: 1fr 1fr) - a row's height is
    set by its taller card, same as the browser's own grid layout will
    resolve at render time; this only needs to match that for the
    page-count decision, not reproduce it pixel-for-pixel."""
    rows_h = [_CARD_HEADER_H_MM + max(card_bodies_mm[i:i + 2])
              for i in range(0, len(card_bodies_mm), 2)]
    return sum(rows_h) + max(0, len(rows_h) - 1) * CARD_GAP_MM


def _table_col_height_mm(text, width_mm):
    lines = _wrapped_lines(text, width_mm - _TABLE_CELL_H_PADDING_MM, _TABLE_CHAR_WIDTH_MM)
    return lines * _TABLE_LINE_HEIGHT_MM + _TABLE_CELL_V_PADDING_MM


def _sub_row_height_mm(title, summary):
    return max(_table_col_height_mm(title, COL_TITLE_MM), _table_col_height_mm(summary, COL_SUM_MM))


def _block_height_mm(item):
    """One regulation's full table block: an English row, plus a German
    row underneath if there is German text - each sub-row's height is
    driven by whichever of Title/Summary needs more room at its own
    column width (they wrap independently, side by side)."""
    h = _sub_row_height_mm(item["title_en"], item["summary_en"])
    if item["title_de"] or item["summary_de"]:
        h += _sub_row_height_mm(item["title_de"], item["summary_de"])
    return h


def _chunk_by_height(items, available_height_mm, max_rows):
    """Fills each page in the given order, adding regulations until the
    next one's predicted block height would no longer fit - then starts
    a new page. Never reorders: real predicted content height drives
    the break, a row count cap is a secondary safety net."""
    chunks = []
    current, current_h = [], 0.0
    for it in items:
        h = _block_height_mm(it)
        exceeds_height = current and current_h + h > available_height_mm
        exceeds_count = current and len(current) >= max_rows
        if exceeds_height or exceeds_count:
            chunks.append(current)
            current, current_h = [], 0.0
        current.append(it)
        current_h += h
    if current:
        chunks.append(current)
    return chunks


# ---------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------
def _blank(v):
    return "" if v in (None, "Information pending", "\u2013", "-") else v


def esc(t):
    import html as _html
    return _html.escape(t or "")


def infer_type(number, explicit=""):
    e = (explicit or "").strip().capitalize()
    if e in ("Circular", "Standard"):
        return e
    return "Standard" if "_ST_" in number.upper() else "Circular"


def _parse_date(v):
    for pat, order in ((r"^(\d{2})-(\d{2})-(\d{4})$", "dmy"),
                        (r"^(\d{4})-(\d{2})-(\d{2})(?:[ T]00:00:00)?$", "ymd")):
        m = re.match(pat, (v or "").strip())
        if m:
            a, b, c = (int(x) for x in m.groups())
            return (a, b, c) if order == "dmy" else (c, b, a)
    return None


def parse_quarter(v):
    m = re.match(r"^\s*Q\s*([1-4])\s*[/\- ]\s*(\d{4})\s*$", str(v or ""), re.I)
    return (int(m.group(2)), int(m.group(1))) if m else None


NO_VERSION = {"", "TBD", "TBA", "N/A", "-", "OFFEN", "OPEN"}


def version_display(current, planned):
    cur, pln = (current or "").strip(), (planned or "").strip()
    if pln.upper() in NO_VERSION:
        pln = ""
    if cur.upper() in NO_VERSION:
        cur = ""
    if cur and pln and cur != pln:
        return f"{cur} {ARROW} {pln}"
    return pln or cur


def archiving_text(kind, target, pubdate_str):
    d = _parse_date(pubdate_str)
    en_date = f"{EN_MONTHS[d[1] - 1]} {d[0]}, {d[2]}" if d else pubdate_str
    de_date = f"{d[0]}. {DE_MONTHS[d[1] - 1]} {d[2]}" if d else pubdate_str
    noun_en, noun_de = ("Standard", "Standard") if kind == "Standard" else ("Circular", "Rundschreiben")
    en = f"This Siemens {noun_en} has been archived effective {en_date}"
    de = f"Der Siemens {noun_de}" if kind == "Standard" else f"Das {noun_de}"
    de += f" wurde am {de_date} archiviert"
    if target:
        en += f", and relevant contents have been transferred to Siemens {noun_en} No. {target}."
        de += f" und relevante Inhalte wurden in {'den' if kind=='Standard' else 'das'} Siemens {noun_de} Nr. {target} \u00fcberf\u00fchrt."
    else:
        en += "."
        de += "."
    return en, de


def _finish_item(it, pubdate_str):
    """Fills the boilerplate archiving wording when no summary was typed;
    a real description, however short, is never overwritten."""
    if it["change"] == "Archiving" and not it["summary_en"] and not it["summary_de"]:
        en, de = archiving_text(it["type"], it["archived_into"], pubdate_str)
        it["summary_en"], it["summary_de"] = en, de
        if not it["title_de"]:
            it["title_de"] = it["title_en"]
    return it


# ---------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------
CIRCULARS_SHEET = "Circulars"
SUMMARY_SHEETS = ("Summary", "Changes", "Zusammenfassung")
META_SHEET = "Meta"


def _detect_quarter(circulars_rows):
    """The earliest 'Planned Quarter' still present that is not already
    behind today's date - same rule as before, now computed once here
    at load time instead of inside a separate script layer."""
    found = sorted({q for q in (parse_quarter(c.get("plannedQuarter", "")) for c in circulars_rows) if q})
    if not found:
        return ""
    today = date.today()
    now = (today.year, (today.month - 1) // 3 + 1)
    upcoming = [q for q in found if q >= now]
    y, n = (upcoming or found)[0] if upcoming else found[-1]
    return f"Q{n}/{y}"


def load_from_xlsx(xlsx_path: str, quarter: str = None) -> dict:
    """
    Reads bus_data.xlsx and returns the dict build() expects:
    {"meta": {...}, "items": [...]}. Prefers the "Summary" sheet
    (English + German, archivings, anything hand-added); falls back to
    reading "Circulars" directly if there is no Summary sheet yet
    (English only, no archivings - everything that sheet cannot
    express: unknown stays unknown, nothing is invented).
    `quarter` overrides auto-detection (e.g. "Q4/2026").
    """
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, data_only=True)
    sheets = wb.sheetnames

    meta = {}
    if META_SHEET in sheets:
        ws = wb[META_SHEET]
        for row in ws.iter_rows(min_row=2, values_only=True):
            # A blank Value cell is treated as if the row weren't there at all,
            # not as "the value is the empty string" - otherwise every
            # meta.setdefault(...) below silently becomes a no-op whenever the
            # Meta sheet lists a key with nothing filled in yet, which it
            # normally does (found via a real render: the footer's copyright
            # year came out blank because "Copyright year" was such a row).
            if row and row[0] and row[1] is not None and str(row[1]).strip():
                meta[str(row[0]).strip()] = str(row[1]).strip()

    pubdate = meta.get("Publication date", "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T]00:00:00)?$", pubdate)
    if m:
        pubdate = f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    if not pubdate:
        pubdate = date.today().strftime("%d-%m-%Y")
    meta["Publication date"] = pubdate
    meta.setdefault("Classification", "Restricted")
    meta.setdefault("Governance Office", "Siemens Governance Office (P&O TO OD RG GOV)")
    meta.setdefault("Contact unit", "P&O TO OD RG GOV")
    meta.setdefault("Contact mail", "governance@siemens.com")
    meta.setdefault("Contact unit link", "#")

    road_rows = []
    if CIRCULARS_SHEET in sheets:
        ws = wb[CIRCULARS_SHEET]
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        idx = {h: i for i, h in enumerate(headers) if h}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[idx.get("Circular Number", 0)]:
                continue

            def g(col):
                i = idx.get(col)
                v = row[i] if i is not None and i < len(row) else None
                return "" if v is None else str(v).strip()
            road_rows.append({
                "number": g("Circular Number"), "gov": g("Governance Responsibility"),
                "title": g("Circular Name"), "summary": g("Description of Changes"),
                "change": g("Impact on Existing Circular"),
                "currentVersion": g("Current Version"), "plannedVersion": g("Planned Version"),
                "plannedQuarter": g("Planned Quarter"), "link": g("Document URL"),
            })

    if quarter is None:
        quarter = meta.get("Quarter", "").strip() or _detect_quarter(road_rows)
    meta["Quarter"] = quarter

    sheet_name = next((n for n in SUMMARY_SHEETS if n in sheets), None)
    items = []
    if sheet_name:
        ws = wb[sheet_name]
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        idx = {h: i for i, h in enumerate(headers) if h}

        def col(*names):
            for n in names:
                if n in idx:
                    return idx[n]
            return None

        c_num, c_type, c_chg = col("Number", "Circular Number"), col("Type"), col("Change")
        c_gov = col("Governance responsibility", "Governance Responsibility")
        # Version: prefer the current two-column split (Current Version /
        # Planned Version), matching Circulars' own shape - combined into
        # the "X -> Y" display at read time here, not stored pre-combined.
        # Falls back to a single "Version" column for an older-format
        # sheet that still has one.
        c_curver, c_plnver = col("Current Version"), col("Planned Version")
        c_ver_legacy = col("Version")
        c_ten, c_len, c_sen = col("Title EN", "Circular Name"), col("Link EN", "Document URL"), col("Summary EN", "Description of Changes")
        c_tde, c_lde, c_sde = col("Title DE", "Circular Name (DE)"), col("Link DE", "Document URL (DE)"), col("Summary DE", "Description of Changes (DE)")
        # Archived into: optional. Without this column, an archiving row's
        # auto-generated boilerplate (_finish_item, below) omits the
        # "...transferred to Siemens Circular No. X" clause, since there is
        # nowhere to read the target number from - type the full sentence
        # into Summary EN / Summary DE by hand in that case.
        c_arch, c_q = col("Archived into"), col("Quarter", "Planned Quarter")

        def cell(row, i):
            if i is None or i >= len(row):
                return ""
            v = row[i]
            return "" if v is None else str(v).strip()

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            num = cell(row, c_num)
            title_en = cell(row, c_ten)
            if not num and not title_en:
                continue
            # rq is "" when there's no Quarter/Planned Quarter column at all
            # (the current Summary schema), which makes this filter a no-op
            # below - every row present is treated as belonging to the
            # target quarter, since there's nothing on the row itself to
            # say otherwise. If a Quarter column does exist (an
            # older-format sheet), it still filters as before.
            rq = cell(row, c_q)
            if quarter and rq and parse_quarter(rq) != parse_quarter(quarter):
                continue
            if c_curver is not None or c_plnver is not None:
                ver = version_display(cell(row, c_curver), cell(row, c_plnver))
            else:
                ver = cell(row, c_ver_legacy)
            it = {
                "type": infer_type(num, cell(row, c_type)),
                "change": cell(row, c_chg),
                "gov": cell(row, c_gov), "number": num, "version": ver,
                "title_en": title_en, "link_en": cell(row, c_len), "summary_en": cell(row, c_sen),
                "title_de": cell(row, c_tde), "link_de": cell(row, c_lde), "summary_de": cell(row, c_sde),
                "archived_into": cell(row, c_arch),
            }
            items.append(_finish_item(it, pubdate))
    else:
        # No Summary sheet yet: derive English-only rows straight from
        # Circulars for the target quarter. No archivings (they are never
        # on Circulars) and no German text (there is nowhere to read it from).
        for r in road_rows:
            if quarter and parse_quarter(r["plannedQuarter"]) != parse_quarter(quarter):
                continue
            it = {
                "type": infer_type(r["number"]), "change": r["change"], "gov": r["gov"],
                "number": r["number"],
                "version": version_display(r["currentVersion"], r["plannedVersion"]),
                "title_en": r["title"], "link_en": r["link"], "summary_en": r["summary"],
                "title_de": "", "link_de": r["link"], "summary_de": "", "archived_into": "",
            }
            items.append(_finish_item(it, pubdate))

    for it in items:
        if it["change"] not in ("New", "Upgrade", "Update", "Archiving"):
            raise ValueError(
                f"{it['number']}: Change is '{it['change']}'. Must be one of "
                f"New / Upgrade / Update / Archiving - 'Update/Upgrade' means the decision "
                f"is still open and cannot be published. Fix it in the Summary sheet.")

    return {"meta": meta, "items": items}




# ---------------------------------------------------------------------
# HTML assembly
# ---------------------------------------------------------------------
def _icon_svg(key, cls=""):
    return f'<svg class="ico {cls}" viewBox="0 0 96 96" xmlns="http://www.w3.org/2000/svg" fill="currentColor">{ICON_PATHS[key]}</svg>'


def _card_icon(kind, badge):
    if kind == "Circular":
        return f'<span class="cardicon">{_icon_svg("doc")}{_icon_svg(badge, "badge")}</span>'
    return f'<span class="cardicon">{_icon_svg(badge, "solo")}</span>'


def _bullet_html(it):
    suffix = f" ({esc(it['number'])})" if it["number"] else ""
    return f"<li>{esc(it['title_en'])}{suffix}</li>"


def _card_html(label, kind, change, icon, items):
    bullets = [it for it in items if it["type"] == kind and it["change"] == change]
    if bullets:
        lis = "".join(_bullet_html(it) for it in bullets)
    else:
        lis = "<li class='empty'>&nbsp;</li>"
    return (f'<div class="card"><div class="card-h"><span>{esc(label)}</span>{_card_icon(kind, icon)}</div>'
            f'<div class="card-b"><ul>{lis}</ul></div></div>')


def _title_cell(title, link):
    if not title:
        return ""
    if link:
        return f'<a href="{esc(link)}">{esc(title)}</a>'
    return esc(title)


def _block_rows_html(it):
    langs = [(it["title_en"], it["link_en"], it["summary_en"])]
    if it["title_de"] or it["summary_de"]:
        langs.append((it["title_de"], it["link_de"], it["summary_de"]))
    span = len(langs)
    rows = []
    for n, (title, link, summary) in enumerate(langs):
        meta_cells = ""
        if n == 0:
            meta_cells = (f'<td class="c-gov" rowspan="{span}">{esc(it["gov"])}</td>'
                          f'<td class="c-no" rowspan="{span}">{esc(it["number"])}</td>')
        ver_cell = f'<td class="c-ver" rowspan="{span}">{esc(it["version"])}</td>' if n == 0 else ""
        rows.append(f'<tr>{meta_cells}<td class="c-title">{_title_cell(title, link)}</td>{ver_cell}'
                    f'<td class="c-sum">{esc(summary)}</td></tr>')
    return "".join(rows)


def _table_head_html(kind):
    return ("<thead><tr><th class='c-gov'>Governance responsibility</th>"
            f"<th class='c-no'>{kind} number</th><th class='c-title'>{kind} title</th>"
            "<th class='c-ver'>Version</th><th class='c-sum'>Summary</th></tr></thead>")


def _footer_html(page_num, total_pages, meta):
    return (f'<div class="foot"><span class="pg">Page {page_num} of {total_pages}</span>'
            f'<span class="meta">{esc(meta.get("Classification","Restricted"))} | '
            f'\u00a9 Siemens {esc(meta.get("Copyright year") or meta.get("Publication date","")[-4:])} | '
            f'{esc(meta.get("Governance Office",""))} | {esc(meta.get("Publication date",""))}</span>'
            f'{WORDMARK}</div>')


def _cover_pages_html(items, quarter, meta):
    subtitle = f"New, revised or archived Siemens Circulars and Siemens Standards in {esc(quarter)}"
    circ_bullets = [_card_html(*spec, items) for spec in [(l, "Circular", c, i) for l, c, i in CIRCULAR_CARDS]]
    std_bullets = [_card_html(*spec, items) for spec in [(l, "Standard", c, i) for l, c, i in STANDARD_CARDS]]

    def _bucket_texts(kind, change):
        return [f"{b['title_en']} ({b['number']})" if b["number"] else b["title_en"]
                for b in items if b["type"] == kind and b["change"] == change]

    circ_bodies = [_card_body_height_mm(_bucket_texts("Circular", chg)) for _, chg, _ in CIRCULAR_CARDS]
    std_bodies = [_card_body_height_mm(_bucket_texts("Standard", chg)) for _, chg, _ in STANDARD_CARDS]
    circ_h = _band_height_mm(circ_bodies)
    std_h = _band_height_mm(std_bodies)

    note = (f'For more details regarding the changes made, please have a look at the following pages.<br>'
            f'If you are interested in previous summaries of regulatory changes, please feel free to contact us '
            f'(<a href="{esc(meta.get("Contact unit link","#"))}">{esc(meta.get("Contact unit",""))}</a>, '
            f'<a href="mailto:{esc(meta.get("Contact mail",""))}">{esc(meta.get("Contact mail",""))}</a>).')

    pages = []
    if circ_h + std_h <= COVER_BANDS_BUDGET_MM:
        body = (f'<h1>Summary of regulatory changes</h1><div class="subtitle">{subtitle}</div>'
                f'<div class="band circulars"><div class="band-label">CIRCULARS</div><div class="grid">{"".join(circ_bullets)}</div></div>'
                f'<div class="band"><div class="band-label">STANDARDS</div><div class="grid">{"".join(std_bullets)}</div></div>'
                f'<div class="note">{note}</div>')
        pages.append(f'<div class="page cover">{body}</div>')
    else:
        body1 = (f'<h1>Summary of regulatory changes</h1><div class="subtitle">{subtitle}</div>'
                 f'<div class="band circulars"><div class="band-label">CIRCULARS</div><div class="grid">{"".join(circ_bullets)}</div></div>')
        pages.append(f'<div class="page cover">{body1}</div>')
        body2 = (f'<h1>Summary of regulatory changes</h1><div class="subtitle">{subtitle}</div>'
                 f'<div class="band"><div class="band-label">STANDARDS</div><div class="grid">{"".join(std_bullets)}</div></div>'
                 f'<div class="note">{note}</div>')
        pages.append(f'<div class="page cover">{body2}</div>')
    return pages


def _detail_pages_html(items, quarter):
    pages = []
    for kind, changes, prefix, noun in SECTIONS:
        entries = [it for it in items if it["type"] == kind and it["change"] in changes]
        if not entries:
            continue
        subtitle = f'<u>{esc(prefix)} </u>Siemens <u>{esc(noun)}</u> in {esc(quarter)}'
        chunks = _chunk_by_height(entries, TABLE_BODY_BUDGET_MM, MAX_ROWS_PER_TABLE_PAGE)
        head = _table_head_html(kind)
        for chunk in chunks:
            rows = "".join(_block_rows_html(it) for it in chunk)
            body = (f'<h1>Summary of regulatory changes</h1><div class="subtitle">{subtitle}</div>'
                    f'<table>{head}<tbody>{rows}</tbody></table>')
            pages.append(f'<div class="page">{body}</div>')
    return pages


def build(data: dict, out_dir: str = "."):
    """
    data: the parsed dataset dict (must have "meta" and "items") -
          normally produced by load_from_xlsx(), not built by hand.
    out_dir: directory to write the output HTML into.
    Returns (html_path, None) - the second value is always None; PDF
    generation is intentionally not attempted (see module docstring).
    There is no preliminary/final distinction here - this always
    renders the current, complete state of the target quarter's
    Summary sheet.
    """
    meta = data.get("meta", {})
    items = data.get("items", [])
    quarter = meta.get("Quarter", "")
    if not quarter:
        raise ValueError("No quarter could be determined - set Meta!Quarter, pass quarter= "
                         "explicitly, or check 'Planned Quarter' values in Circulars.")
    if not items:
        raise ValueError(f"No regulations found for {quarter}.")

    raw_pages = _cover_pages_html(items, quarter, meta) + _detail_pages_html(items, quarter)
    total = len(raw_pages)
    # The footer needs the complete page count ("Page N of M"), so it's
    # spliced in only after the full page list is known - each page's
    # HTML ends in "...</div></div>" (content div, then .page div); the
    # footer goes just before that LAST closing </div>.
    pages = []
    for i, p_html in enumerate(raw_pages, start=1):
        stripped = p_html.rstrip()
        assert stripped.endswith("</div>")
        pages.append(stripped[:-len("</div>")] + _footer_html(i, total, meta) + "</div>")

    html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8" /><title>Summary of regulatory changes</title>
<style>{CSS}</style></head><body>{''.join(pages)}</body></html>"""

    m = re.match(r"Q(\d)/(\d{4})", quarter)
    quarter_slug = f"{m.group(2)}_Q{m.group(1)}" if m else quarter.replace("/", "_")
    base_name = f"SummaryRegChanges_{quarter_slug}"

    os.makedirs(out_dir, exist_ok=True)
    html_path = os.path.join(out_dir, base_name + ".html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    # PDF generation is intentionally not attempted - see module docstring.
    # pdf_path is always None; kept in the return value so calling code
    # that unpacks (html_path, pdf_path) doesn't break.
    return html_path, None


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python generate_summary.py bus_data.xlsx [out_dir] [quarter]")
        sys.exit(1)
    data_path = sys.argv[1]
    if not data_path.lower().endswith(".xlsx"):
        print("Input must be an .xlsx spreadsheet.")
        sys.exit(1)

    out_dir_arg = sys.argv[2] if len(sys.argv) > 2 else "."
    quarter_arg = sys.argv[3] if len(sys.argv) > 3 else None
    summary_data = load_from_xlsx(data_path, quarter=quarter_arg)
    html_out, pdf_out = build(summary_data, out_dir=out_dir_arg)
    print("HTML written to:", html_out)
