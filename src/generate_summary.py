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
regulation, English + German side by side); falls back to reading the
roadmap sheet directly if there is no Summary sheet yet (English only,
no archivings - see load_from_xlsx; accepts either "Roadmap" or
"Circulars" as that sheet's name - see CIRCULARS_SHEETS). Only the
CURRENT bus quarter is rendered - the roadmap sheet may list several
quarters at once, but this document is scoped to what is being
published this cycle. "Current" is
auto-detected (the earliest quarter still present in the data that is
not already behind today's date) unless a quarter is passed explicitly
- see load_from_xlsx's `quarter` argument.

LAYOUT: A4 portrait, every page on the same dark navy background. Page 1
is an overview with up to seven cards (New/Upgrade/Update/Archivings
for Circulars; New/Upgrade/Archivings for Standards), each colored by
its Change state and listing the regulations in that bucket. The
remaining pages are detail tables (light rows on the same dark page),
one section per (type, bucket) combination that actually has rows: New
or revised Circulars, New or revised Standards, Archived Circulars,
Archived Standards - each starting on its own page and spanning as
many pages as its row count needs, bordered in teal for Circulars or
blue for Standards, with a colored left-edge stripe per row marking
its Change state. All of this is decided in Python before any HTML is
written (see "Pagination" below) - there is no script tag in
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
# automatically follows the page's own text color - no CSS filter/recolor
# hack needed. An earlier version referenced it as an
# <img src="data:...svg;base64,...")> with the fill hardcoded to white,
# which needed a CSS filter to recolor it for a light-page variant of this
# document; that was tried and produced a visibly wrong purple-ish tint
# (confirmed via an actual rendered page, not just reading the filter
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
GREY = "#DFDFD9"
LINK = "#00557C"
WHITE = "#FFFFFF"

# Two "family" colors distinguishing Circulars from Standards - applied
# consistently everywhere the family matters: the cover page's vertical
# band label, the underlined word in each detail page's subtitle, and
# the detail table's header row / outer border. Both are existing
# Siemens brand colors, chosen for a clear enough hue difference that
# they don't read as "the same blue" side by side on the cover page.
CIRC = "#009999"   # Circulars - the original brand teal
STD = "#0087BE"    # Standards - a distinctly bluer brand accent
TEAL = CIRC        # kept as an alias: used for links, the h1 title, etc.
                    # (page-level branding, not the family-color system)

# Four "state" colors for Change = New / Upgrade / Update / Archiving -
# applied to each cover-page card header and, on the detail tables, as a
# left-edge stripe on every row block, so the same color means the same
# thing whether you're looking at the overview or the detail. All four
# were darkened from their brighter starting points specifically so
# white text/icons sit on them at a real, checked contrast ratio (WCAG,
# computed - not eyeballed): each is >=4:1 against white. Upgrade is
# violet rather than a third blue on purpose, so it can never be
# mistaken for the Standards family color STD when a Standards card
# happens to be an Upgrade.
NEW = "#008B68"
UPGRADE = "#6A5AA8"
UPDATE = "#9D6D00"
ARCHIVED = "#6B7280"
STATE_COLOR = {"New": NEW, "Upgrade": UPGRADE, "Update": UPDATE, "Archiving": ARCHIVED}
FAMILY_COLOR = {"Circular": CIRC, "Standard": STD}

ARROW = "\u2192"
DE_MONTHS = ["Januar", "Februar", "M\u00e4rz", "April", "Mai", "Juni",
             "Juli", "August", "September", "Oktober", "November", "Dezember"]
EN_MONTHS = ["January", "February", "March", "April", "May", "June",
             "July", "August", "September", "October", "November", "December"]

# Simple stroke-based icons (viewBox 0 0 24 24, round caps/joins) - one
# per Change state, colored via the card header background they sit on
# (STATE_COLOR) rather than needing their own fill color. Replaces an
# earlier compound "document + small badge" icon lifted from the source
# PowerPoint deck: cleaner at small size, and the card header text
# already says "New Siemens Circular" / "Upgrade" / etc., so a second
# "this is a document" icon in the corner was redundant, not just
# unpolished.
ICON_PATHS = {
    "add": '<path d="M12 5v14M5 12h14"/>',
    "caret": '<path d="M12 19V5M5 12l7-7 7 7"/>',
    "repeat": ('<path d="M3 12a9 9 0 0 1 15-6.7L21 8"/><path d="M21 3v5h-5"/>'
               '<path d="M21 12a9 9 0 0 1-15 6.7L3 16"/><path d="M3 21v-5h5"/>'),
    "archive": ('<rect x="3" y="4" width="18" height="4" rx="1.3"/>'
                '<path d="M5 8v10a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8"/><path d="M10 13h4"/>'),
}

CIRCULAR_CARDS = [("New Siemens Circular", "New", "add"), ("Upgrade", "Upgrade", "caret"),
                   ("Archivings", "Archiving", "archive"), ("Update", "Update", "repeat")]
STANDARD_CARDS = [("New Siemens Standard", "New", "add"), ("Upgrade", "Upgrade", "caret"),
                   ("Archivings", "Archiving", "archive")]

# Detail pages are one continuous section per family now, not split by
# New/Revised vs Archived - the per-row color stripe already carries
# that distinction, so a separate "chapter" for it was redundant. Order
# entries are sorted into within a section, matching the cover cards'
# own bucket order.
SECTION_STATE_ORDER = ["New", "Upgrade", "Update", "Archiving"]
SECTIONS = [
    ("Circular", "Circulars"),
    ("Standard", "Standards"),
]

CSS = """
@page { size: A4 portrait; margin: 0; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body { font-family: 'Siemens Sans', Arial, Helvetica, sans-serif; color: %(WHITE)s; }
/* Every page shares one background now - previously only the cover page
   was navy and the detail pages were white/light. table rows stay light
   (see below), so the effect is light "cards" floating on a dark page,
   consistent front to back rather than switching palette mid-document. */
.page { position: relative; width: 210mm; height: 297mm; overflow: hidden;
        padding: 13mm 11mm 14mm 11mm; background: %(NAVY)s; display: flex; flex-direction: column; }
/* page-break-after on every .page, including the last one, forces a genuine
   trailing blank page in Chromium's print pipeline - confirmed by counting
   physical PDF pages against logical ones. :not(:last-child) avoids it. */
.page:not(:last-child) { page-break-after: always; }

h1 { font-size: 21pt; font-weight: 700; color: %(TEAL)s; margin: 0; line-height: 1.1; flex-shrink: 0; }
.subtitle { font-size: 8.5pt; font-weight: 700; margin: 1.5mm 0 0 0; flex-shrink: 0; color: %(WHITE)s; }

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
.foot { flex-shrink: 0; padding-top: 4mm; display: flex; align-items: flex-end; gap: 4mm; font-size: 6.5pt; color: %(WHITE)s; }
.foot .pg { font-weight: 700; white-space: nowrap; }
.foot .meta { flex: 1; }
.wordmark { width: 26mm; display: block; flex-shrink: 0; }

/* ---------- color legend: a single row of the four state colors, on
   every page EXCEPT the cover (see LEGEND_HTML/_footer_html) - the
   cover page already establishes what each color means through the
   cards themselves, so repeating it there was redundant rather than
   reassuring. No family (Circulars/Standards) swatches here either -
   each detail page's own colored table header/border already answers
   that, on that same page. Sits above the normal page-number/
   classification line, inside the same flex-shrink:0 footer block -
   counted into TABLE_BODY_BUDGET_MM like everything else in .foot. */
.legend { flex-shrink: 0; padding-top: 3mm; border-top: 0.6pt solid rgba(255,255,255,0.25); }
.legend-row { display: flex; align-items: center; gap: 4mm; font-size: 6.5pt; color: %(WHITE)s; }
.legend-item { display: flex; align-items: center; gap: 1.3mm; }
.legend-dot { width: 2.6mm; height: 2.6mm; border-radius: 50%%; flex-shrink: 0; display: inline-block; }

/* ---------- cover ---------- */
/* Both bands now share the page's available vertical space equally
   (flex: 1, same ratio for both) rather than sizing to their own
   natural content - the whole point of this layout is that every box
   on the page, Circulars or Standards, ends up the same size
   regardless of how much content it holds, which only happens if
   both bands claim an equal share of the same total space. An
   earlier version gave bands their natural content height instead
   (with .note absorbing whatever was left over) specifically to avoid
   stretching a sparse band - that was the right call when box count
   and content varied page to page and stretching made the
   inconsistency more visible, not less. It's the opposite call here
   on purpose: a fixed 2x2 layout (see CIRCULAR_CARDS/STANDARD_CARDS)
   makes uniform, page-filling boxes the actual goal, not a bug to
   avoid.

   Flex's own default min-height-from-content is what's supposed to
   protect against a genuinely tall block of text being squeezed
   smaller than it needs, growing a band beyond its equal share rather
   than letting content overflow it - but only if nothing in the chain
   overrides that default. An earlier version of this rule added
   min-height:0 here (and on .band-content, .card) specifically to
   force the equal split - which also silently disabled that exact
   protection. Confirmed as a real bug, not a theoretical one: with
   that override in place, a card's grid was measured 33px taller than
   its own containing band, invisibly overflowing past .page's own
   overflow:hidden. Removed everywhere in this chain - _band_height_mm
   / COVER_BANDS_BUDGET_MM in the Python code still separately decides
   the 1-page-vs-2-page split from content's natural minimum, so the
   two mechanisms now agree instead of fighting each other. */
.band { display: flex; gap: 4mm; margin-top: 12mm; flex: 1; }
/* Not :first-of-type - a div.subtitle sits before both bands, so
   neither .band is ever literally the first div among its siblings
   and the rule never matched at all (confirmed by inspecting computed
   margin-top on both bands: both reported 12mm). Targeting the
   .circulars / .standards classes _band_html already gives each band
   is unambiguous instead of relying on sibling position. */
.band.circulars { margin-top: 9mm; }
/* Vertical orientation, same as the original design - a full-width
   horizontal bar was tried instead and reverted per feedback (too
   heavy-handed a departure). Kept vertical, but bigger and paired with
   a solid color bar of its own (not just colored text) - that combo is
   what actually makes the boundary between the two bands unmistakable
   at a glance, which plain rotated text on its own wasn't. */
.band-sidebar { display: flex; align-items: stretch; gap: 3mm; flex-shrink: 0; }
.band-bar { width: 4mm; border-radius: 2mm; flex-shrink: 0; }
.band-bar.circ { background: %(CIRC)s; }
.band-bar.std { background: %(STD)s; }
.band-label { writing-mode: vertical-rl; transform: rotate(180deg); font-size: 15pt; font-weight: 700;
              letter-spacing: 1.4pt; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
.band-label.circ { color: %(CIRC)s; }
.band-label.std { color: %(STD)s; }
.band-content { flex: 1; display: flex; flex-direction: column; min-width: 0; }
/* grid-template-rows: 1fr 1fr (not auto) is what makes both rows in a
   band's 2x2 layout the same height as each other - matching
   .band's own equal-share flex-grow above, so a box's size depends on
   neither its own content nor which band it's in, only on how much
   total space the page has to give. grid:1 makes the grid itself fill
   .band-content's full height rather than just its own natural
   content height, the same reasoning as .band above. Standards has
   only 3 cards (see STANDARD_CARDS) placed into a grid still sized
   for 4 - CSS grid's own auto-placement already leaves the 4th cell
   genuinely blank without any extra markup, at the same size the
   Update card would have taken up. */
.grid { flex: 1; display: grid; grid-template-columns: 1fr 1fr; grid-template-rows: 1fr 1fr; gap: 3mm; }
.card { display: flex; flex-direction: column; }
/* Rounded on all four corners (not just the top) now that the header and
   body are visually one continuous card rather than a flat-topped tab -
   part of the general move to rounder, softer edges throughout. Header
   background is set inline per card (one of the four STATE_COLOR values)
   rather than fixed here, since New/Upgrade/Update/Archiving each need
   their own color - see _card_html(). */
.card-h { color: %(WHITE)s; font-size: 10pt; font-weight: 700; padding: 2.4mm 3.6mm 2.4mm 3mm;
          display: flex; align-items: center; justify-content: space-between; gap: 2mm;
          border-radius: 2.6mm 2.6mm 0 0; min-height: 10mm; flex: 0 0 auto; }
/* flex:1 so the body fills whatever height the card ends up with (see
   .card/.grid above) rather than staying at its own natural content
   height and leaving a gap of bare page background below it - the
   grey fill needs to reach the card's actual bottom edge now that a
   card's size isn't just its own content's doing. */
.card-b { flex: 1; background: %(GREY)s; color: %(NAVY)s; padding: 2.5mm 3mm 3mm 3mm; border-radius: 0 0 2.6mm 2.6mm; }
.card-b ul { margin: 0; padding-left: 3.4mm; }
.card-b li { font-size: 8pt; line-height: 1.25; margin-bottom: 1.2mm; }
.card-b li:last-child { margin-bottom: 0; }
.cardicon { display: flex; align-items: center; justify-content: center; width: 6.4mm; height: 6.4mm; flex: 0 0 auto; }
.cardicon svg { width: 100%%; height: 100%%; }
/* A fixed set of cards is always shown now (4 for Circulars, 3 for
   Standards), even when a bucket is empty - shown once, then dropped
   per feedback that a variable number of boxes looked inconsistent
   from quarter to quarter. An empty bucket gets this instead of an
   empty bullet list: centered (both axes, via .card-b.empty - the
   card still stretches to match its row-mate's height via the grid's
   own default align-items:stretch, so this can end up taller than one
   line and should stay centered rather than pinned to the top), muted
   and italic so it reads as a deliberate "nothing here" rather than
   missing content. */
.card-b.empty { display: flex; align-items: center; justify-content: center; }
.card-empty { margin: 0; text-align: center; font-style: italic; font-size: 8pt;
              color: rgba(0, 0, 40, 0.45); }

/* margin-top is a fixed value now, not auto - the two bands claim the
   page's leftover space themselves (flex:1 each, see .band above), so
   there's nothing left over for .note to absorb the way it used to
   when bands sized to their own content instead. */
.note { margin-top: 6mm; border: 0.5pt solid %(WHITE)s; border-radius: 2.2mm; padding: 2.6mm 3mm;
        font-size: 7.5pt; line-height: 1.35; flex-shrink: 0; }
.note a { color: %(WHITE)s; }

/* ---------- detail tables ---------- */
/* Rounded outer corners come from this wrapper (border + border-radius +
   overflow:hidden clipping the table's own square corners), not from the
   table element itself - a <table> can't reliably take a border-radius
   directly (border-collapse fights it), but a div wrapping it can, and
   clips whatever's inside to match. Border color is the same family
   system as the band label / subtitle underline - set inline per table
   (CIRC or STD) by build(), not fixed here. */
/* No border-left here (see below) - the left edge belongs entirely to
   each row's own state-color stripe, not to this wrapper's family
   color. The two used to sit immediately next to each other as two
   thin, differently-colored lines, which read as a cluttered "mixed"
   edge rather than the one clear color it was meant to be - confirmed
   by zooming into an actual rendered page, not assumed. */
.table-wrap { margin-top: 4mm; border-radius: 3mm; overflow: hidden; border: 1pt solid; border-left: none; }
table { width: 100%%; border-collapse: collapse; table-layout: fixed; }
th, td { background: %(GREY)s; border: 0.9pt solid %(WHITE)s; padding: 1.6mm 2mm; font-size: 7.5pt;
         line-height: 1.3; vertical-align: middle; overflow-wrap: break-word; }
/* Header row background is the family color (not grey) with white text -
   an immediate "this whole table is Circulars / Standards" signal that
   doesn't depend on reading the subtitle above it. Set inline per table
   alongside the wrapper border, for the same reason. */
th { font-weight: 700; text-align: center; color: %(WHITE)s; }
/* Body cells need their own dark text - white body text (the page-wide
   default, correct for the navy page itself) on this light grey is
   ~1.3:1 contrast, nowhere near legible; confirmed by computing the
   actual ratio, not by judging it on screen. Navy on this same grey is
   the pattern .card-b already uses successfully (~15:1) - applying it
   here too rather than introducing a different fix for what is the
   same underlying problem (dark text needed on a light background). */
td { color: %(NAVY)s; }
td.c-gov, td.c-no, td.c-ver { text-align: center; }
td.c-title { text-align: center; overflow-wrap: anywhere; word-break: break-word; hyphens: auto; }
td.c-sum { text-align: left; }
/* The Change-state stripe on c-gov (New/Upgrade/Update/Archiving, one of
   the four STATE_COLOR values, set inline per row block by
   _block_rows_html()) carries the cover page's color code into the
   detail tables - the same visual language both places, not just a
   color scheme that happens to reappear. c-gov carries it rather than a
   dedicated column because it already spans the full row block via
   rowspan, so one border runs the block's whole height for free. */
/* Widened from 3.5pt - the one and only color on this edge now that
   the wrapper no longer has its own competing left border (see
   .table-wrap), so it can afford - and benefits from - being more
   prominent. */
/* th.c-gov included here too (not just td.c-gov) - matching width is
   what makes the header's own left-edge line (see _table_head_html)
   read as a continuation of the body stripe below it, rather than an
   abrupt change from a thin 0.9pt white border to a thick 6pt colored
   one right at the header/body boundary. */
th.c-gov, td.c-gov { border-left-width: 6pt; }
/* The header row gets the same left-edge treatment, in its own color
   (the family color it already has as a background, set inline
   alongside it in _table_head_html so the border exactly matches the
   fill rather than approximating it) - without this, the state-color
   stripe just starts abruptly at the first data row with nothing
   above it, which reads as the bar stopping short rather than a
   deliberate design element. Same width as the data-row stripes below
   it, so the two read as one continuous bar that happens to change
   color at each section boundary, not as two unrelated details. */
th.c-gov { border-left-width: 6pt; }
/* c-gov is 23mm, not a rounder-looking 21 or 22 - measured directly (real
   Chromium, canvas.measureText): the single word "responsibility" alone at
   bold 7.5pt is 65.0px wide, and 21mm only gives 64.25px of usable width
   after padding - 0.75px short, enough to force an orphaned "y" onto its
   own third line. 23mm gives 71.8px, comfortably clear. Re-measure this
   word specifically before narrowing this column again. */
.c-gov { width: 23mm; } .c-no { width: 19mm; } .c-title { width: 34mm; } .c-ver { width: 16mm; }
a { color: %(LINK)s; }
""" % dict(NAVY=NAVY, TEAL=TEAL, GREY=GREY, WHITE=WHITE, LINK=LINK, CIRC=CIRC, STD=STD)


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
# .band-header no longer exists as a horizontal bar (reverted to a
# vertical sidebar that stretches alongside the content instead - see
# .band-sidebar), so it no longer adds its own height on top of the
# band's content - removed from the height sum below accordingly.
# .band-empty placeholder box shown when a whole band has nothing in it
_BAND_EMPTY_H_MM = 20

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
COVER_BANDS_BUDGET_MM = 200

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
    """An empty bucket (no bullets) still needs a real height estimate
    now that it always renders as a card - one centered line of muted
    "Nothing to report this quarter" text (see .card-empty), not zero."""
    if not bullets:
        return _CARD_HEADER_H_MM + _CARD_LINE_HEIGHT_MM + _CARD_BODY_V_PADDING_MM
    lines = sum(_wrapped_lines(b, CARD_BODY_TEXT_WIDTH_MM, _CARD_CHAR_WIDTH_MM) for b in bullets)
    return (_CARD_HEADER_H_MM + lines * _CARD_LINE_HEIGHT_MM
            + max(0, len(bullets) - 1) * _CARD_BULLET_GAP_MM + _CARD_BODY_V_PADDING_MM)


def _band_height_mm(card_bodies_mm):
    """card_bodies_mm: one pre-computed total height (header + body) per
    card, always the full fixed set (4 for Circulars, 3 for Standards -
    an empty bucket still contributes its own "nothing to report"
    height, it's never left out). Cards pair up 2-per-row (CSS
    grid-template-columns: 1fr 1fr), except a trailing lone card
    (Standards' 3rd), which spans both columns (.card.wide) and so
    counts as its own full-width row. A row's height is set by its
    taller card, same as the browser's own grid layout will resolve at
    render time; this only needs to match that for the page-count
    decision, not reproduce it pixel-for-pixel. The vertical sidebar
    label stretches to match whatever height this comes out to (see
    .band-sidebar), so it doesn't add anything of its own."""
    rows_h = []
    i = 0
    while i < len(card_bodies_mm):
        if i == len(card_bodies_mm) - 1:
            rows_h.append(card_bodies_mm[i])  # lone trailing card, full width, its own row
            i += 1
        else:
            rows_h.append(max(card_bodies_mm[i], card_bodies_mm[i + 1]))
            i += 2
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


def _ensure_decimal(v):
    """Version numbers like '3' (no decimal point - typically because the
    source cell held a plain Excel number, not text, so a whole number
    lost its trailing '.0' when the cell converted to text) are shown as
    '3.0', matching the one-decimal-or-more format every other version
    uses. Anything that already has a decimal point, or isn't a plain
    integer to begin with (blank, "TBD", etc.), is left exactly as-is."""
    return v + ".0" if v and re.match(r"^\d+$", v) else v


def version_display(current, planned, change=None):
    """New and Archiving don't have a version *transition* - only Update
    and Upgrade do. Forcing that here (rather than just trusting Current
    Version / Planned Version to be filled in "correctly" per Change)
    means a stray value in the "wrong" cell can't accidentally produce a
    nonsensical arrow - e.g. someone fills in a Current Version for a
    brand-new regulation, or the same version in both cells for an
    archiving: New always shows just the one version (whichever is
    filled - normally Planned, since New has nothing to have a "current"
    version of); Archiving always shows just the one version being
    archived (whichever is filled - normally Current). Neither ever
    shows an arrow, regardless of what's in the other cell."""
    cur, pln = (current or "").strip(), (planned or "").strip()
    if pln.upper() in NO_VERSION:
        pln = ""
    if cur.upper() in NO_VERSION:
        cur = ""
    cur, pln = _ensure_decimal(cur), _ensure_decimal(pln)
    if change == "New":
        return pln or cur
    if change == "Archiving":
        return cur or pln
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
CIRCULARS_SHEETS = ("Roadmap", "Circulars")
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


def load_from_xlsx(xlsx_path: str, quarter: str | None = None) -> dict:
    """
    Reads bus_data.xlsx and returns the dict build() expects:
    {"meta": {...}, "items": [...]}. Prefers the "Summary" sheet
    (English + German, archivings, anything hand-added); falls back to
    reading the roadmap sheet directly if there is no Summary sheet yet
    (accepts "Roadmap" or "Circulars" as that sheet's name - see
    CIRCULARS_SHEETS) (English only, no archivings - everything that
    sheet cannot express: unknown stays unknown, nothing is invented).
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
    # "RG" dropped from the org code (was "P&O TO OD RG GOV") - the
    # reorganization effective October 1 retires that segment, and this
    # document is always published after that date, so the old code
    # would already be wrong on every document it's ever used for.
    meta.setdefault("Governance Office", "Siemens Governance Office (P&O TO OD GOV)")
    meta.setdefault("Contact unit", "P&O TO OD GOV")
    meta.setdefault("Contact mail", "governance@siemens.com")
    meta.setdefault("Contact unit link", "#")

    road_rows = []
    circulars_sheet = next((n for n in CIRCULARS_SHEETS if n in sheets), None)
    if circulars_sheet:
        ws = wb[circulars_sheet]
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
        # Case-insensitive, whitespace-trimmed matching - consistent with
        # every other header lookup in this file. This one used to be a
        # raw exact-string match, which is exactly the kind of thing a
        # header like "Number " (trailing space) or "number" (different
        # case) silently defeats: col() would return None, cell(row, None)
        # returns "" for every row, and the whole document renders with
        # blank Circular/Standard numbers throughout - every OTHER field
        # still works, since their own header happened to still match,
        # which is what makes this failure mode so easy to miss until
        # someone actually reads the rendered output.
        idx = {str(h).strip().lower(): i for i, h in enumerate(headers) if h}

        def col(*names):
            for n in names:
                if n.lower() in idx:
                    return idx[n.lower()]
            return None

        c_num, c_type, c_chg = col("Number", "Circular Number", "Standard Number", "Document Number"), col("Type"), col("Change")
        if c_num is None:
            raise ValueError(
                f"'{sheet_name}' has no column recognizable as the regulation number - "
                f"looked for 'Number', 'Circular Number', 'Standard Number', or 'Document "
                f"Number' (case and stray whitespace don't matter, this matching already "
                f"ignores both) - every regulation number in the rendered document would "
                f"come out blank, which is worse than failing loudly here. Found these "
                f"headers instead: {', '.join(str(h) for h in headers if h)}. Rename the "
                f"column to one of the four accepted names and try again.")
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
            change = cell(row, c_chg)
            if c_curver is not None or c_plnver is not None:
                ver = version_display(cell(row, c_curver), cell(row, c_plnver), change)
            else:
                ver = cell(row, c_ver_legacy)
            it = {
                "type": infer_type(num, cell(row, c_type)),
                "change": change,
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
                "version": version_display(r["currentVersion"], r["plannedVersion"], r["change"]),
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
def _icon_svg(key):
    return (f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" fill="none" '
            f'stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round">'
            f'{ICON_PATHS[key]}</svg>')


def _bullet_html(it):
    suffix = f" ({esc(it['number'])})" if it["number"] else ""
    return f"<li>{esc(it['title_en'])}{suffix}</li>"


def _card_html(label, kind, change, icon, items):
    """Always renders - a fixed set of boxes (4 for Circulars, 3 for
    Standards) shown every time regardless of content, per explicit
    preference over the alternative (dropping empty buckets entirely)
    tried first. A bucket with nothing in it gets a clean "nothing to
    report" body instead of an empty bullet list, so it still reads as
    an intentional, organized empty state rather than a broken box."""
    bullets = [it for it in items if it["type"] == kind and it["change"] == change]
    color = STATE_COLOR[change]
    if bullets:
        body_cls, body = "card-b", f'<ul>{"".join(_bullet_html(it) for it in bullets)}</ul>'
    else:
        body_cls, body = "card-b empty", '<p class="card-empty">Nothing to report this quarter.</p>'
    return (f'<div class="card"><div class="card-h" style="background:{color}">'
            f'<span>{esc(label)}</span><span class="cardicon">{_icon_svg(icon)}</span></div>'
            f'<div class="{body_cls}">{body}</div></div>')


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
    stripe = STATE_COLOR.get(it["change"], GREY)
    rows = []
    for n, (title, link, summary) in enumerate(langs):
        meta_cells = ""
        if n == 0:
            meta_cells = (f'<td class="c-gov" rowspan="{span}" style="border-left-color:{stripe}">{esc(it["gov"])}</td>'
                          f'<td class="c-no" rowspan="{span}">{esc(it["number"])}</td>')
        ver_cell = f'<td class="c-ver" rowspan="{span}">{esc(it["version"])}</td>' if n == 0 else ""
        rows.append(f'<tr>{meta_cells}<td class="c-title">{_title_cell(title, link)}</td>{ver_cell}'
                    f'<td class="c-sum">{esc(summary)}</td></tr>')
    return "".join(rows)


def _table_head_html(kind):
    color = FAMILY_COLOR[kind]
    # c-gov's border-left is set to the same color as its own background -
    # continuing the same 6pt-wide colored edge the body rows use (see
    # .c-gov CSS), just in the family color here since the header has no
    # per-row Change state of its own. Matching width/color at this
    # boundary is what makes it read as one continuous line down the
    # table rather than stopping abruptly at the header.
    th = lambda cls, text, extra="": f"<th class='{cls}' style='background:{color}{extra}'>{text}</th>"
    gov_th = th("c-gov", "Governance responsibility", f";border-left-color:{color}")
    return ("<thead><tr>" + gov_th + th("c-no", f"{kind} number")
            + th("c-title", f"{kind} title") + th("c-ver", "Version") + th("c-sum", "Summary") + "</tr></thead>")


# The color legend - a single row explaining the four state colors
# (New/Upgrade/Update/Archived). Shown from page 2 onward only: page 1
# (the cover) already establishes what each color means through the
# cards themselves - repeating it there would be redundant, not
# reassuring. Family colors (Circulars=teal, Standards=blue) aren't in
# this legend either - each detail page's own colored header/border
# already shows which family it is, on that same page. Built once at
# import time rather than per page, since it never varies.
_LEGEND_STATES = [("New", "New"), ("Upgrade", "Upgrade"), ("Update", "Update"), ("Archiving", "Archived")]

LEGEND_HTML = '<div class="legend"><div class="legend-row">' + "".join(
    f'<span class="legend-item"><span class="legend-dot" style="background:{STATE_COLOR[chg]}"></span>{esc(label)}</span>'
    for chg, label in _LEGEND_STATES) + '</div></div>'


def _footer_html(page_num, total_pages, meta):
    legend = LEGEND_HTML if page_num > 1 else ""
    return (legend
            + f'<div class="foot"><span class="pg">Page {page_num} of {total_pages}</span>'
            f'<span class="meta">{esc(meta.get("Classification","Restricted"))} | '
            f'\u00a9 Siemens {esc(meta.get("Copyright year") or meta.get("Publication date","")[-4:])} | '
            f'{esc(meta.get("Governance Office",""))} | {esc(meta.get("Publication date",""))}</span>'
            f'{WORDMARK}</div>')


def _band_html(css_class, label_cls, header_text, cards_html):
    """cards_html: the full fixed set of card HTML strings for this
    band (always 4 for Circulars, always 3 for Standards - see
    _cover_pages_html), in CIRCULAR_CARDS/STANDARD_CARDS order. A
    trailing lone card (Standards' 3rd) already carries the "wide"
    class from _cover_pages_html, so it spans both grid columns
    instead of sitting next to a visibly blank cell - a fixed
    structural fact now that the card count itself is fixed, not a
    per-render decision."""
    sidebar = (f'<div class="band-sidebar"><span class="band-bar {label_cls}"></span>'
              f'<span class="band-label {label_cls}">{esc(header_text)}</span></div>')
    body = f'<div class="grid">{"".join(cards_html)}</div>'
    return f'<div class="band {css_class}">{sidebar}<div class="band-content">{body}</div></div>'


def _cover_pages_html(items, quarter, meta):
    subtitle = f"New, revised or archived Siemens Circulars and Siemens Standards in {esc(quarter)}"

    def _bucket_texts(kind, change):
        return [f"{b['title_en']} ({b['number']})" if b["number"] else b["title_en"]
                for b in items if b["type"] == kind and b["change"] == change]

    def _build_cards(specs, kind):
        return [_card_html(label, kind, chg, icon, items) for label, chg, icon in specs]

    circ_cards = _build_cards(CIRCULAR_CARDS, "Circular")
    std_cards = _build_cards(STANDARD_CARDS, "Standard")

    circ_bodies = [_card_body_height_mm(_bucket_texts("Circular", chg)) for _, chg, _ in CIRCULAR_CARDS]
    std_bodies = [_card_body_height_mm(_bucket_texts("Standard", chg)) for _, chg, _ in STANDARD_CARDS]
    circ_h = _band_height_mm(circ_bodies)
    std_h = _band_height_mm(std_bodies)

    note = (f'For more details regarding the changes made, please have a look at the following pages.<br>'
            f'If you are interested in previous summaries of regulatory changes, please feel free to contact us '
            f'(<a href="{esc(meta.get("Contact unit link","#"))}">{esc(meta.get("Contact unit",""))}</a>, '
            f'<a href="mailto:{esc(meta.get("Contact mail",""))}">{esc(meta.get("Contact mail",""))}</a>).')

    circ_band = _band_html("circulars", "circ", "CIRCULARS", circ_cards)
    std_band = _band_html("standards", "std", "STANDARDS", std_cards)

    pages = []
    if circ_h + std_h <= COVER_BANDS_BUDGET_MM:
        body = (f'<h1>Summary of regulatory changes</h1><div class="subtitle">{subtitle}</div>'
                f'{circ_band}{std_band}<div class="note">{note}</div>')
        pages.append(f'<div class="page cover">{body}</div>')
    else:
        body1 = f'<h1>Summary of regulatory changes</h1><div class="subtitle">{subtitle}</div>{circ_band}'
        pages.append(f'<div class="page cover">{body1}</div>')
        body2 = f'<h1>Summary of regulatory changes</h1><div class="subtitle">{subtitle}</div>{std_band}<div class="note">{note}</div>'
        pages.append(f'<div class="page cover">{body2}</div>')
    return pages


def _detail_pages_html(items, quarter):
    pages = []
    for kind, noun in SECTIONS:
        entries = [it for it in items if it["type"] == kind]
        if not entries:
            continue
        # Stable sort: New, Upgrade, Update, Archiving, in that order -
        # matches the cover cards' own bucket order. Ties (same Change)
        # keep their original relative order from the source data.
        entries = sorted(entries, key=lambda it: SECTION_STATE_ORDER.index(it["change"])
                         if it["change"] in SECTION_STATE_ORDER else len(SECTION_STATE_ORDER))
        family_color = FAMILY_COLOR[kind]
        subtitle = f"Changes to {esc(noun)} in {esc(quarter)}"
        chunks = _chunk_by_height(entries, TABLE_BODY_BUDGET_MM, MAX_ROWS_PER_TABLE_PAGE)
        head = _table_head_html(kind)
        for chunk in chunks:
            rows = "".join(_block_rows_html(it) for it in chunk)
            body = (f'<h1>Summary of regulatory changes</h1><div class="subtitle">{subtitle}</div>'
                    f'<div class="table-wrap" style="border-color:{family_color}">'
                    f'<table>{head}<tbody>{rows}</tbody></table></div>')
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
                         "explicitly (e.g. quarter=\"Q4/2026\" - quarter number first, then a "
                         "4-digit year, separated by / or - or a space; \"2026 Q4\" or \"Q4 FY26\" "
                         "will NOT parse), or check that 'Planned Quarter' in Circulars/Roadmap "
                         "actually has values in that same Q#/YYYY shape - a blank column, or one "
                         "converted to an Excel date by accident, both look like no quarter at "
                         "all to this check.")
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


def _run_cli():
    """Only ever called from the __main__ guard below, and only when
    sys.argv[1] already looks like a genuine .xlsx path - never calls
    sys.exit(). That matters more than it looks: __name__ == "__main__"
    is true not just for a real `python generate_summary.py ...`
    invocation, but also whenever this file is exec()'d into a hosted
    Code Interpreter session - exactly how the chat-bot deployment is
    documented (bot_instructions.md) to load it, specifically to avoid
    Python's import bytecode cache going stale across turns. Confirmed
    concretely, not just in theory: running
    `exec(open("generate_summary.py").read())` in a plain interactive
    session reliably hit this block with whatever sys.argv the host
    process happened to have (not file paths at all) and sys.exit(1)
    tore down the entire session before load_from_xlsx/build were ever
    reached - the exact failure a person hit for real, not a
    hypothetical. There is no reliable way to distinguish "real
    terminal invocation" from "exec'd into a notebook-like host" by
    inspecting __name__ or sys.argv alone, so the fix isn't a cleverer
    check - it's making every path through this block harmless to a
    hosted session: if argv doesn't already look like a real
    invocation, do nothing and let exec() continue normally, rather
    than ever exiting the whole process."""
    out_dir_arg = sys.argv[2] if len(sys.argv) > 2 else "."
    quarter_arg = sys.argv[3] if len(sys.argv) > 3 else None
    summary_data = load_from_xlsx(sys.argv[1], quarter=quarter_arg)
    html_out, pdf_out = build(summary_data, out_dir=out_dir_arg)
    print("HTML written to:", html_out)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith(".xlsx"):
        _run_cli()
    elif len(sys.argv) == 1:
        # A genuine `python generate_summary.py` with nothing else on
        # the line - most likely a real terminal user who forgot the
        # path, so worth a hint. stderr, not stdout, and no sys.exit():
        # harmless even if this fires from an exec()'d hosted session
        # instead (sys.argv there is whatever the host happened to
        # have, not usually empty like this, but better safe).
        print("Usage: python generate_summary.py bus_data.xlsx [out_dir] [quarter]",
              file=sys.stderr)
    # Any other shape of sys.argv (2+ entries, first one not .xlsx) is
    # left completely alone - almost certainly a hosted session's own
    # argv, not a person's typo, and printing anything here risks
    # looking like a real error to whatever's watching stdout.