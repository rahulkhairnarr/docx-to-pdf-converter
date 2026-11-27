# Design notes

Why the converter is built the way it is, and the measurements behind its choices.
The code comments point here for the details.

## Why pure Python

The usual ways to convert `.docx` to PDF drive another program: Word (through COM,
Windows only, with an Office licence) or LibreOffice (`soffice --headless`, a
large install). Both give good fidelity. Both are also unavailable in many places
Python runs: containers, serverless functions, locked-down machines and CI.

This package needs neither. It reads the document with **python-docx** and draws
the PDF with **reportlab**, whose platypus layout engine handles paragraphs,
tables and page templates. Both are pure Python or ship wheels, with no system
libraries.

Alternatives that were considered:

- **weasyprint**, from HTML. It needs Pango/Cairo system libraries, and turning
  .docx into HTML loses page-level features (sections, headers, footers).
- **fpdf2**. This was not tried. reportlab was chosen for its mature layout of
  splitting tables and paragraphs.
- **docx → HTML in a browser engine.** That brings a browser along: the very kind
  of dependency this package avoids.

Legacy `.doc` and password-protected files are refused with a clear error. Both
are OLE binary containers that nothing pure-Python reads reliably.

## Nothing is lost silently

A converter that misses a construct usually drops its text without a word. Here
every `w:t` (text) element the renderer handles is marked, whether it draws the
text or skips it on purpose (deleted, hidden). After rendering, any unmarked text
in the body, the headers and footers drawn, or the notes is reported in a warning.

The same report counts:
- characters drawn as `?`;
- placeholders (charts, SmartArt, embedded documents, EMF/WMF);
- tracked changes shown as accepted.

The guard has found real gaps. Footnote text was once lost without any warning,
because notes live in a separate part the guard didn't check. That led to notes
being rendered, and to the guard covering every part that gets drawn.

## Fonts: measured, not guessed

Line breaks depend on character widths. Free fonts that are *metric-compatible*
with Microsoft's were measured against the originals: the advance width of each
character at 1000 units, compared with the Microsoft font on a Windows machine.

| Microsoft font | Substitute | Width difference |
|---|---|---|
| Calibri | Carlito | 0.00% |
| Arial | Liberation Sans | 0.00% |
| Times New Roman | Liberation Serif | 0.00% |
| Courier New | Liberation Mono | 0.00% |
| Cambria | Caladea | about −6% |

**Caladea is not a scaled Cambria.** The per-character ratios are not uniform:
digits are much narrower ("1" is 0.65 of Cambria's), while capital I is wider
(1.06). The 2013 original release measures the same, so no build of Caladea
matches Cambria. It is still the closest free look-alike.

**Line height** is Word's own figure for the original font, (usWinAscent +
usWinDescent) / unitsPerEm. The substitute's figure isn't used, because the two
can differ: Calibri is 1.2207, but Carlito's own win metrics give about 1.30.

| Font | Word's single line height |
|---|---|
| Calibri | 1.2207 |
| Cambria | 1.1724 |
| Arial | 1.1172 |
| Times New Roman | 1.1074 |
| Courier New | 1.1328 |

reportlab sizes self-sizing lines from a font's *typo* metrics. So each bundled
face is registered with its ascent and descent set to Word's line height, split
in the face's own ascent:descent ratio.

**Theme fonts.** Most Word documents name their fonts through the theme
(`w:asciiTheme="minorHAnsi"`). These resolve through the theme part's
`a:fontScheme`. Within one `w:rFonts`, the theme reference wins over the plain
name.

**Characters outside the font** are drawn in DejaVu Sans. Scripts that need
shaping are *reported, never drawn*: right-to-left scripts (bidi class R/AL),
Indic, Thai, Khmer and others, plus unmatched combining marks. reportlab draws
characters one by one, so such text would come out reversed or unjoined while
looking rendered.

Before the check, text is NFC-normalised, so a decomposed "é" doesn't count as a
bare combining mark. Invisible format characters (zero-width joiners, direction
marks, the BOM, the soft hyphen, variation selectors) are dropped, as Word draws
nothing for them.

reportlab 5 can shape text through `uharfbuzz` (the `shaping` extra) and has a
`bidi` extra. Supporting these scripts would need those, plus script fonts.

## Layout decisions, with the evidence

These were found by comparing the output against a reference rendering, page by
page. Glyph positions were measured with pypdfium2.

- **Widow/orphan control is on when absent.** Word writes `w:widowControl` only
  as `w:val="0"`. LibreOffice's DOCX import also defaults widows and orphans to
  2.
- **Odd/even headers follow the displayed page number**, not the physical page.
  So a section that restarts at 1 on physical page 2 gets the *odd* header.
- **NUMPAGES and SECTIONPAGES need totals.** The document is laid out twice, and
  only when such a field is present. Headers and footers don't reflow the body,
  so the second pass is final.
- **Page-number fields may have no cached result.** Some generators write
  `begin`/`separate`/`end` with nothing between them, all in one run. The number
  is drawn at `separate`.
- **Tab stops.** reportlab has none. A paragraph that fits on one line is cut at
  its tabs, and each piece is placed at its stop: left, centre, right or decimal,
  with leaders. Positions are measured from the margin. A hanging indent is an
  implicit stop. Default stops come from `w:defaultTabStop`.
- **Inline pictures are part of the line** (`<img>` markup), so a logo and the
  text beside it share a baseline. reportlab grows the line for the picture
  (`autoLeading="max"`). But it still puts the first baseline one font size
  below the top (`rl_config.paraFontSizeHeightOffset`, read at import), so the
  picture would overlap the line above. The paragraph's drawing is lowered by the
  difference.
- **Mixed font sizes.** Word sizes each line by its own largest font. Leading
  every line by the largest font in the paragraph pushed one test page 45pt too
  low. Lines now size themselves, and the paragraph's line-spacing multiple
  (often 1.15) is applied to each line.
- **A paragraph holding only a page break** keeps its mark before the break. It
  adds no blank line at the top of the next page.
- **Table cells.** Cell margins come from `w:tcMar`, then `w:tblCellMar`, then
  the table style, then Word's default (0 top and bottom, 0.08" at the sides).
  reportlab drops a cell's first space-before and last space-after, which Word
  keeps, so these are added to the cell padding. An earlier attempt used
  zero-height spacers, and splitting a table across pages then never finished:
  each split produced another zero-height part.
- **Keep-with-next.** reportlab's KeepTogether moves the whole group when it
  doesn't fit: a heading and its 45-row table jumped to the next page and left
  half a page empty. Word needs only the next item's first line or row on the
  page. `KeptWithNext` splits the next item in place.
- **Superscript and subscript** are drawn at 0.65× the size, raised 0.33× or
  lowered 0.14×. A run's own `<font size>` would otherwise override reportlab's
  reduction.
- **Bullets.** Real Unicode bullets (● ○ ■ ◆ ➢ ✓) are kept. Symbol and Wingdings
  codes (private-use characters, or Latin-1 letters such as "§") become "•",
  because they mean nothing outside those fonts.

## Known limits

The README lists them. The largest is floating images: text doesn't wrap around
them, which can push later content down by a page in image-heavy documents.
After that come footnotes (collected at the end, not at the foot of each page)
and equations (drawn as plain text).
