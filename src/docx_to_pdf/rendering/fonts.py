"""Word fonts -> the open-licensed fonts bundled in `docx_to_pdf/fonts/`.

Where a free font has the same character widths as Word's (measured, docs/DESIGN.md
2026-11-26 "fonts"), lines break where Word breaks them: Carlito for Calibri,
Liberation Sans/Serif/Mono for Arial, Times New Roman and Courier New. Caladea
stands in for Cambria but measured about 6% narrower. Other fonts get the
Liberation font of their class (serif, monospace, else sans).

A character the chosen font lacks is drawn in DejaVu Sans; one neither has is
drawn as "?" and reported. Scripts that need shaping (right-to-left, Indic,
Thai...) are always reported, never drawn: reportlab draws characters one by one,
in order, so they would come out reversed or unjoined while looking "rendered".
ponytail: add a shaping engine (e.g. uharfbuzz) and bidi reordering to draw those.
"""

from __future__ import annotations

import struct
import unicodedata
from dataclasses import dataclass
from functools import cache
from pathlib import Path

from reportlab.pdfbase.pdfmetrics import getFont, registerFont
from reportlab.pdfbase.ttfonts import TTFont

UNRENDERABLE = "?"
FONT_DIR = Path(__file__).resolve().parent.parent / "fonts"

# family -> file stems for (regular, bold, italic, bold italic)
_FAMILIES = {
    "Carlito": ("Carlito-Regular", "Carlito-Bold", "Carlito-Italic", "Carlito-BoldItalic"),
    "Caladea": ("Caladea-Regular", "Caladea-Bold", "Caladea-Italic", "Caladea-BoldItalic"),
    "LiberationSans": ("LiberationSans-Regular", "LiberationSans-Bold", "LiberationSans-Italic",
                       "LiberationSans-BoldItalic"),
    "LiberationSerif": ("LiberationSerif-Regular", "LiberationSerif-Bold", "LiberationSerif-Italic",
                        "LiberationSerif-BoldItalic"),
    "LiberationMono": ("LiberationMono-Regular", "LiberationMono-Bold", "LiberationMono-Italic",
                       "LiberationMono-BoldItalic"),
    # Fallback only: regular and bold are enough to keep a missing character's weight.
    "DejaVuSans": ("DejaVuSans", "DejaVuSans-Bold", "DejaVuSans", "DejaVuSans-Bold"),
}
_FALLBACK = "DejaVuSans"

_SUBSTITUTES = {
    "calibri": "Carlito", "calibri light": "Carlito",
    "cambria": "Caladea",
    "arial": "LiberationSans", "helvetica": "LiberationSans", "arial narrow": "LiberationSans",
    "times new roman": "LiberationSerif", "times": "LiberationSerif",
    "courier new": "LiberationMono", "courier": "LiberationMono", "consolas": "LiberationMono",
    "lucida console": "LiberationMono",
}
_SERIF = ("roman", "georgia", "garamond", "book antiqua", "palatino", "baskerville", "century", "serif")
_MONOSPACE = ("mono", "courier", "consolas", "console", "menlo")

# Word's single line height for its own fonts, as a multiple of the font size:
# (usWinAscent + usWinDescent) / unitsPerEm of the Microsoft font, measured once
# (docs/DESIGN.md). Used instead of the substitute's own metrics because the layout
# being reproduced is Word's. Fonts not listed use their substitute's metrics.
_WORD_LINE_HEIGHTS = {
    "calibri": 1.2207, "calibri light": 1.2207, "cambria": 1.1724, "arial": 1.1172,
    "times new roman": 1.1074, "courier new": 1.1328,
}
# The Microsoft font each bundled family stands in for, for its line height above.
_ORIGINALS = {"Carlito": "calibri", "Caladea": "cambria", "LiberationSans": "arial",
              "LiberationSerif": "times new roman", "LiberationMono": "courier new"}
# Scripts drawn wrongly without shaping or bidi reordering (beyond the right-to-left
# ones, which `unicodedata.bidirectional` catches): Indic, Sinhala, Thai, Lao,
# Tibetan, Myanmar, Khmer, Mongolian.
_SHAPED_RANGES = ((0x0900, 0x0DFF), (0x0E00, 0x0FFF), (0x1000, 0x109F), (0x1780, 0x18AF))


@dataclass(frozen=True)
class Segment:
    face: str | None  # None: the run's own face; else a fallback face
    text: str


def family_for(word_font: str | None) -> str:
    name = (word_font or "").strip().lower()
    if name in _SUBSTITUTES:
        return _SUBSTITUTES[name]
    if any(word in name for word in _MONOSPACE):
        return "LiberationMono"
    if any(word in name for word in _SERIF):
        return "LiberationSerif"
    return "LiberationSans"  # Aptos (Word 365's default) and other sans fonts land here


def face_for(word_font: str | None, bold: bool, italic: bool) -> str:
    """The registered reportlab font name to draw `word_font` with."""
    return _register(family_for(word_font), bold + 2 * italic)


def line_height(word_font: str | None) -> float:
    """Word's single line spacing for `word_font`, as a multiple of the font size."""
    name = (word_font or "").strip().lower()
    return _WORD_LINE_HEIGHTS.get(name) or _win_line_height(_FAMILIES[family_for(word_font)][0])


def split_by_coverage(text: str, face: str, unrenderable: list[str]) -> list[Segment]:
    """`text` cut into runs drawn in `face` (Segment.face None) or in the fallback
    face; characters neither can draw become "?" and are appended to `unrenderable`."""
    fallback = _register(_FALLBACK, 1 if face.endswith(("-Bold", "-BoldItalic")) else 0)
    primary, secondary = _coverage(face), _coverage(fallback)
    segments: list[Segment] = []
    for character in unicodedata.normalize("NFC", text):
        if _is_invisible(character):
            continue  # Word draws nothing for these either
        if _needs_shaping(character):
            unrenderable.append(character)
            chosen, character = None, UNRENDERABLE
        elif ord(character) in primary or character.isspace():
            chosen = None
        elif ord(character) in secondary:
            chosen = fallback
        else:
            unrenderable.append(character)
            chosen, character = None, UNRENDERABLE
        if segments and segments[-1].face == chosen:
            segments[-1] = Segment(chosen, segments[-1].text + character)
        else:
            segments.append(Segment(chosen, character))
    return segments


def face_covering(text: str, face: str) -> str:
    """`face`, or the fallback face when `face` lacks one of `text`'s characters: for
    text drawn in a single font, like a list label ("●", "✓")."""
    return next((segment.face for segment in split_by_coverage(text, face, []) if segment.face), face)


def _is_invisible(character: str) -> bool:
    """Format characters (zero-width space and joiners, direction marks, BOM, soft
    hyphen) and variation selectors (the U+FE0F in "✔️"): common in text pasted from
    email or the web, and never drawn. Dropped before the shaping check, which would
    otherwise report a direction mark as right-to-left text."""
    code = ord(character)
    return (unicodedata.category(character) == "Cf" or 0xFE00 <= code <= 0xFE0F
            or 0xE0100 <= code <= 0xE01EF)


def _needs_shaping(character: str) -> bool:
    if unicodedata.bidirectional(character) in ("R", "AL"):
        return True
    if unicodedata.category(character) == "Mn":  # a combining mark NFC could not compose
        return True
    code = ord(character)
    return any(start <= code <= end for start, end in _SHAPED_RANGES)


@cache
def _register(family: str, style: int) -> str:
    """Register the font file once per process; a missing or broken bundled font
    raises (a packaging bug), rather than drawing in some other font."""
    stem = _FAMILIES[family][style]
    font = TTFont(stem, str(FONT_DIR / f"{stem}.ttf"))
    # reportlab takes ascent/descent from the typo metrics, which lines that size
    # themselves (autoLeading) then use. Word's line is the original font's win
    # metrics: Calibri 1.2207, though Carlito's own are 1.30 (docs/DESIGN.md).
    # So the height is Word's figure, split in this font's ascent:descent ratio.
    win_ascent, win_descent, units_per_em = _win_metrics(stem)
    height = _WORD_LINE_HEIGHTS.get(_ORIGINALS.get(family, "")) or (win_ascent + win_descent) / units_per_em
    share = win_ascent / (win_ascent + win_descent)
    font.face.ascent, font.face.descent = 1000 * height * share, -1000 * height * (1 - share)
    registerFont(font)
    return stem


@cache
def _coverage(face: str) -> frozenset[int]:
    return frozenset(getFont(face).face.charToGlyph)


def _win_line_height(stem: str) -> float:
    """(usWinAscent + usWinDescent) / unitsPerEm."""
    win_ascent, win_descent, units_per_em = _win_metrics(stem)
    return (win_ascent + win_descent) / units_per_em


@cache
def _win_metrics(stem: str) -> tuple[int, int, int]:
    """(usWinAscent, usWinDescent, unitsPerEm), read from the font file's tables."""
    data = (FONT_DIR / f"{stem}.ttf").read_bytes()
    table_count = struct.unpack(">H", data[4:6])[0]
    offsets = {
        data[12 + 16 * index:16 + 16 * index]: struct.unpack(">I", data[20 + 16 * index:24 + 16 * index])[0]
        for index in range(table_count)
    }
    head, os2 = offsets[b"head"], offsets[b"OS/2"]
    units_per_em = struct.unpack(">H", data[head + 18:head + 20])[0]
    win_ascent, win_descent = struct.unpack(">HH", data[os2 + 74:os2 + 78])
    return win_ascent, win_descent, units_per_em
