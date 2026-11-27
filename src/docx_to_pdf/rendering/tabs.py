"""Tab stops: a paragraph cut at its tabs, each piece placed at its stop.

reportlab's Paragraph has no tab stops, so `inline` marks each `w:tab`/`w:ptab` in
the markup (outside any font tag), and a paragraph that fits on one line is laid
out here: left, centre, right and decimal stops, with leaders (the dots of a table
of contents). Positions are from the left margin, as in Word; past the last stop,
the document's default stops apply (every half inch unless settings.xml says).

ponytail: a tabbed paragraph that wraps keeps a fixed gap per tab (`as_gaps`);
laying out tabs across wrapped lines needs a line breaker of our own.
"""

from __future__ import annotations

import html
import re

from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import Flowable, Paragraph

from .paragraphs import first_baseline, make_paragraph
from .styles import TabStop

TAB = "\t"
_PTAB = "\x0b"  # control characters can't occur in XML text, so they are safe marks
_MARKS = re.compile(r"(\t|\x0b[^\x0b]*\x0b)")
_GAP = "&nbsp;" * 4
_LEADERS = {"dot": ".", "hyphen": "-", "underscore": "_", "heavy": "_", "middleDot": "·"}
_TAGS = re.compile(r"<[^>]*>")


def ptab(alignment: str | None, relative_to: str | None, leader: str | None) -> str:
    """The mark for a positional tab: aligned to the margin or indent, not a stop."""
    return f"{_PTAB}{alignment or 'left'}:{relative_to or 'margin'}:{leader or 'none'}{_PTAB}"


def has_tabs(markup: str) -> bool:
    return _MARKS.search(markup) is not None


def as_gaps(markup: str) -> str:
    """Each tab as a fixed gap, for a paragraph that can't be laid out at its stops."""
    return _MARKS.sub(_GAP, markup)


class TabbedLine(Flowable):
    """One line of pieces, each a Paragraph drawn at its own x, on a shared baseline
    (header text stands on the same baseline as the logo beside it)."""

    def __init__(self, placed: list[tuple[float, Paragraph, float]], style: ParagraphStyle) -> None:
        super().__init__()
        self.placed = placed  # (x, paragraph, width)
        self.style = style
        self.keepWithNext = style.keepWithNext
        self._baseline = 0.0

    def wrap(self, availWidth: float, availHeight: float) -> tuple[float, float]:
        heights = [paragraph.wrap(width + 2, availHeight)[1] for _, paragraph, width in self.placed]
        baselines = [first_baseline(paragraph) for _, paragraph, _ in self.placed]
        self._baseline = max(baselines, default=0.0)
        self.width = availWidth
        self.height = max((self._baseline - baseline + height for baseline, height in zip(baselines, heights)),
                          default=self.style.leading)
        return self.width, self.height

    def draw(self) -> None:
        for x, paragraph, _ in self.placed:
            top = self.height - (self._baseline - first_baseline(paragraph))
            paragraph.drawOn(self.canv, x, top - paragraph.height)

    def getSpaceBefore(self) -> float:
        return self.style.spaceBefore

    def getSpaceAfter(self) -> float:
        return self.style.spaceAfter


def tabbed_line(markup: str, style: ParagraphStyle, stops: tuple[TabStop, ...], width: float,
                default_interval: float = 36.0) -> TabbedLine | None:
    """The paragraph laid out at its tab stops, or None if it isn't one line that fits
    in `width` (the caller then falls back to `as_gaps`). Past the last stop, tabs go
    to multiples of `default_interval` (settings.xml w:defaultTabStop)."""
    parts = _MARKS.split(markup)  # text, mark, text, mark, ..., text
    if any("<br/>" in text for text in parts[::2]):
        return None
    plain = ParagraphStyle("tab-piece", parent=style, leftIndent=0, rightIndent=0, firstLineIndent=0,
                           alignment=TA_LEFT)
    x = style.leftIndent + style.firstLineIndent
    # A hanging indent is also a stop: "Step 5:<tab>text" lines up with the wrapped lines.
    hanging = style.leftIndent if style.firstLineIndent < 0 else None
    placed: list[tuple[float, Paragraph, float]] = []
    for index in range(0, len(parts), 2):
        text, piece_width = parts[index], 0.0
        paragraph = None
        if _TAGS.sub("", text) or "<img " in text:
            paragraph = make_paragraph(text, plain)
            paragraph.wrap(1e6, 1e6)
            piece_width = max(paragraph.getActualLineWidths0(), default=0.0)
        if index:
            stop = _stop(parts[index - 1], x, stops, hanging, style, width, default_interval)
            start = max(x, stop.position - _offset(stop.alignment, text, piece_width, style))
            if stop.leader and start - x > 1:
                placed.append(_leader(stop.leader, x, start, plain))
            x = start
        if paragraph is not None:
            placed.append((x, paragraph, piece_width))
        x += piece_width
        if x > width + 0.5:
            return None
    return TabbedLine(placed, style)


def _stop(mark: str, x: float, stops: tuple[TabStop, ...], hanging: float | None, style: ParagraphStyle,
          width: float, default_interval: float) -> TabStop:
    if mark != TAB:  # a positional tab
        alignment, relative_to, leader = mark.strip(_PTAB).split(":")
        left, right = (style.leftIndent, width - style.rightIndent) if relative_to == "indent" else (0.0, width)
        position = {"center": (left + right) / 2, "right": right}.get(alignment, left)
        return TabStop(position, alignment, None if leader == "none" else leader)
    candidates = [stop for stop in stops if stop.position > x + 0.01]
    if hanging is not None and hanging > x + 0.01:
        candidates.append(TabStop(hanging, "left", None))
    if candidates:
        return min(candidates, key=lambda stop: stop.position)
    return TabStop((x // default_interval + 1) * default_interval, "left", None)


def _offset(alignment: str, text: str, piece_width: float, style: ParagraphStyle) -> float:
    """How far left of the stop the piece starts."""
    if alignment == "right":
        return piece_width
    if alignment == "center":
        return piece_width / 2
    if alignment == "decimal":
        plain = html.unescape(_TAGS.sub("", text))
        if "." not in plain:
            return piece_width  # a whole number ends at the stop
        whole = stringWidth(plain, style.fontName, style.fontSize) or 1.0
        return piece_width * stringWidth(plain[:plain.index(".")], style.fontName, style.fontSize) / whole
    return 0.0


def _leader(leader: str, start: float, end: float, style: ParagraphStyle) -> tuple[float, Paragraph, float]:
    """Leader characters filling start..end, flush against the text after them."""
    character = _LEADERS.get(leader, ".")
    size = stringWidth(character, style.fontName, style.fontSize) or 1.0
    count = int((end - start) / size)
    paragraph = Paragraph(html.escape(character * count), style)
    return end - count * size, paragraph, count * size
