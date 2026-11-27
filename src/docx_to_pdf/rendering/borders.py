"""Paragraph borders and shading (w:pBdr, w:pPr/w:shd): heading rules, header
underlines, boxed and shaded paragraphs. reportlab's ParagraphStyle can only draw a
full box in one colour, so a wrapper draws each side Word sets.

ponytail: consecutive paragraphs with the same borders are boxed one by one, where
Word draws one box around the group; "double" is two thin lines, other line styles
(wave, 3-D) are drawn as single lines.
"""

from __future__ import annotations

from reportlab.lib.colors import HexColor
from reportlab.platypus import Flowable

from .styles import Border, ParagraphFormat

_DASHES = {"dashed": (3, 2), "dotted": (1, 1.5), "dotDash": (3, 1.5, 1, 1.5), "dashSmallGap": (3, 1)}


class Bordered(Flowable):
    """A paragraph's flowable with its borders and shading. The lines sit `space`
    outside the text; shading spans the indents, as in Word."""

    def __init__(self, inner: Flowable, paragraph_format: ParagraphFormat) -> None:
        super().__init__()
        self.inner = inner
        self.format = paragraph_format
        self.keepWithNext = inner.getKeepWithNext()
        borders = paragraph_format.borders
        self._top = _reach(borders.get("top"))
        self._bottom = _reach(borders.get("bottom"))

    def wrap(self, availWidth: float, availHeight: float) -> tuple[float, float]:
        _, height = self.inner.wrap(availWidth, availHeight - self._top - self._bottom)
        self.width, self.height = availWidth, height + self._top + self._bottom
        return self.width, self.height

    def split(self, availWidth: float, availHeight: float) -> list[Flowable]:
        parts = self.inner.split(availWidth, availHeight - self._top - self._bottom)
        return [Bordered(part, self.format) for part in parts]

    def getSpaceBefore(self) -> float:
        return self.inner.getSpaceBefore()

    def getSpaceAfter(self) -> float:
        return self.inner.getSpaceAfter()

    def draw(self) -> None:
        canvas, borders = self.canv, self.format.borders
        left = self.format.left - _reach(borders.get("left"))
        right = self.width - self.format.right + _reach(borders.get("right"))
        if self.format.shading:
            canvas.setFillColor(HexColor(self.format.shading))
            canvas.rect(left, 0, right - left, self.height, stroke=0, fill=1)
        self.inner.drawOn(canvas, 0, self._bottom)
        lines = {"top": (left, self.height, right, self.height), "bottom": (left, 0, right, 0),
                 "left": (left, 0, left, self.height), "right": (right, 0, right, self.height)}
        for side, border in borders.items():
            # a double line's second line goes inside: down from the top, left from the right
            _line(canvas, border, *lines[side], second_below=side in ("top", "right"))


def _reach(border: Border | None) -> float:
    """How far a border stands out from the text: its spacing plus its line."""
    return border.space + border.width if border else 0.0


def _line(canvas, border: Border, x0: float, y0: float, x1: float, y1: float, second_below: bool) -> None:
    canvas.saveState()
    canvas.setStrokeColor(HexColor(border.color))
    canvas.setDash(*([_DASHES[border.style]] if border.style in _DASHES else []))
    if border.style == "double":
        canvas.setLineWidth(max(border.width / 3, 0.25))
        gap = border.width * 2 / 3 * (-1 if second_below else 1)
        vertical = x0 == x1
        canvas.line(x0, y0, x1, y1)
        canvas.line(x0 + gap * vertical, y0 + gap * (not vertical), x1 + gap * vertical, y1 + gap * (not vertical))
    else:
        canvas.setLineWidth(border.width)
        canvas.line(x0, y0, x1, y1)
    canvas.restoreState()
