"""The one place a text Paragraph is made, so lines sized one by one get it right."""

from __future__ import annotations

from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph


class _LineSizedParagraph(Paragraph):
    """A paragraph whose lines are sized one by one (autoLeading "max"): it holds
    inline pictures or mixed font sizes, and Word gives each line the height of
    its own largest item (docs/DESIGN.md). Two corrections to reportlab:

    - a line's height is multiplied by the paragraph's line spacing (`lineMultiple`
      on the style, 1.15 in many templates), the extra going below the text;
    - reportlab puts the first baseline one font size below the top
      (rl_config.paraFontSizeHeightOffset, read once at import), so a picture taller
      than the text would stick out above the paragraph: drawing is dropped by the
      difference.
    """

    def wrap(self, availWidth: float, availHeight: float) -> tuple[float, float]:
        width, height = super().wrap(availWidth, availHeight)
        multiple = getattr(self.style, "lineMultiple", 1.0)
        if multiple != 1.0 and getattr(self.blPara, "kind", 0) == 1:
            for line in self.blPara.lines:
                if not getattr(line, "spaced", False):
                    line.descent -= (multiple - 1) * (line.ascent - line.descent)
                    line.spaced = True
            height = self.height = sum(max(line.ascent - line.descent, self.style.leading)
                                       for line in self.blPara.lines)
        return width, height

    def drawOn(self, canvas, x: float, y: float, _sW: float = 0) -> None:
        super().drawOn(canvas, x, y - _picture_drop(self), _sW)


def _picture_drop(paragraph: Paragraph) -> float:
    lines = paragraph.blPara
    if getattr(lines, "kind", 0) == 1 and lines.lines:
        first = lines.lines[0]
        return max(getattr(first, "ascent", 0.0) - first.fontSize, 0.0)
    return 0.0


def first_baseline(paragraph: Paragraph) -> float:
    """How far below its top a wrapped paragraph draws its first baseline."""
    lines = paragraph.blPara
    if getattr(lines, "kind", 0) == 1 and lines.lines:
        return lines.lines[0].fontSize + (_picture_drop(paragraph) if isinstance(paragraph, _LineSizedParagraph)
                                          else 0.0)
    return lines.fontSize


def make_paragraph(markup: str, style: ParagraphStyle, label: str | None = None) -> Paragraph:
    if "<img " in markup or getattr(style, "autoLeading", "") == "max":
        sized = ParagraphStyle("line-sized", parent=style, autoLeading="max")
        sized.lineMultiple = getattr(style, "lineMultiple", 1.0)
        return _LineSizedParagraph(markup, sized, bulletText=label)
    return Paragraph(markup, style, bulletText=label)
