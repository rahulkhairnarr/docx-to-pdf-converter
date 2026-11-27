"""State shared while rendering one document, and the text guard.

The guard is the reason a missing construct can't lose text silently: every `w:t`
(text) element the renderer handles, by drawing it or by skipping it on purpose
(deleted or hidden text), is marked. Whatever is left unmarked when the PDF is built
was never reached, and is reported as a warning.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

from docx.opc.part import Part
from lxml.etree import _Element

from .numbering import Numbering, format_number
from .ooxml import qn
from .styles import StyleSheet

_TEXT = qn("w:t")
_FALLBACK = qn("mc:Fallback")  # an older copy of mc:Choice's content, drawn instead of it


class TextGuard:
    def __init__(self) -> None:
        self._handled: set[_Element] = set()  # holding the proxies keeps lxml identities stable
        self.unrenderable: Counter[str] = Counter()  # characters drawn as "?"
        self.not_drawn: Counter[str] = Counter()  # "image (EMF)", "chart", ...: shown as a placeholder
        self.tracked_changes = 0  # revisions shown as accepted

    def handled(self, text: _Element) -> None:
        self._handled.add(text)

    def skipped(self, element: _Element) -> None:
        """Mark all text below `element` as left out on purpose."""
        self._handled.update(element.iter(_TEXT))

    def unhandled_text(self, root: _Element) -> str:
        return "".join(
            text.text or ""
            for text in root.iter(_TEXT)
            if text not in self._handled and not any(a.tag == _FALLBACK for a in text.iterancestors())
        )


@dataclass
class FieldState:
    """Complex fields (fldChar begin / instrText / separate / end) nest, so a stack."""

    stack: list[dict] = field(default_factory=list)

    def in_page_field_result(self, page_fields: PageFields) -> bool:
        """Inside the cached result of a page-number field (replaced by the real value)."""
        return (bool(self.stack) and self.stack[-1]["in_result"]
                and page_fields.value(self.stack[-1]["instruction"]) is not None)


PAGE_FIELDS = ("PAGE", "NUMPAGES", "SECTIONPAGES")


@dataclass
class PageFields:
    """The page-number fields' values while drawing one header or footer."""

    values: dict[str, str]  # field name -> displayed value; totals are absent in the first pass
    wanted: set[str]  # shared by the whole render: names asked for but not yet known

    def value(self, instruction: str | None) -> str | None:
        """The text for a page-number field's instruction; None for any other field."""
        words = (instruction or "").split()
        name = words[0].upper() if words else ""
        if name not in PAGE_FIELDS:
            return None
        if name not in self.values:
            self.wanted.add(name)
        return self.values.get(name, "")


class NoteRegistry:
    """Footnotes and endnotes in the order the body cites them. Word numbers
    footnotes 1, 2, 3 and endnotes i, ii, iii by default."""

    def __init__(self) -> None:
        self.cited: dict[str, list[str]] = {"footnote": [], "endnote": []}

    def cite(self, kind: str, note_id: str) -> str:
        cited = self.cited[kind]
        if note_id not in cited:
            cited.append(note_id)
        return self.mark(kind, cited.index(note_id) + 1)

    @staticmethod
    def mark(kind: str, number: int) -> str:
        return format_number(number, "lowerRoman" if kind == "endnote" else "decimal")


@dataclass
class RenderContext:
    styles: StyleSheet
    numbering: Numbering
    part: Part  # the part whose relationships resolve image ids (document, header, footer)
    guard: TextGuard
    width: float  # available width in points
    height: float  # available height in points, to scale images that would not fit
    # Set while drawing a header/footer, so PAGE/NUMPAGES/SECTIONPAGES show real values.
    # ponytail: a page field in the body keeps its cached result; it needs the page at draw time.
    page_fields: PageFields | None = None
    fields: FieldState = field(default_factory=FieldState)
    notes: NoteRegistry = field(default_factory=NoteRegistry)
    note_mark: str | None = None  # set while rendering a note's own text, for its w:footnoteRef
    default_tab: float = 36.0  # settings.xml w:defaultTabStop, in points (Word's default 0.5")
