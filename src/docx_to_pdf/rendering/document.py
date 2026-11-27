"""A whole .docx as a PDF: one page template per section (page size, margins,
orientation), the body as flowables, and headers/footers drawn on each page."""

from __future__ import annotations

import io
from collections.abc import Iterable
from dataclasses import dataclass, replace

import docx
from docx.document import Document
from docx.enum.section import WD_SECTION_START
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part, XmlPart
from docx.oxml import parse_xml
from docx.section import Section
from lxml.etree import _Element
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    HRFlowable,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Spacer,
)

from . import fonts
from .blocks import render_block, render_blocks
from .context import NoteRegistry, PageFields, RenderContext, TextGuard
from .keeping import keep_with_next
from .numbering import Numbering, format_number
from .ooxml import attr, child, qn, twips_to_pt
from .styles import StyleSheet, theme_fonts

_DEFAULT_MARGIN = 72.0  # 1 inch, for a section that doesn't say
_DEFAULT_DISTANCE = 36.0  # header/footer distance from the page edge
_SAMPLE = 60  # characters of unrendered text quoted in a warning


@dataclass(frozen=True)
class RenderedDocument:
    pdf: bytes
    warnings: tuple[str, ...]


@dataclass(frozen=True)
class _Page:
    width: float
    height: float
    left: float
    right: float
    top: float
    bottom: float
    header: float
    footer: float
    columns: int = 1
    column_space: float = 36.0
    column_separator: bool = False

    @property
    def frame_width(self) -> float:
        return self.width - self.left - self.right

    @property
    def column_width(self) -> float:
        """ponytail: equal columns; unequal widths (w:col children) are not read."""
        return (self.frame_width - self.column_space * (self.columns - 1)) / self.columns

    def column_left(self, column: int) -> float:
        return self.left + column * (self.column_width + self.column_space)

    @property
    def frame_height(self) -> float:
        return self.height - self.top - self.bottom


@dataclass(frozen=True)
class _Totals:
    """Page counts from a first layout pass, for NUMPAGES and SECTIONPAGES."""

    pages: int
    section_pages: tuple[int, ...]


class _Renderer:
    """Holds what every part of one document shares: styles, list counters, the guard.
    One instance renders once: list counters and note numbers are consumed as it goes."""

    def __init__(self, document: Document, totals: _Totals | None = None) -> None:
        self.document = document
        self.totals = totals
        theme = _related_xml(document, RT.THEME)
        self.styles = StyleSheet(document.styles.element, theme_fonts(theme[1] if theme else None))
        numbering = _related_xml(document, RT.NUMBERING)
        self.numbering = Numbering(numbering[1] if numbering else None)
        self.guard = TextGuard()
        self.notes = NoteRegistry()
        self.note_parts = {"footnote": _related_xml(document, RT.FOOTNOTES),
                           "endnote": _related_xml(document, RT.ENDNOTES)}
        self.sections = list(document.sections)
        self.pages = [_page(section) for section in self.sections]
        self.odd_and_even = document.settings.odd_and_even_pages_header_footer
        self.default_tab = twips_to_pt(attr(child(document.settings.element, "w:defaultTabStop"))) or 36.0
        self.drawn_parts: dict[str, _Element] = {}  # header/footer parts drawn, for the guard
        self.page_number = 0  # the displayed number of the page being drawn
        self.section_pages = [0] * len(self.sections)  # pages drawn per section so far
        self.wanted_fields: set[str] = set()  # page fields whose value needs a total

    def context(self, section_index: int, part: Part, page_fields: PageFields | None = None) -> RenderContext:
        """Body content is laid out a column wide; headers and footers (with
        `page_fields`) span the page between the margins."""
        page = self.pages[section_index]
        width = page.frame_width if page_fields else page.column_width
        return RenderContext(self.styles, self.numbering, part, self.guard, width, page.frame_height,
                             page_fields, notes=self.notes, default_tab=self.default_tab)

    def render(self) -> RenderedDocument:
        templates = [self._template(index) for index in range(len(self.sections))]
        output = io.BytesIO()
        pdf = BaseDocTemplate(output, pagesize=templates[0].pagesize, pageTemplates=templates,
                              title=self.document.core_properties.title or "", creator="docx-to-pdf",
                              # else every page declares Helvetica, reportlab's default
                              initialFontName=fonts.face_for(None, False, False))
        pdf.build(self._story() or [Spacer(0, 0)])
        roots = [self.document.element.body, *self.drawn_parts.values()]
        roots += [notes for _, notes in filter(None, self.note_parts.values())]
        return RenderedDocument(output.getvalue(), self._warnings(roots))

    def _story(self) -> list:
        story: list = []
        section_index = 0
        context = self.context(0, self.document.part)
        for element in self.document.element.body:
            story.extend(render_block(element, context))
            if element.tag == qn("w:p") and child(element, "w:pPr", "w:sectPr") is not None:
                section_index += 1  # this paragraph ends a section
                if section_index < len(self.sections):
                    story.append(NextPageTemplate(f"section-{section_index}"))
                    # ponytail: a continuous break can't change the page size mid-page; the
                    # new section's page setup starts on the next page.
                    if self.sections[section_index].start_type != WD_SECTION_START.CONTINUOUS:
                        story.append(PageBreak())
                    context = self.context(section_index, self.document.part)
        return keep_with_next(story + self._notes(context))

    def _notes(self, context: RenderContext) -> list:
        """The cited footnotes, then endnotes, after the body behind a short rule.
        ponytail: footnotes go at the end, not at the foot of their page; placing them
        needs a frame per page that shrinks by the notes cited on it."""
        flowables: list = []
        for kind, cited in self.notes.cited.items():
            if (found := self.note_parts[kind]) is None:
                continue
            part, root = found
            notes = {attr(note, "w:id"): note for note in root.iterfind(qn(f"w:{kind}"))}
            for note_id, note in notes.items():
                if note_id not in cited:  # separators, and notes nothing cites: Word shows neither
                    self.guard.skipped(note)
            for number, note_id in enumerate(cited, start=1):
                if (note := notes.get(note_id)) is not None:
                    note_context = replace(context, part=part, note_mark=NoteRegistry.mark(kind, number))
                    flowables += render_blocks(note, note_context)
        if not flowables:
            return []
        return [HRFlowable(width="33%", thickness=0.5, color=colors.black, hAlign="LEFT", spaceBefore=12,
                           spaceAfter=6), *flowables]

    def _template(self, index: int) -> PageTemplate:
        page = self.pages[index]
        frames = [Frame(page.column_left(column), page.bottom, page.column_width, page.frame_height,
                        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id=f"column-{column}")
                  for column in range(page.columns)]
        start, _ = _page_numbering(self.sections[index])

        def decorate(canvas: Canvas, _doc: BaseDocTemplate) -> None:
            self.section_pages[index] += 1
            first = self.section_pages[index] == 1
            self.page_number = start if first and start is not None else self.page_number + 1
            self._draw_header_footer(canvas, index, first)
            if page.column_separator:
                canvas.setLineWidth(0.5)
                for column in range(1, page.columns):
                    x = page.column_left(column) - page.column_space / 2
                    canvas.line(x, page.bottom, x, page.bottom + page.frame_height)

        return PageTemplate(id=f"section-{index}", frames=frames, pagesize=(page.width, page.height),
                            onPage=decorate)

    def _draw_header_footer(self, canvas: Canvas, index: int, first: bool) -> None:
        section, page = self.sections[index], self.pages[index]
        if first and section.different_first_page_header_footer:
            header, footer = section.first_page_header, section.first_page_footer
        elif self.odd_and_even and self.page_number % 2 == 0:  # Word goes by the displayed number
            header, footer = section.even_page_header, section.even_page_footer
        else:
            header, footer = section.header, section.footer
        values = {"PAGE": format_number(self.page_number, _page_numbering(section)[1])}
        if self.totals:
            values |= {"NUMPAGES": str(self.totals.pages), "SECTIONPAGES": str(self.totals.section_pages[index])}
        page_fields = PageFields(values, self.wanted_fields)
        for story, at_top in ((header, True), (footer, False)):
            part = story.part
            self.drawn_parts[part.partname] = part.element
            context = self.context(index, part, page_fields)
            flowables = render_blocks(part.element, context)
            sizes = [flowable.wrap(page.frame_width, page.height) for flowable in flowables]
            # A header hangs down from its distance; a footer stands up from its distance.
            y = page.height - page.header if at_top else page.footer + sum(height for _, height in sizes)
            for flowable, (_, height) in zip(flowables, sizes):
                y -= height
                flowable.drawOn(canvas, page.left, y)

    def _warnings(self, roots: Iterable[_Element]) -> tuple[str, ...]:
        warnings = []
        missing = "".join(self.guard.unhandled_text(root) for root in roots)
        if missing:
            sample = missing[:_SAMPLE] + ("..." if len(missing) > _SAMPLE else "")
            warnings.append(f"{len(missing)} character(s) of text were not rendered: {sample!r}")
        if unrenderable := self.guard.unrenderable:
            listed = ", ".join(f"{char!r} x{count}" for char, count in unrenderable.most_common(10))
            warnings.append(
                f"{unrenderable.total()} character(s) drawn as '?' (no bundled font has them, or their script "
                f"needs shaping): {listed}"
            )
        if not_drawn := self.guard.not_drawn:
            listed = ", ".join(f"{name} x{count}" for name, count in not_drawn.items())
            warnings.append(f"{not_drawn.total()} item(s) not drawn, shown as a placeholder: {listed}")
        if changes := self.guard.tracked_changes:
            warnings.append(f"{changes} tracked change(s) shown as accepted (insertions kept, deletions dropped)")
        return tuple(warnings)


def render_docx(data: bytes) -> RenderedDocument:
    document = docx.Document(io.BytesIO(data))
    first = _Renderer(document)
    rendered = first.render()
    if not first.wanted_fields:
        return rendered
    # NUMPAGES/SECTIONPAGES: totals are known only after one layout. Headers and
    # footers don't reflow the body, so a second pass with the totals is final.
    return _Renderer(document, _Totals(sum(first.section_pages), tuple(first.section_pages))).render()


def _page_numbering(section: Section) -> tuple[int | None, str]:
    """(restart number or None to continue, number format) from w:pgNumType."""
    numbering = child(section._sectPr, "w:pgNumType")
    start = attr(numbering, "w:start")
    return (int(start) if start else None), attr(numbering, "w:fmt") or "decimal"


def _page(section: Section) -> _Page:
    def points(length, default: float) -> float:
        return length.pt if length is not None else default

    width, height = A4  # Word's default in Australia, for a section with no page size
    columns = child(section._sectPr, "w:cols")
    return _Page(
        columns=max(int(attr(columns, "w:num") or 1), 1),
        column_space=twips_to_pt(attr(columns, "w:space")) or 36.0,
        column_separator=attr(columns, "w:sep") in ("1", "true", "on"),
        width=points(section.page_width, width),
        height=points(section.page_height, height),
        left=points(section.left_margin, _DEFAULT_MARGIN),
        right=points(section.right_margin, _DEFAULT_MARGIN),
        top=points(section.top_margin, _DEFAULT_MARGIN),
        bottom=points(section.bottom_margin, _DEFAULT_MARGIN),
        header=points(section.header_distance, _DEFAULT_DISTANCE),
        footer=points(section.footer_distance, _DEFAULT_DISTANCE),
    )


def _related_xml(document: Document, relationship: str) -> tuple[Part, _Element] | None:
    """The document's part of this relationship type and its XML root, or None if it
    has none. (python-docx's own numbering accessor would create an empty one, and it
    loads parts it doesn't model, like footnotes, as plain bytes.)"""
    try:
        part = document.part.part_related_by(relationship)
    except KeyError:
        return None
    return part, part.element if isinstance(part, XmlPart) else parse_xml(part.blob)
