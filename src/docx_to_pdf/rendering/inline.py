"""A paragraph's content (runs, and the wrappers runs sit in) as reportlab markup.

`python-docx`'s `paragraph.runs` only sees direct `w:r` children, which would drop
the text of hyperlinks, tracked insertions, content controls and fields. So this
walks the XML: `_CONTAINERS` are descended into, `_SKIPPED` are left out on purpose
(and marked as such for the text guard), and runs are rendered.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from xml.sax.saxutils import escape

from lxml.etree import _Element

from . import fonts, tabs
from .context import RenderContext
from .ooxml import EMU_PER_PT, attr, qn
from .styles import RunFormat

_CONTAINERS = {qn(tag) for tag in ("w:hyperlink", "w:smartTag", "w:customXml", "w:sdtContent", "w:fldSimple",
                                    "w:dir", "w:bdo")}
# Tracked changes are shown accepted, as the final version: insertions kept, deletions dropped.
_INSERTED = {qn("w:ins"), qn("w:moveTo")}
_DELETED = {qn("w:del"), qn("w:moveFrom")}
_NOTE_REFERENCES = {qn("w:footnoteReference"): "footnote", qn("w:endnoteReference"): "endnote"}
_NOTE_OWN_MARKS = {qn("w:footnoteRef"), qn("w:endnoteRef")}
# Graphics whose content lives in another part; drawn as a placeholder and reported.
_NOT_DRAWN = {"c:chart": "chart", "dgm:relIds": "SmartArt diagram"}


@dataclass(frozen=True)
class Picture:
    data: bytes | None  # None: the relationship couldn't be resolved
    width: float | None  # points; None: use the image's own size
    height: float | None
    floating: bool = False  # anchored (wp:anchor), not in the line of text


@dataclass(frozen=True)
class TextBox:
    content: _Element  # a w:txbxContent, rendered as block content


PAGE_BREAK = object()
COLUMN_BREAK = object()  # on a one-column page, a page break (as in Word)

Piece = str | Picture | TextBox | object  # markup, a picture, a text box, PAGE_BREAK or COLUMN_BREAK


@dataclass
class InlineContent:
    pieces: list[Piece] = field(default_factory=list)
    largest_font: float = 0.0

    def add_markup(self, markup: str) -> None:
        if self.pieces and isinstance(self.pieces[-1], str):
            self.pieces[-1] += markup
        else:
            self.pieces.append(markup)


def render_inline(paragraph: _Element, context: RenderContext) -> InlineContent:
    content = InlineContent()
    _walk(paragraph, paragraph, context, content)
    return content


def _walk(paragraph: _Element, parent: _Element, context: RenderContext, content: InlineContent) -> None:
    for element in parent:
        tag = element.tag
        if tag == qn("w:r"):
            _run(paragraph, element, context, content)
        elif (tag == qn("w:fldSimple") and context.page_fields is not None
              and (value := context.page_fields.value(attr(element, "w:instr"))) is not None):
            context.guard.skipped(element)  # the cached value, replaced by the real one
            content.add_markup(_styled(value, context.styles.run_format(paragraph, None)))
        elif tag in _CONTAINERS:
            _walk(paragraph, element, context, content)
        elif tag in _INSERTED:
            context.guard.tracked_changes += 1
            _walk(paragraph, element, context, content)
        elif tag in _DELETED:
            context.guard.tracked_changes += 1
            context.guard.skipped(element)
        elif tag == qn("w:sdt"):
            for sdt_content in element.iterfind(qn("w:sdtContent")):
                _walk(paragraph, sdt_content, context, content)
        elif tag == qn("m:oMathPara") or tag == qn("m:oMath"):
            # ponytail: equations as their plain text, not typeset
            content.add_markup(escape("".join(t.text or "" for t in element.iter(qn("m:t")))))


def _run(paragraph: _Element, run: _Element, context: RenderContext, content: InlineContent) -> None:
    run_format = context.styles.run_format(paragraph, run)
    if run_format.hidden:
        context.guard.skipped(run)
        return
    content.largest_font = max(content.largest_font, run_format.size)
    text: list[str] = []

    def flush() -> None:
        if text:
            content.add_markup(_styled("".join(text), run_format))
            text.clear()

    for element in run:
        tag = element.tag
        if tag == qn("w:fldChar"):
            if (value := _field_char(element, context)) is not None:  # a page field's result starts
                text.append(value)
        elif tag == qn("w:instrText") and context.fields.stack:
            context.fields.stack[-1]["instruction"] += element.text or ""
        elif context.page_fields is not None and context.fields.in_page_field_result(context.page_fields):
            context.guard.skipped(element)  # the cached value, replaced by the real one
        elif tag == qn("w:t"):
            context.guard.handled(element)
            text.append(_text(element.text or "", run_format, context))
        elif tag == qn("w:tab"):
            flush()  # outside the run's font tags, so the markup can be cut at it
            content.add_markup(tabs.TAB)
        elif tag == qn("w:ptab"):
            flush()
            content.add_markup(tabs.ptab(attr(element, "w:alignment"), attr(element, "w:relativeTo"),
                                         attr(element, "w:leader")))
        elif tag == qn("w:br") and attr(element, "w:type") in ("page", "column"):
            flush()
            content.pieces.append(PAGE_BREAK if attr(element, "w:type") == "page" else COLUMN_BREAK)
        elif tag in (qn("w:br"), qn("w:cr")):
            text.append("<br/>")
        elif tag == qn("w:noBreakHyphen"):
            text.append("-")
        elif tag in _NOTE_REFERENCES:
            flush()
            mark = context.notes.cite(_NOTE_REFERENCES[tag], attr(element, "w:id"))
            content.add_markup(_styled(mark, replace(run_format, vertical="superscript")))
        elif tag in _NOTE_OWN_MARKS and context.note_mark:
            flush()
            content.add_markup(_styled(context.note_mark, replace(run_format, vertical="superscript")))
        elif tag == qn("w:sym"):  # a Symbol/Wingdings glyph: its code means nothing in other fonts
            context.guard.unrenderable[f"symbol {attr(element, 'w:char')}"] += 1
            text.append(fonts.UNRENDERABLE)
        elif tag in (qn("w:drawing"), qn("w:pict"), qn("w:object"), qn("mc:AlternateContent")):
            flush()
            content.pieces.extend(_graphics(element, context))
    flush()


def _field_char(element: _Element, context: RenderContext) -> str | None:
    """Track the field stack. At the "separate" of a page-number field in a header or
    footer, return its real value, drawn there (the cached result after it is skipped,
    and may be absent: docx-js writes begin/separate/end with nothing between)."""
    kind = attr(element, "w:fldCharType")
    stack = context.fields.stack
    if kind == "begin":
        stack.append({"instruction": "", "in_result": False})
    elif kind == "separate" and stack:
        stack[-1]["in_result"] = True
        if context.page_fields is not None:
            return context.page_fields.value(stack[-1]["instruction"])
    elif kind == "end" and stack:
        stack.pop()
    return None


def _text(raw: str, run_format: RunFormat, context: RenderContext) -> str:
    if run_format.caps:
        raw = raw.upper()
    unrenderable: list[str] = []
    face = fonts.face_for(run_format.font, run_format.bold, run_format.italic)
    markup = []
    for segment in fonts.split_by_coverage(raw, face, unrenderable):
        # reportlab collapses runs of spaces, which Word keeps
        text = re.sub(r"(?<= ) ", "&nbsp;", escape(segment.text))
        markup.append(text if segment.face is None else f'<font face="{segment.face}">{text}</font>')
    context.guard.unrenderable.update(unrenderable)
    return "".join(markup)


def _styled(markup: str, run_format: RunFormat) -> str:
    face = fonts.face_for(run_format.font, run_format.bold, run_format.italic)
    # Word draws super/subscript at about 2/3 size; set here, as this font tag's size
    # would override the one <super>/<sub> sets.
    size = run_format.size * _SCRIPT_SIZE if run_format.vertical else run_format.size
    attributes = f'face="{face}" size="{size:.3g}"'
    if run_format.color:
        attributes += f' color="{run_format.color}"'
    if run_format.background:
        attributes += f' backColor="{run_format.background}"'
    markup = f"<font {attributes}>{markup}</font>"
    if run_format.underline:
        markup = f"<u>{markup}</u>"
    if run_format.strike:
        markup = f"<strike>{markup}</strike>"
    if run_format.vertical == "superscript":
        markup = f'<super rise="{run_format.size * 0.33:.3g}">{markup}</super>'
    elif run_format.vertical == "subscript":
        markup = f'<sub rise="{run_format.size * 0.14:.3g}">{markup}</sub>'
    return markup


_SCRIPT_SIZE = 0.65


def _graphics(element: _Element, context: RenderContext) -> list[Piece]:
    """Pictures and text boxes inside a drawing, VML picture or alternate content.
    ponytail: floating (anchored) shapes are placed in the flow at their anchor,
    not at their page position; shapes, charts and SmartArt draw nothing."""
    if element.tag == qn("mc:AlternateContent"):
        choice = element.find(qn("mc:Choice"))
        return [] if choice is None else [p for e in choice for p in _graphics(e, context)]
    pieces: list[Piece] = []
    for tag, name in _NOT_DRAWN.items():
        for _ in element.iter(qn(tag)):
            context.guard.not_drawn[name] += 1
            pieces.append(f"[{name} not shown]")
    # Only this level's own pictures and boxes: a text box renders what is inside it.
    for blip in _outside_boxes(element, "a:blip"):
        floating = any(ancestor.tag == qn("wp:anchor") for ancestor in blip.iterancestors())
        pieces.append(_picture(attr(blip, "r:embed"), _extent(blip), context, floating))
    pieces += [_picture(attr(image, "r:id"), None, context) for image in _outside_boxes(element, "v:imagedata")]
    pieces += [TextBox(box) for box in _outside_boxes(element, "w:txbxContent")]
    return pieces


def _outside_boxes(element: _Element, tag: str) -> list[_Element]:
    box = qn("w:txbxContent")

    def in_a_box(candidate: _Element) -> bool:
        for ancestor in candidate.iterancestors():
            if ancestor is element:  # boxes further out don't matter
                return False
            if ancestor.tag == box:
                return True
        return False

    return [candidate for candidate in element.iter(qn(tag)) if not in_a_box(candidate)]


def _extent(blip: _Element) -> tuple[float, float] | None:
    for ancestor in blip.iterancestors():
        if ancestor.tag in (qn("wp:inline"), qn("wp:anchor")):
            extent = ancestor.find(qn("wp:extent"))
            if extent is not None:
                return int(extent.get("cx")) / EMU_PER_PT, int(extent.get("cy")) / EMU_PER_PT
    return None


def _picture(relationship_id: str | None, extent: tuple[float, float] | None, context: RenderContext,
             floating: bool = False) -> Picture:
    part = context.part.related_parts.get(relationship_id) if relationship_id else None
    width, height = extent or (None, None)
    return Picture(data=getattr(part, "blob", None), width=width, height=height, floating=floating)
