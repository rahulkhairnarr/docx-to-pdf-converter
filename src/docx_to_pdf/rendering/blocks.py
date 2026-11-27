"""Block content (paragraphs, tables, content controls) as reportlab flowables.

`BLOCK_RENDERERS` maps a body element's tag to the function that renders it. To
support another construct, write `render_x(element, context) -> list[Flowable]` and
add it there; until then its text is reported by the text guard, never lost silently.
"""

from __future__ import annotations

import base64
import io
from collections.abc import Callable
from dataclasses import replace

from lxml.etree import _Element
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Flowable, FrameBreak, Image, PageBreak, Paragraph, Table, TableStyle

from . import fonts, tabs
from .borders import Bordered
from .context import RenderContext
from .inline import COLUMN_BREAK, PAGE_BREAK, Picture, TextBox, render_inline
from .paragraphs import make_paragraph
from .ooxml import attr, child, qn, twips_to_pt
from .styles import ParagraphFormat

BlockRenderer = Callable[[_Element, RenderContext], list[Flowable]]

_ALIGNMENTS = {"left": TA_LEFT, "center": TA_CENTER, "right": TA_RIGHT, "justify": TA_JUSTIFY}
# Formats Pillow reads the same on every platform. EMF/WMF need Windows' GDI to draw.
_PORTABLE_IMAGE_FORMATS = {"PNG", "JPEG", "GIF", "BMP", "TIFF", "WEBP"}


def render_blocks(parent: _Element, context: RenderContext) -> list[Flowable]:
    return [flowable for element in parent for flowable in render_block(element, context)]


def render_block(element: _Element, context: RenderContext) -> list[Flowable]:
    """Elements with no renderer (sectPr, bookmarks) draw nothing."""
    renderer = BLOCK_RENDERERS.get(element.tag)
    return renderer(element, context) if renderer else []


def render_paragraph(paragraph: _Element, context: RenderContext) -> list[Flowable]:
    reference = context.styles.numbering_reference(paragraph)
    level_ppr = context.numbering.level_properties(*reference)[0] if reference else None
    label = context.numbering.next_label(*reference) if reference else None
    paragraph_format = context.styles.paragraph_format(paragraph, level_ppr)
    mark = context.styles.run_format(paragraph, None)
    content = render_inline(paragraph, context)
    size = max(content.largest_font, mark.size)
    face = fonts.face_for(mark.font, mark.bold, mark.italic)
    if paragraph_format.line_spacing[0] == "auto" and content.largest_font > mark.size:
        # Mixed sizes: Word sizes each line by its own largest font, so lines of small
        # text stay close. Lead by the paragraph mark; taller lines grow to fit.
        style = _paragraph_style(paragraph_format, face, mark.size * fonts.line_height(mark.font), size, label)
        style.autoLeading = "max"
    else:
        style = _paragraph_style(paragraph_format, face, size * fonts.line_height(mark.font), size, label)

    flowables: list[Flowable] = [PageBreak()] if paragraph_format.page_break_before else []
    has_text = False
    for piece in _inline_pictures(content.pieces, context):
        if isinstance(piece, str):
            flowables.append(_text(piece, style, paragraph_format, None if has_text else label, context))
            has_text = True
        elif piece is PAGE_BREAK:
            flowables.append(PageBreak())
        elif piece is COLUMN_BREAK:
            flowables.append(FrameBreak())
        elif isinstance(piece, Picture):
            flowables.append(_picture(piece, context))
        elif isinstance(piece, TextBox):
            flowables.extend(render_blocks(piece.content, context))
    if not has_text and not any(isinstance(piece, (Picture, TextBox)) or piece in (PAGE_BREAK, COLUMN_BREAK)
                                for piece in content.pieces):
        # An empty paragraph still takes up a line in Word; a label on its own still shows.
        # One holding only a break keeps its mark on the line before the break, so it
        # adds no blank line at the top of the next page.
        flowables.append(Paragraph("&nbsp;", style, bulletText=label))
    if paragraph_format.borders or paragraph_format.shading:
        flowables = [Bordered(flowable, paragraph_format) if isinstance(flowable, (Paragraph, tabs.TabbedLine))
                     else flowable for flowable in flowables]
    return flowables


def _inline_pictures(pieces: list, context: RenderContext) -> list:
    """Inline pictures join the paragraph's text as <img> markup, as in Word: they sit
    on the baseline, line up side by side, and follow the paragraph's alignment (a
    logo beside header text, an icon in a sentence, a centred figure). Floating
    pictures, and ones that can't be drawn (placeholders), stay blocks."""
    merged: list = []
    for piece in pieces:
        if isinstance(piece, Picture) and not piece.floating and (markup := _picture_markup(piece, context)):
            piece = markup
        if isinstance(piece, str) and merged and isinstance(merged[-1], str):
            merged[-1] += piece
        else:
            merged.append(piece)
    return merged



def _text(markup: str, style: ParagraphStyle, paragraph_format: ParagraphFormat, label: str | None,
          context: RenderContext) -> Flowable:
    if tabs.has_tabs(markup):
        if label is None and paragraph_format.alignment in ("left", "justify"):
            line = tabs.tabbed_line(markup, style, paragraph_format.tabs, context.width, context.default_tab)
            if line is not None:
                return line
        markup = tabs.as_gaps(markup)
    return make_paragraph(markup, style, label)


def render_content_control(element: _Element, context: RenderContext) -> list[Flowable]:
    """A block-level content control (w:sdt) or custom XML wrapper: render what's inside."""
    inner = child(element, "w:sdtContent")
    return render_blocks(element if inner is None else inner, context)


def render_table(table: _Element, context: RenderContext) -> list[Flowable]:
    widths = _column_widths(table, context.width)
    if not widths:
        return []
    rows = table.findall(qn("w:tr"))
    data: list[list[object]] = [["" for _ in widths] for _ in rows]
    table_margins = _table_cell_margins(table, context)
    commands: list[tuple] = [("VALIGN", (0, 0), (-1, -1), "TOP")]
    merge_start: dict[int, int] = {}  # column -> row where its vertical merge began
    for row_index, row in enumerate(rows):
        column = int(attr(child(row, "w:trPr", "w:gridBefore")) or 0)
        for cell in row.iterfind(qn("w:tc")):
            span = int(attr(child(cell, "w:tcPr", "w:gridSpan")) or 1)
            last = min(column + span, len(widths)) - 1
            if column > last:
                context.guard.skipped(cell)  # more cells than grid columns: malformed, nothing to place it in
                continue
            merge = child(cell, "w:tcPr", "w:vMerge")
            if merge is not None and attr(merge) != "restart" and column in merge_start:
                context.guard.skipped(cell)  # continuation of the cell above; Word shows only that one
                commands.append(("SPAN", (column, merge_start[column]), (last, row_index)))
            else:
                if merge is not None:
                    merge_start[column] = row_index
                else:
                    merge_start.pop(column, None)
                margins = table_margins | _margins(child(cell, "w:tcPr", "w:tcMar"))
                cell_width = sum(widths[column:last + 1]) - margins["left"] - margins["right"]
                content = render_blocks(cell, replace(context, width=cell_width))
                data[row_index][column] = content or ""
                if content:  # reportlab drops these two spacings in a cell; Word keeps them
                    margins["top"] += content[0].getSpaceBefore()
                    margins["bottom"] += content[-1].getSpaceAfter()
                commands += [(f"{side.upper()}PADDING", (column, row_index), (last, row_index), value)
                             for side, value in margins.items()]
                if last > column:
                    commands.append(("SPAN", (column, row_index), (last, row_index)))
                commands.extend(_cell_style(cell, (column, row_index), (last, row_index)))
            column = last + 1
    if _has_borders(table, context):
        commands.append(("GRID", (0, 0), (-1, -1), 0.5, colors.black))
    header_rows = 0
    while header_rows < len(rows) and child(rows[header_rows], "w:trPr", "w:tblHeader") is not None:
        header_rows += 1
    # A vertical merge spans rows: reportlab wants the SPAN once, over the whole run.
    commands = _merge_vertical_spans(commands)
    return [Table(data, colWidths=widths, style=TableStyle(commands), repeatRows=header_rows,
                  splitInRow=1, hAlign=_TABLE_ALIGNMENTS.get(attr(child(table, "w:tblPr", "w:jc")), "LEFT"))]


_TABLE_ALIGNMENTS = {"center": "CENTER", "right": "RIGHT", "end": "RIGHT"}


def render_embedded_document(_element: _Element, context: RenderContext) -> list[Flowable]:
    """w:altChunk: a whole HTML/RTF/docx file embedded as its own part, which Word
    converts when it opens the document. ponytail: shown as a placeholder; render
    an embedded .docx by recursing into render_docx if real documents use them."""
    return [_placeholder("embedded document (altChunk)", context)]


BLOCK_RENDERERS: dict[str, BlockRenderer] = {
    qn("w:p"): render_paragraph,
    qn("w:tbl"): render_table,
    qn("w:sdt"): render_content_control,
    qn("w:customXml"): render_content_control,
    qn("w:altChunk"): render_embedded_document,
}


def _placeholder(name: str, context: RenderContext) -> Flowable:
    """A visible marker for content that can't be drawn, counted for the warnings."""
    context.guard.not_drawn[name] += 1
    style = ParagraphStyle("placeholder", fontName=fonts.face_for(None, False, False), fontSize=8)
    return Paragraph(f"[{name} not shown]", style)


def _paragraph_style(
    paragraph_format: ParagraphFormat, face: str, single: float, size: float, label: str | None
) -> ParagraphStyle:
    """`single` is Word's single line height for this paragraph's font and size."""
    rule, value = paragraph_format.line_spacing
    leading = single * value if rule == "auto" else value if rule == "exact" else max(value, single)
    left, first_line = paragraph_format.left, paragraph_format.first_line
    style = ParagraphStyle(
        "word",
        fontName=face,
        fontSize=size,
        leading=leading,
        alignment=_ALIGNMENTS[paragraph_format.alignment],
        leftIndent=left,
        rightIndent=paragraph_format.right,
        # A list label hangs at the first-line position; the text starts at the left indent.
        firstLineIndent=0 if label else first_line,
        bulletIndent=max(left + first_line, 0),
        bulletFontName=fonts.face_covering(label, face) if label else face,
        bulletFontSize=size,
        spaceBefore=paragraph_format.space_before,
        spaceAfter=paragraph_format.space_after,
        keepWithNext=int(paragraph_format.keep_with_next),
        # Word's widow/orphan control: no single line of a paragraph alone on a page.
        allowWidows=int(not paragraph_format.widow_control),
        allowOrphans=int(not paragraph_format.widow_control),
    )
    style.lineMultiple = value if rule == "auto" else 1.0  # for lines sized one by one (paragraphs.py)
    return style


def _picture(picture: Picture, context: RenderContext) -> Flowable:
    if (size := _picture_size(picture, context)) is None:
        return _placeholder(f"image ({_image_format(picture.data) or 'unreadable'})", context)
    return Image(io.BytesIO(picture.data), width=size[0], height=size[1], hAlign="LEFT")


def _picture_markup(picture: Picture, context: RenderContext) -> str | None:
    """An <img> tag carrying the picture as a data: URI; None if it can't be drawn
    (the caller keeps it as a block, which draws the placeholder)."""
    if (size := _picture_size(picture, context)) is None:
        return None
    data = base64.b64encode(picture.data).decode("ascii")
    return f'<img src="data:image/{_image_format(picture.data).lower()};base64,{data}" width="{size[0]:.2f}" ' \
           f'height="{size[1]:.2f}"/>'


def _picture_size(picture: Picture, context: RenderContext) -> tuple[float, float] | None:
    """Points to draw the picture at, scaled down to fit; None for a format that isn't portable."""
    if _image_format(picture.data) not in _PORTABLE_IMAGE_FORMATS:
        return None
    with PILImage.open(io.BytesIO(picture.data)) as image:
        natural = (image.width * 0.75, image.height * 0.75)  # pixels at 96 dpi -> points
    width, height = (picture.width, picture.height) if picture.width and picture.height else natural
    scale = min(1.0, context.width / width, context.height / height)
    return width * scale, height * scale


def _image_format(data: bytes | None) -> str | None:
    if not data:
        return None
    try:
        with PILImage.open(io.BytesIO(data)) as image:
            return image.format
    except Exception:  # noqa: BLE001 - any undecodable image is reported as such, not raised
        return None


def _column_widths(table: _Element, available: float) -> list[float]:
    widths = [twips_to_pt(attr(column, "w:w")) or 0.0 for column in table.iterfind(f"{qn('w:tblGrid')}/{qn('w:gridCol')}")]
    if not widths:
        cells = max((len(row.findall(qn("w:tc"))) for row in table.iterfind(qn("w:tr"))), default=0)
        widths = [0.0] * cells
    if widths and not all(widths):
        widths = [available / len(widths)] * len(widths)
    total = sum(widths)
    return [width * available / total for width in widths] if total > available else widths


def _cell_style(cell: _Element, start: tuple[int, int], end: tuple[int, int]) -> list[tuple]:
    commands: list[tuple] = []
    fill = attr(child(cell, "w:tcPr", "w:shd"), "w:fill")
    if fill and fill != "auto" and len(fill) == 6:
        commands.append(("BACKGROUND", start, end, colors.HexColor(f"#{fill}")))
    vertical = attr(child(cell, "w:tcPr", "w:vAlign"))
    if vertical in ("center", "bottom"):
        commands.append(("VALIGN", start, end, "MIDDLE" if vertical == "center" else "BOTTOM"))
    return commands


_DEFAULT_CELL_MARGINS ={"top": 0.0, "bottom": 0.0, "left": 5.4, "right": 5.4}  # Word's: 0.08" each side


def _table_cell_margins(table: _Element, context: RenderContext) -> dict[str, float]:
    """Cell margins from the table, else its style chain, else Word's defaults. A
    cell's own w:tcMar is laid over these, side by side."""
    style_id = attr(child(table, "w:tblPr", "w:tblStyle"))
    margins = dict(_DEFAULT_CELL_MARGINS)
    chain = [child(style, "w:tblPr", "w:tblCellMar") for style in context.styles.style_chain(style_id)]
    for found in reversed([child(table, "w:tblPr", "w:tblCellMar"), *chain]):
        margins |= _margins(found)
    return margins


def _margins(element: _Element | None) -> dict[str, float]:
    """The sides a w:tblCellMar / w:tcMar sets (dxa widths), with start/end as left/right."""
    sides = {"top": "top", "bottom": "bottom", "left": "left", "start": "left", "right": "right", "end": "right"}
    found: dict[str, float] = {}
    for edge in element if element is not None else ():
        side = sides.get(edge.tag.split("}")[-1])
        if side and attr(edge, "w:type") in (None, "dxa"):
            found[side] = twips_to_pt(attr(edge, "w:w")) or 0.0
    return found


def _has_borders(table: _Element, context: RenderContext) -> bool:
    """ponytail: borders are all or nothing, from the table's or its style's tblBorders;
    per-cell borders and border widths/colours are not reproduced."""
    style_id = attr(child(table, "w:tblPr", "w:tblStyle"))
    candidates = [child(table, "w:tblPr", "w:tblBorders")]
    candidates += [child(style, "w:tblPr", "w:tblBorders") for style in context.styles.style_chain(style_id)]
    for borders in candidates:
        if borders is not None:
            return any(attr(edge) not in ("nil", "none") for edge in borders)
    return False


def _merge_vertical_spans(commands: list[tuple]) -> list[tuple]:
    """Collapse the per-row SPANs of a vertical merge into one SPAN over all its rows."""
    spans: dict[tuple[int, int], tuple] = {}
    others = []
    for command in commands:
        if command[0] != "SPAN":
            others.append(command)
            continue
        _, (c0, r0), (c1, r1) = command
        key = (c0, r0)
        if key in spans:
            _, start, (old_c1, old_r1) = spans[key]
            spans[key] = ("SPAN", start, (max(old_c1, c1), max(old_r1, r1)))
        else:
            spans[key] = command
    return others + list(spans.values())
