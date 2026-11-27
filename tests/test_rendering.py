"""Real conversions, read back with pypdfium2."""

from __future__ import annotations

import io
import re
from pathlib import Path

import docx
import pypdfium2
import pytest
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_LINE_SPACING
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.shared import Inches, Mm, Pt
from PIL import Image as PILImage
from conftest import (
    DELETED_WORDS,
    RICH_WORDS,
    W,
    add_note,
    pdf_page_sizes,
    pdf_pages_text,
    pdf_text_x,
    write_docx,
    write_rich_docx,
)

from docx_to_pdf.errors import ConversionFailed
from docx_to_pdf.rendering import blocks, convert_file
from docx_to_pdf.rendering.numbering import Numbering, format_number
from docx_to_pdf.rendering.ooxml import qn


@pytest.fixture
def rich_pdf(tmp_path: Path) -> tuple[Path, tuple[str, ...]]:
    target = tmp_path / "rich.pdf"
    rendered = convert_file(write_rich_docx(tmp_path / "rich.docx"), target)
    return target, rendered.warnings


def test_every_word_reaches_the_pdf(rich_pdf: tuple[Path, tuple[str, ...]]) -> None:
    target, warnings = rich_pdf
    text = " ".join(pdf_pages_text(target))
    # The fixture's w:ins, w:del and w:moveFrom; nothing else may be reported.
    assert warnings == ("3 tracked change(s) shown as accepted (insertions kept, deletions dropped)",)
    assert [word for word in RICH_WORDS if word not in text] == []
    assert [word for word in DELETED_WORDS if word in text] == []


def test_text_box_is_drawn_once(rich_pdf: tuple[Path, tuple[str, ...]]) -> None:
    text = " ".join(pdf_pages_text(rich_pdf[0]))
    assert text.count("Boxed") == 1  # mc:Choice only, not its mc:Fallback copy


def test_merged_cells_are_not_duplicated(rich_pdf: tuple[Path, tuple[str, ...]]) -> None:
    text = " ".join(pdf_pages_text(rich_pdf[0]))
    assert text.count("Mergedwide") == 1 and text.count("Mergedtall") == 1


def test_list_labels_are_generated(rich_pdf: tuple[Path, tuple[str, ...]]) -> None:
    text = pdf_pages_text(rich_pdf[0])[0]
    assert re.search(r"1\.\s*Numbered first", text) and re.search(r"2\.\s*Numbered second", text)
    assert re.search(r"•\s*Bullet first", text)


def test_footer_shows_the_real_page_number(rich_pdf: tuple[Path, tuple[str, ...]]) -> None:
    pages = pdf_pages_text(rich_pdf[0])
    assert len(pages) == 3
    for number, page in enumerate(pages, start=1):
        assert f"Page {number}" in page
    assert "Headertext" in pages[0]


def test_page_field_in_one_run_with_no_cached_result(tmp_path: Path) -> None:
    """docx-js's shape: the whole field in one run, nothing between separate and end."""
    document = docx.Document()
    document.add_paragraph("One")
    document.add_page_break()
    document.add_paragraph("Two")
    document.sections[0].footer.paragraphs[0]._p.append(parse_xml(
        f'<w:r {W}><w:t xml:space="preserve">Page </w:t><w:fldChar w:fldCharType="begin"/>'
        '<w:instrText>PAGE</w:instrText><w:fldChar w:fldCharType="separate"/>'
        '<w:fldChar w:fldCharType="end"/><w:t xml:space="preserve"> end</w:t></w:r>'
    ))
    document.save(tmp_path / "field.docx")
    convert_file(tmp_path / "field.docx", tmp_path / "field.pdf")
    for number, page in enumerate(pdf_pages_text(tmp_path / "field.pdf"), start=1):
        assert re.search(rf"Page {number}\s*end", page)


def test_page_numbers_restart_format_totals_and_pick_odd_even_headers(tmp_path: Path) -> None:
    """Section 2 restarts at i (lower Roman). Word picks the even header by the
    displayed number, and NUMPAGES/SECTIONPAGES need the totals (a second pass)."""
    document = docx.Document()
    document.settings.odd_and_even_pages_header_footer = True
    document.add_paragraph("Cover")
    second = document.add_section(WD_SECTION_START.NEW_PAGE)
    second._sectPr.append(parse_xml(f'<w:pgNumType {W} w:start="1" w:fmt="lowerRoman"/>'))
    document.add_paragraph("Body one")
    document.add_page_break()
    document.add_paragraph("Body two")
    for section in document.sections:
        section.header.paragraphs[0].text = "Oddhead"
        section.even_page_header.paragraphs[0].text = "Evenhead"
        for footer in (section.footer, section.even_page_footer):
            for label, instruction in (("P ", " PAGE "), (" of ", "NUMPAGES"), (" sec ", "SECTIONPAGES")):
                footer.paragraphs[0]._p.append(parse_xml(
                    f'<w:r {W}><w:t xml:space="preserve">{label}</w:t></w:r>'))
                footer.paragraphs[0]._p.append(parse_xml(
                    f'<w:fldSimple {W} w:instr="{instruction}"><w:r><w:t>9</w:t></w:r></w:fldSimple>'))
    document.save(tmp_path / "numbers.docx")
    convert_file(tmp_path / "numbers.docx", tmp_path / "numbers.pdf")
    cover, one, two = pdf_pages_text(tmp_path / "numbers.pdf")
    assert re.search(r"P 1\s*of 3\s*sec 1", cover) and "Oddhead" in cover
    assert re.search(r"P i\s*of 3\s*sec 2", one) and "Oddhead" in one  # i is odd
    assert re.search(r"P ii\s*of 3\s*sec 2", two) and "Evenhead" in two


def test_tab_stops_right_leader_decimal_and_centre(tmp_path: Path) -> None:
    document = docx.Document()
    margin = document.sections[0].left_margin.pt
    right_edge = margin + 432.0  # the stop at 8640 twips
    lines = (
        ('<w:tab w:val="right" w:pos="8640" w:leader="dot"/>', "Intro", "Rightword"),
        ('<w:tab w:val="decimal" w:pos="5000"/>', "Price", "12.5"),
        ('<w:tab w:val="decimal" w:pos="5000"/>', "Tax", "1,234.567"),
        ('<w:tab w:val="center" w:pos="4320"/>', "Start", "Centred"),
    )
    for stop, before, after in lines:
        document.element.body.sectPr.addprevious(parse_xml(
            f'<w:p {W}><w:pPr><w:tabs>{stop}</w:tabs></w:pPr>'
            f'<w:r><w:t>{before}</w:t><w:tab/><w:t>{after}</w:t></w:r></w:p>'))
    document.save(tmp_path / "tabs.docx")
    convert_file(tmp_path / "tabs.docx", tmp_path / "tabs.pdf")
    target = tmp_path / "tabs.pdf"
    assert abs(pdf_text_x(target, "Rightword")[1] - right_edge) < 1.5
    assert "......" in pdf_pages_text(target)[0]
    assert abs(pdf_text_x(target, "2.5")[0] - pdf_text_x(target, "4.5")[0]) < 1.0  # the points align
    assert abs(pdf_text_x(target, ".567")[0] - (margin + 250)) < 1.5
    left, right = pdf_text_x(target, "Centred")
    assert abs((left + right) / 2 - (margin + 216)) < 1.5


def test_inline_picture_shares_a_tabbed_line_with_text(tmp_path: Path) -> None:
    """The usual header: a logo, a tab to a right stop, then text, all on one line."""
    image = io.BytesIO()
    PILImage.new("RGB", (40, 40), "navy").save(image, "PNG")
    document = docx.Document()
    margin = document.sections[0].left_margin.pt
    paragraph = document.add_paragraph()
    paragraph._p.get_or_add_pPr().append(parse_xml(f'<w:tabs {W}><w:tab w:val="right" w:pos="8640"/></w:tabs>'))
    paragraph.add_run().add_picture(image, width=Pt(30))
    paragraph.add_run().add_tab()
    paragraph.add_run("Headerright")
    document.save(tmp_path / "logo.docx")
    convert_file(tmp_path / "logo.docx", tmp_path / "logo.pdf")
    page = pypdfium2.PdfDocument(str(tmp_path / "logo.pdf"))[0]
    images = list(page.get_objects(filter=[pypdfium2.raw.FPDF_PAGEOBJ_IMAGE]))
    assert len(images) == 1
    image_left, image_bottom, _, _ = images[0].get_pos()
    assert abs(image_left - margin) < 1
    assert abs(pdf_text_x(tmp_path / "logo.pdf", "Headerright")[1] - (margin + 432)) < 1.5
    text_bottom = page.get_textpage().get_charbox(0)[1]  # same line: the logo stands on the text's baseline
    assert abs(image_bottom - text_bottom) < 4


@pytest.mark.parametrize("beside", ["", "Caption"])
def test_inline_picture_stays_between_its_neighbours(tmp_path: Path, beside: str) -> None:
    """A tall inline picture, alone or with text, must not overlap the lines around it."""
    image = io.BytesIO()
    PILImage.new("RGB", (100, 100), "navy").save(image, "PNG")
    document = docx.Document()
    document.add_paragraph("Abovetext")
    paragraph = document.add_paragraph()
    paragraph.add_run().add_picture(image, width=Pt(100))
    paragraph.add_run(beside)
    document.add_paragraph("Belowtext")
    document.save(tmp_path / "tall.docx")
    convert_file(tmp_path / "tall.docx", tmp_path / "tall.pdf")
    page = pypdfium2.PdfDocument(str(tmp_path / "tall.pdf"))[0]
    _, bottom, _, top = next(iter(page.get_objects(filter=[pypdfium2.raw.FPDF_PAGEOBJ_IMAGE]))).get_pos()
    textpage = page.get_textpage()
    above = textpage.get_charbox(textpage.search("Abovetext").get_next()[0])
    below = textpage.get_charbox(textpage.search("Belowtext").get_next()[0])
    assert above[1] > top and bottom > below[3]


def test_two_columns_with_a_column_break(tmp_path: Path) -> None:
    document = docx.Document()
    section = document.sections[0]
    columns = section._sectPr.find(qn("w:cols"))  # the template has one: one column, 720 apart
    columns.set(qn("w:num"), "2")
    columns.set(qn("w:sep"), "1")
    left_margin, width = section.left_margin.pt, section.page_width.pt - section.left_margin.pt - section.right_margin.pt
    document.add_paragraph("Firstcolumn " * 30)
    document.element.body.sectPr.addprevious(parse_xml(
        f'<w:p {W}><w:r><w:br w:type="column"/><w:t>Secondcolumn</w:t></w:r></w:p>'))
    document.save(tmp_path / "columns.docx")
    convert_file(tmp_path / "columns.docx", tmp_path / "columns.pdf")
    column_width = (width - 36) / 2
    _, first_right = pdf_text_x(tmp_path / "columns.pdf", "Firstcolumn")
    second_left, _ = pdf_text_x(tmp_path / "columns.pdf", "Secondcolumn")
    assert len(pdf_pages_text(tmp_path / "columns.pdf")) == 1
    assert first_right < left_margin + column_width + 1  # the text wraps within its column
    assert abs(second_left - (left_margin + column_width + 36)) < 1.5
    page = pypdfium2.PdfDocument(str(tmp_path / "columns.pdf"))[0]
    assert len(list(page.get_objects(filter=[pypdfium2.raw.FPDF_PAGEOBJ_PATH]))) == 1  # the separator


def test_heading_stays_with_the_first_row_not_the_whole_table(tmp_path: Path) -> None:
    """Keep-with-next needs only the next item's first row on the page (reportlab's
    KeepTogether would push heading and whole table to a new page)."""
    document = docx.Document()
    for _ in range(22):  # most of page 1
        document.add_paragraph("Filler line")
    document.add_heading("Tableheading", level=1)  # Heading 1 keeps with next
    table = document.add_table(rows=40, cols=2)
    for index, row in enumerate(table.rows):
        row.cells[0].text = f"Row{index}"
    document.save(tmp_path / "keep.docx")
    convert_file(tmp_path / "keep.docx", tmp_path / "keep.pdf")
    first = pdf_pages_text(tmp_path / "keep.pdf")[0]
    assert "Tableheading" in first and "Row0" in first


def test_a_break_only_paragraph_adds_no_line_on_the_next_page(tmp_path: Path) -> None:
    document = docx.Document()
    document.add_paragraph("One")
    document.element.body.sectPr.addprevious(parse_xml(f'<w:p {W}><w:r><w:br w:type="page"/></w:r></w:p>'))
    document.add_paragraph("Two")
    document.save(tmp_path / "break.docx")
    convert_file(tmp_path / "break.docx", tmp_path / "break.pdf")
    section = document.sections[0]
    page = pypdfium2.PdfDocument(str(tmp_path / "break.pdf"))[1]
    top = page.get_textpage().get_charbox(0)[3]
    assert top > section.page_height.pt - section.top_margin.pt - 12  # on the first line, not the second


def test_mixed_sizes_size_each_line_by_its_own_largest_font(tmp_path: Path) -> None:
    document = docx.Document()
    paragraph = document.add_paragraph()
    for text, size, last in (("Tiny", 8, False), ("Huge", 40, False), ("Last", 8, True)):
        run = paragraph.add_run(text)
        run.font.size = Pt(size)
        if not last:
            run.add_break()
    document.save(tmp_path / "sizes.docx")
    convert_file(tmp_path / "sizes.docx", tmp_path / "sizes.pdf")
    page = pypdfium2.PdfDocument(str(tmp_path / "sizes.pdf"))[0]
    textpage = page.get_textpage()
    tiny, last = (textpage.get_charbox(textpage.search(word).get_next()[0])[3] for word in ("Tiny", "Last"))
    # Baseline 1 to 3: an 8pt line's descent, the whole 40pt line, an 8pt line's ascent,
    # at Calibri's 1.2207 and the template's 1.15 spacing; not three lines led like the 40pt one.
    assert (tiny - last) < 2 * 40 * 1.2207
    assert abs((tiny - last) - (40 + 8) * 1.2207 * 1.15) < 4


def test_cell_margins_and_paragraph_spacing_make_the_row_height(tmp_path: Path) -> None:
    document = docx.Document()
    table = document.add_table(rows=3, cols=1)
    for index, row in enumerate(table.rows):
        row.cells[0].text = f"Cell{index}"
        row.cells[0].paragraphs[0].paragraph_format.space_after = Pt(6)
        row.cells[0]._tc.get_or_add_tcPr().append(parse_xml(
            f'<w:tcMar {W}><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/></w:tcMar>'))
    document.save(tmp_path / "cells.docx")
    convert_file(tmp_path / "cells.docx", tmp_path / "cells.pdf")
    textpage = pypdfium2.PdfDocument(str(tmp_path / "cells.pdf"))[0].get_textpage()
    first, second = (textpage.get_charbox(textpage.search(word).get_next()[0])[3] for word in ("Cell0", "Cell1"))
    line = 11 * 1.2207 * 1.15  # the template's Normal: 11pt Calibri at 1.15 lines
    assert abs((first - second) - (4 + line + 6 + 4)) < 1.5


@pytest.mark.parametrize(("paragraph_properties", "paths"), [
    ("", 0),
    ('<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="1F4E79"/></w:pBdr>', 1),
    ('<w:pBdr><w:top w:val="double" w:sz="12" w:space="1"/></w:pBdr><w:shd w:val="clear" w:fill="FFF2CC"/>', 3),
])
def test_paragraph_borders_and_shading_are_drawn(tmp_path: Path, paragraph_properties: str, paths: int) -> None:
    document = docx.Document()
    document.element.body.sectPr.addprevious(parse_xml(
        f'<w:p {W}><w:pPr>{paragraph_properties}</w:pPr><w:r><w:t>Ruled</w:t></w:r></w:p>'))
    document.save(tmp_path / "rule.docx")
    convert_file(tmp_path / "rule.docx", tmp_path / "rule.pdf")
    page = pypdfium2.PdfDocument(str(tmp_path / "rule.pdf"))[0]
    assert len(list(page.get_objects(filter=[pypdfium2.raw.FPDF_PAGEOBJ_PATH]))) == paths
    assert "Ruled" in pdf_pages_text(tmp_path / "rule.pdf")[0]


@pytest.mark.parametrize(("level_text", "label"), [
    ("●", "●"), ("✓", "✓"), ("➢", "➢"), ("o", "o"),
    ("", "•"),  # Symbol font's bullet, a private-use code
    ("§", "•"),  # a Wingdings square, not the section sign
])
def test_bullet_labels(level_text: str, label: str) -> None:
    numbering = Numbering(parse_xml(
        f'<w:numbering {W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
        f'<w:numFmt w:val="bullet"/><w:lvlText w:val="{level_text}"/></w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    ))
    assert numbering.next_label("1", 0) == label


def test_footnotes_and_endnotes_are_rendered(tmp_path: Path) -> None:
    document = docx.Document()
    add_note(document, "footnote", "Footnotetext one")
    add_note(document, "footnote", "Footnotetext two")
    add_note(document, "endnote", "Endnotetext")
    document.save(tmp_path / "notes.docx")
    rendered = convert_file(tmp_path / "notes.docx", tmp_path / "notes.pdf")
    text = " ".join(pdf_pages_text(tmp_path / "notes.pdf"))
    assert rendered.warnings == ()
    assert re.search(r"1\s*Footnotetext one", text) and re.search(r"2\s*Footnotetext two", text)
    assert re.search(r"i\s*Endnotetext", text)
    assert re.search(r"Cited claim\s*1", text)


def test_charts_and_embedded_documents_are_placeholders_with_a_warning(tmp_path: Path) -> None:
    document = docx.Document()
    body = document.element.body
    body.sectPr.addprevious(parse_xml(
        f'<w:p {W} xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:r><w:drawing><wp:inline><a:graphic><a:graphicData><c:chart r:id="rId99"/>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    ))
    body.sectPr.addprevious(parse_xml(
        f'<w:altChunk {W} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'r:id="rId98"/>'
    ))
    document.save(tmp_path / "chart.docx")
    rendered = convert_file(tmp_path / "chart.docx", tmp_path / "chart.pdf")
    text = pdf_pages_text(tmp_path / "chart.pdf")[0]
    assert "[chart not shown]" in text and "[embedded document (altChunk) not shown]" in text
    assert rendered.warnings == (
        "2 item(s) not drawn, shown as a placeholder: chart x1, embedded document (altChunk) x1",
    )


def test_first_page_header_differs(tmp_path: Path) -> None:
    document = docx.Document()
    section = document.sections[0]
    section.different_first_page_header_footer = True
    section.first_page_header.paragraphs[0].text = "Coverheader"
    section.header.paragraphs[0].text = "Runningheader"
    document.add_paragraph("One")
    document.add_page_break()
    document.add_paragraph("Two")
    document.save(tmp_path / "first.docx")
    convert_file(tmp_path / "first.docx", tmp_path / "first.pdf")
    first, second = pdf_pages_text(tmp_path / "first.pdf")
    assert "Coverheader" in first and "Runningheader" not in first
    assert "Runningheader" in second and "Coverheader" not in second


def test_page_break_and_landscape_section(rich_pdf: tuple[Path, tuple[str, ...]]) -> None:
    pages = pdf_pages_text(rich_pdf[0])
    assert "Before break" in pages[0] and "After break" in pages[1]
    (portrait_w, portrait_h), _, (landscape_w, landscape_h) = pdf_page_sizes(rich_pdf[0])
    assert portrait_h > portrait_w and landscape_w > landscape_h


def test_unicode_beyond_western_european_is_drawn(tmp_path: Path) -> None:
    text = "Zoë – ‘quoted’ — Nguyễn Ελένη Дмитрий ☐ ☒ ✓ €"
    source = write_docx(tmp_path / "names.docx", text)
    rendered = convert_file(source, tmp_path / "names.pdf")
    assert rendered.warnings == ()
    assert text in pdf_pages_text(tmp_path / "names.pdf")[0]


def test_invisible_characters_are_not_drawn_or_reported(tmp_path: Path) -> None:
    source = write_docx(tmp_path / "pasted.docx", "Done ✓️​ ok﻿ ‎end‏")
    rendered = convert_file(source, tmp_path / "pasted.pdf")
    text = pdf_pages_text(tmp_path / "pasted.pdf")[0]
    assert rendered.warnings == ()
    assert "?" not in text and "Done ✓ ok end" in text


def test_decomposed_text_is_composed_before_drawing(tmp_path: Path) -> None:
    source = write_docx(tmp_path / "nfd.docx", "Nguyễn")  # NFD of "Nguyễn"
    rendered = convert_file(source, tmp_path / "nfd.pdf")
    assert rendered.warnings == ()
    assert "Nguyễn" in pdf_pages_text(tmp_path / "nfd.pdf")[0]


def test_unsupported_and_shaped_scripts_are_reported_not_drawn(tmp_path: Path) -> None:
    source = write_docx(tmp_path / "scripts.docx", "Wang 王 / سلام / नमस्ते")
    rendered = convert_file(source, tmp_path / "scripts.pdf")
    [warning] = rendered.warnings
    assert "'王' x1" in warning and "'س' x1" in warning and "'न' x" in warning
    assert "Wang ? / ???? /" in pdf_pages_text(tmp_path / "scripts.pdf")[0]


def test_word_fonts_use_their_metric_compatible_substitutes(tmp_path: Path) -> None:
    document = docx.Document()
    for font in ("Calibri", "Arial", "Times New Roman", "Courier New", "Cambria"):
        document.add_paragraph(font).runs[0].font.name = font
    document.save(tmp_path / "fonts.docx")
    convert_file(tmp_path / "fonts.docx", tmp_path / "fonts.pdf")
    pdf = (tmp_path / "fonts.pdf").read_bytes()
    for embedded in (b"Carlito", b"LiberationSans", b"LiberationSerif", b"LiberationMono", b"Caladea"):
        assert embedded in pdf
    assert b"Helvetica" not in pdf and b"Times-Roman" not in pdf  # never the PDF standard fonts


def test_theme_fonts_are_resolved(tmp_path: Path) -> None:
    """python-docx's template, like Word's, names its body font only through the theme
    (docDefaults: asciiTheme="minorHAnsi"; theme minor font: Cambria). Swapping the
    theme's minor font must change the PDF's font."""
    source = write_docx(tmp_path / "theme.docx")
    package = docx.Document(source)
    theme = package.part.part_related_by(RT.THEME)
    swapped, count = re.subn(rb'(<a:minorFont>\s*<a:latin typeface=")Cambria"', rb'\1Courier New"', theme.blob)
    assert count == 1  # the swap really happened, so the assertion below means something
    theme._blob = swapped
    package.save(tmp_path / "theme.docx")
    convert_file(tmp_path / "theme.docx", tmp_path / "theme.pdf")
    pdf = (tmp_path / "theme.pdf").read_bytes()
    assert b"LiberationMono" in pdf and b"Caladea" not in pdf


def test_a_run_naming_only_its_east_asian_font_keeps_the_latin_font(tmp_path: Path) -> None:
    document = docx.Document()
    document.styles["Normal"].font.name = "Courier New"
    run = document.add_paragraph().add_run("Monospaced")
    run._r.get_or_add_rPr().append(parse_xml(f'<w:rFonts {W} w:eastAsia="MS Mincho"/>'))
    document.save(tmp_path / "east.docx")
    convert_file(tmp_path / "east.docx", tmp_path / "east.pdf")
    assert b"LiberationMono" in (tmp_path / "east.pdf").read_bytes()


@pytest.mark.parametrize("widow_control", [True, False, None])
def test_widow_control_keeps_a_first_line_off_the_page_bottom(tmp_path: Path, widow_control: bool | None) -> None:
    """Exact 12pt lines and no paragraph spacing: page 1 (A4, 1" margins, 697.9pt) holds
    58 lines. After 57 filler lines, a 3-line paragraph has room for its first line only.
    None leaves w:widowControl out entirely, which Word treats as on."""
    document = docx.Document()
    section = document.sections[0]
    section.page_height, section.page_width = Mm(297), Mm(210)
    section.top_margin = section.bottom_margin = Inches(1)
    normal = document.styles["Normal"].paragraph_format
    normal.space_before = normal.space_after = Pt(0)
    normal.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    normal.line_spacing = Pt(12)
    for line in range(57):
        document.add_paragraph(f"Filler{line}")
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.widow_control = widow_control  # None: no element written
    assert (paragraph._p.find(qn("w:pPr")) is None or paragraph._p.pPr.find(qn("w:widowControl")) is None) is (
        widow_control is None
    )
    for index, text in enumerate(("FirstLine", "SecondLine", "ThirdLine")):
        run = paragraph.add_run(text)
        if index < 2:
            run.add_break()
    document.save(tmp_path / "widow.docx")
    convert_file(tmp_path / "widow.docx", tmp_path / "widow.pdf")
    first, second = pdf_pages_text(tmp_path / "widow.pdf")
    assert "Filler56" in first
    assert ("FirstLine" in second) is (widow_control is not False)  # moved down with the rest, or left alone


def test_text_no_renderer_reaches_is_reported(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delitem(blocks.BLOCK_RENDERERS, qn("w:tbl"))
    document = docx.Document()
    document.add_paragraph("kept")
    document.add_table(rows=1, cols=1).cell(0, 0).text = "Tablebound"
    document.save(tmp_path / "table.docx")
    rendered = convert_file(tmp_path / "table.docx", tmp_path / "table.pdf")
    [warning] = rendered.warnings
    assert "10 character(s) of text were not rendered: 'Tablebound'" in warning


def test_table_row_taller_than_a_page_splits(tmp_path: Path) -> None:
    document = docx.Document()
    cell = document.add_table(rows=1, cols=2).cell(0, 0)
    for line in range(150):
        cell.add_paragraph(f"Line{line}")
    document.save(tmp_path / "long.docx")
    rendered = convert_file(tmp_path / "long.docx", tmp_path / "long.pdf")
    pages = pdf_pages_text(tmp_path / "long.pdf")
    assert rendered.warnings == () and len(pages) > 1
    assert "Line0" in pages[0] and "Line149" in pages[-1]


def test_corrupt_file_fails_without_writing(tmp_path: Path) -> None:
    source = tmp_path / "corrupt.docx"
    source.write_bytes(b"this is not a zip file")
    with pytest.raises(ConversionFailed, match="BadZipFile|PackageNotFound"):
        convert_file(source, tmp_path / "corrupt.pdf")
    assert not (tmp_path / "corrupt.pdf").exists()


def test_encrypted_or_legacy_file_is_named_as_such(tmp_path: Path) -> None:
    source = tmp_path / "locked.docx"
    source.write_bytes(bytes.fromhex("d0cf11e0a1b11ae1") + b"\0" * 512)
    with pytest.raises(ConversionFailed, match="password-protected, or a legacy .doc"):
        convert_file(source, tmp_path / "locked.pdf")


@pytest.mark.parametrize(
    ("value", "fmt", "expected"),
    [(1, "lowerLetter", "a"), (27, "lowerLetter", "aa"), (1994, "upperRoman", "MCMXCIV"), (4, "lowerRoman", "iv"),
     (7, "decimalZero", "07"), (3, "ordinalText", "3")],
)
def test_number_formats(value: int, fmt: str, expected: str) -> None:
    assert format_number(value, fmt) == expected
