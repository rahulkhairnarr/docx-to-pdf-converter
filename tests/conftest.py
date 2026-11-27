"""Fixture documents, built with python-docx (plus raw XML where it has no API), and
PDF readers built on pypdfium2. Every document is generated here; none is real."""

from __future__ import annotations

import io
from pathlib import Path

import docx
import pypdfium2
import pytest
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import parse_xml
from docx.shared import Inches
from PIL import Image

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
WPS = ('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
       'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
       'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
       'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
       'xmlns:v="urn:schemas-microsoft-com:vml"')


def write_docx(path: Path, text: str = "Hello from docx-to-pdf") -> Path:
    document = docx.Document()
    document.add_paragraph(text)
    document.save(path)
    return path


def write_rich_docx(path: Path) -> Path:
    """One of each construct the renderer handles. Every word in RICH_WORDS must reach
    the PDF; the words in DELETED_WORDS must not."""
    document = docx.Document()
    document.add_heading("Rehabilitation Plan Alpha", level=1)

    paragraph = document.add_paragraph("Plain ")
    paragraph.add_run("bold").bold = True
    paragraph.add_run(" and ")
    paragraph.add_run("italic").italic = True
    paragraph.add_run(" with markup characters A&B <tag> and   spaces.")

    document.add_paragraph("Bullet first", style="List Bullet")
    document.add_paragraph("Bullet second", style="List Bullet")
    document.add_paragraph("Numbered first", style="List Number")
    document.add_paragraph("Numbered second", style="List Number")

    body = document.element.body  # raw XML goes before the body's final sectPr, as Word writes it
    body.sectPr.addprevious(parse_xml(
        f'<w:p {W} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:hyperlink r:id="rIdMissing"><w:r><w:t>Hyperlinked</w:t></w:r></w:hyperlink>'
        '<w:ins w:id="1" w:author="x"><w:r><w:t xml:space="preserve"> Inserted</w:t></w:r></w:ins>'
        '<w:del w:id="2" w:author="x"><w:r><w:delText>Removed</w:delText></w:r></w:del>'
        '<w:moveFrom w:id="3" w:author="x"><w:r><w:t>Moveaway</w:t></w:r></w:moveFrom>'
        '<w:r><w:rPr><w:vanish/></w:rPr><w:t>Invisible</w:t></w:r>'
        '<w:sdt><w:sdtContent><w:r><w:t xml:space="preserve"> Controlled</w:t></w:r></w:sdtContent></w:sdt>'
        '<w:fldSimple w:instr="DATE"><w:r><w:t xml:space="preserve"> Fielded</w:t></w:r></w:fldSimple>'
        '</w:p>'
    ))
    body.sectPr.addprevious(parse_xml(
        f'<w:sdt {W}><w:sdtContent><w:p><w:r><w:t>Blockcontrol</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    ))
    body.sectPr.addprevious(parse_xml(
        f'<w:p {W} {WPS}><w:r><mc:AlternateContent>'
        '<mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/>'
        '<a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>'
        '<w:p><w:r><w:t>Boxed</w:t></w:r></w:p>'
        '</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>'
        '<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>'
        '<w:p><w:r><w:t>Boxed</w:t></w:r></w:p>'
        '</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>'
        '</mc:AlternateContent></w:r></w:p>'
    ))

    table = document.add_table(rows=3, cols=3)
    table.style = "Table Grid"
    for row_index, row in enumerate(table.rows):
        for column_index, cell in enumerate(row.cells):
            cell.text = f"Cell{row_index}{column_index}"
    merged = table.cell(0, 0).merge(table.cell(0, 1))  # horizontal
    merged.text = "Mergedwide"
    tall = table.cell(1, 2).merge(table.cell(2, 2))  # vertical
    tall.text = "Mergedtall"
    nested = table.cell(2, 0).add_table(rows=1, cols=1)
    nested.cell(0, 0).text = "Nestedcell"

    image = io.BytesIO()
    Image.new("RGB", (40, 20), "red").save(image, format="PNG")
    image.seek(0)
    document.add_picture(image, width=Inches(1))

    add_note(document, "footnote", "Footnotetext")
    document.add_paragraph("Before break")
    document.add_page_break()
    document.add_paragraph("After break")

    landscape = document.add_section(WD_SECTION_START.NEW_PAGE)
    landscape.orientation = WD_ORIENT.LANDSCAPE
    landscape.page_width, landscape.page_height = landscape.page_height, landscape.page_width
    document.add_paragraph("Landscape section")

    footer = document.sections[0].footer.paragraphs[0]
    footer.text = "Page "
    footer._p.append(parse_xml(f'<w:r {W}><w:fldChar w:fldCharType="begin"/></w:r>'))
    footer._p.append(parse_xml(f'<w:r {W}><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'))
    footer._p.append(parse_xml(f'<w:r {W}><w:fldChar w:fldCharType="separate"/></w:r>'))
    footer._p.append(parse_xml(f'<w:r {W}><w:t>99</w:t></w:r>'))
    footer._p.append(parse_xml(f'<w:r {W}><w:fldChar w:fldCharType="end"/></w:r>'))
    document.sections[0].header.paragraphs[0].text = "Headertext"

    document.save(path)
    return path


_NOTES = {
    "footnote": ("/word/footnotes.xml", RT.FOOTNOTES),
    "endnote": ("/word/endnotes.xml", RT.ENDNOTES),
}


def add_note(document: docx.document.Document, kind: str, text: str) -> None:
    """Append a paragraph citing a new footnote/endnote whose body is `text`. python-docx
    has no notes API, so the notes part is added through its OPC layer."""
    partname, relationship = _NOTES[kind]
    try:
        part = document.part.part_related_by(relationship)
        notes = parse_xml(part.blob) if not hasattr(part, "element") else part.element
    except KeyError:
        notes = parse_xml(
            f'<w:{kind}s {W}><w:{kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p>'
            f'</w:{kind}><w:{kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/>'
            f'</w:r></w:p></w:{kind}></w:{kind}s>'
        )
        part = XmlPart(PackURI(partname), f"application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}s+xml",
                       notes, document.part.package)
        document.part.relate_to(part, relationship)
    note_id = str(len(notes) - 1)
    notes.append(parse_xml(
        f'<w:{kind} {W} w:id="{note_id}"><w:p><w:r><w:{kind}Ref/></w:r>'
        f'<w:r><w:t xml:space="preserve"> {text}</w:t></w:r></w:p></w:{kind}>'
    ))
    citing = document.add_paragraph("Cited claim")
    citing._p.append(parse_xml(f'<w:r {W}><w:{kind}Reference w:id="{note_id}"/></w:r>'))


RICH_WORDS = (
    "Rehabilitation", "Plain", "bold", "italic", "A&B", "<tag>", "spaces.", "Bullet", "Numbered",
    "Hyperlinked", "Inserted", "Controlled", "Fielded", "Blockcontrol", "Boxed", "Mergedwide",
    "Mergedtall", "Cell10", "Nestedcell", "Before", "After", "Landscape", "Headertext", "Cited claim",
    "Footnotetext",
)
DELETED_WORDS = ("Removed", "Moveaway", "Invisible", "99")


def pdf_pages_text(path: Path) -> list[str]:
    document = pypdfium2.PdfDocument(str(path))
    try:
        return [page.get_textpage().get_text_bounded() for page in document]
    finally:
        document.close()


def pdf_text_x(path: Path, text: str) -> tuple[float, float]:
    """(left, right) in points of the first occurrence of `text` on the first page."""
    document = pypdfium2.PdfDocument(str(path))
    try:
        textpage = document[0].get_textpage()
        index, count = textpage.search(text).get_next()
        return textpage.get_charbox(index)[0], textpage.get_charbox(index + count - 1)[2]
    finally:
        document.close()


def pdf_page_sizes(path: Path) -> list[tuple[float, float]]:
    document = pypdfium2.PdfDocument(str(path))
    try:
        return [page.get_size() for page in document]
    finally:
        document.close()


@pytest.fixture
def docx_file(tmp_path: Path) -> Path:
    return write_docx(tmp_path / "letter.docx")


@pytest.fixture
def rich_docx(tmp_path: Path) -> Path:
    return write_rich_docx(tmp_path / "rich.docx")
