"""The public API at the top of the package, as an installed user would call it."""

from __future__ import annotations

from pathlib import Path

import pytest
from conftest import pdf_pages_text, write_docx

import docx_to_pdf


def test_convert_file_takes_a_string_and_writes_beside_it(tmp_path: Path) -> None:
    source = write_docx(tmp_path / "letter.docx", "Dear reader")
    rendered = docx_to_pdf.convert_file(str(source))
    assert rendered.warnings == ()
    assert "Dear reader" in pdf_pages_text(tmp_path / "letter.pdf")[0]


def test_convert_file_to_a_chosen_target(tmp_path: Path) -> None:
    source = write_docx(tmp_path / "letter.docx")
    docx_to_pdf.convert_file(source, tmp_path / "out.pdf")
    assert (tmp_path / "out.pdf").read_bytes().startswith(b"%PDF-")
    assert not (tmp_path / "letter.pdf").exists()


def test_render_docx_is_bytes_in_bytes_out(tmp_path: Path) -> None:
    rendered = docx_to_pdf.render_docx(write_docx(tmp_path / "letter.docx").read_bytes())
    assert rendered.pdf.startswith(b"%PDF-")


def test_a_bad_document_raises_the_package_error(tmp_path: Path) -> None:
    source = tmp_path / "bad.docx"
    source.write_bytes(b"not a zip")
    with pytest.raises(docx_to_pdf.ConversionFailed):
        docx_to_pdf.convert_file(source)
    assert issubclass(docx_to_pdf.ConversionFailed, docx_to_pdf.DocxToPdfError)
