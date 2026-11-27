"""Pure-Python .docx -> PDF: python-docx reads the document, reportlab draws the PDF.
No Word, LibreOffice or other program is needed.

Module map: `document` (sections, pages, headers/footers, warnings) -> `blocks`
(paragraphs, tables; the BLOCK_RENDERERS registry) -> `inline` (runs, fields,
pictures) with `styles` (formatting inheritance), `numbering` (list labels),
`fonts` (standard-font mapping) and `context` (shared state and the text guard).
"""

from __future__ import annotations

import os
from pathlib import Path

from ..errors import ConversionFailed
from .document import RenderedDocument, render_docx

# The OLE container header: an encrypted .docx, or a legacy .doc renamed to .docx.
_OLE_SIGNATURE = bytes.fromhex("d0cf11e0a1b11ae1")

__all__ = ["RenderedDocument", "convert_file", "render_docx"]


def convert_file(source: str | os.PathLike, target: str | os.PathLike | None = None) -> RenderedDocument:
    """Render `source` and write the PDF to `target` (default: beside it, as .pdf), all
    at once, so a failure leaves no partial file. Raises ConversionFailed for anything
    wrong with this document."""
    source = Path(source)
    target = source.with_suffix(".pdf") if target is None else Path(target)
    try:
        data = source.read_bytes()
    except OSError as error:  # locked by another program, permissions
        raise ConversionFailed(f"could not read the file: {error}") from error
    if data.startswith(_OLE_SIGNATURE):
        raise ConversionFailed("password-protected, or a legacy .doc file; neither can be read without Word")
    try:
        rendered = render_docx(data)
    except Exception as error:  # noqa: BLE001 - corrupt zip, bad XML, a layout reportlab can't fit:
        # all are this one document's problem, and must not stop the rest of a batch.
        raise ConversionFailed(f"{type(error).__name__}: {error}") from error
    target.write_bytes(rendered.pdf)
    return rendered
