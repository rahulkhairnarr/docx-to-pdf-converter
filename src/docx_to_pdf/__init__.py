"""Convert Word (.docx) documents to PDF in pure Python: no Word, LibreOffice or other
program needed.

    from docx_to_pdf import convert_file
    rendered = convert_file("letter.docx")          # writes letter.pdf beside it
    rendered.warnings                               # anything that could not be drawn

For bytes in, bytes out, use `render_docx(data).pdf`. For batches with validation
up front, use `service.plan_conversions` and `service.convert`. From the shell, the
same as `docx-to-pdf convert letter.docx`.
"""

from importlib.metadata import PackageNotFoundError, version

from .errors import ConversionFailed, DocxToPdfError, TargetConflict, UnsupportedFileType
from .rendering import RenderedDocument, convert_file, render_docx

try:
    __version__ = version("docx-to-pdf-converter")
except PackageNotFoundError:  # running from a source checkout that isn't installed
    __version__ = "0+unknown"

__all__ = [
    "ConversionFailed",
    "DocxToPdfError",
    "RenderedDocument",
    "TargetConflict",
    "UnsupportedFileType",
    "__version__",
    "convert_file",
    "render_docx",
]
