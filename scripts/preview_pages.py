"""Convert a .docx and render each PDF page to PNG, to look at the layout by eye.

    uv run python scripts/preview_pages.py <in.docx> <out-dir>

With `--fixture`, <in.docx> is first written as the test suite's rich fixture.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pypdfium2

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tests"))
from conftest import write_rich_docx  # noqa: E402  (the tests folder is not a package)

from docx_to_pdf.rendering import convert_file  # noqa: E402


def main(source: Path, out_dir: Path, fixture: bool) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    if fixture:
        write_rich_docx(source)
    pdf = out_dir / f"{source.stem}.pdf"
    rendered = convert_file(source, pdf)
    for warning in rendered.warnings:
        print(f"warning: {warning}")
    document = pypdfium2.PdfDocument(str(pdf))
    for index, page in enumerate(document, start=1):
        image_path = out_dir / f"{source.stem}-page{index}.png"
        page.render(scale=1.2).to_pil().save(image_path)
        print(image_path)
    document.close()


if __name__ == "__main__":
    arguments = [a for a in sys.argv[1:] if a != "--fixture"]
    main(Path(arguments[0]), Path(arguments[1]), "--fixture" in sys.argv)
