# docx-to-pdf-converter

Convert Word documents (`.docx`) to PDF **in pure Python**. No Microsoft Word,
LibreOffice or other program is needed, so it runs the same on Linux, macOS,
Windows, in containers and in serverless functions. [python-docx] reads the
document and [reportlab] draws the PDF.

Source documents are only read, never changed.

```sh
pip install docx-to-pdf-converter
docx-to-pdf convert letter.docx          # writes letter.pdf
```

```python
from docx_to_pdf import convert_file

rendered = convert_file("letter.docx")   # writes letter.pdf
print(rendered.warnings)                 # anything that could not be drawn
```

> **Status: alpha.** The layout is close to Word's, not identical: expect line
> and page breaks in different places in some documents. Anything the converter
> can't draw is reported as a warning rather than dropped silently. See
> [What it renders](#what-it-renders-and-what-it-doesnt).

## Install

Requires Python 3.11 or newer.

```sh
pip install docx-to-pdf-converter
# or straight from GitHub:
pip install "git+https://github.com/OWNER/docx-to-pdf-converter"
# or with uv:
uv add docx-to-pdf-converter
```

Dependencies are python-docx, reportlab, Pillow and Typer. All are pure Python
or ship wheels, so no system libraries are needed. The fonts the converter draws
with are bundled in the package; see [Fonts](#fonts).

## Command line

```sh
docx-to-pdf convert letter.docx report.docx      # letter.pdf, report.pdf beside the inputs
docx-to-pdf convert docs/*.docx -o pdfs/         # into pdfs/
docx-to-pdf convert letter.docx --overwrite      # replace an existing letter.pdf
python -m docx_to_pdf convert letter.docx        # the same, without the script on PATH
```

- **Checked before anything is converted.** These are usage errors and stop the
  whole run (exit code 2):
  - an unsupported file type, including legacy `.doc`;
  - a PDF that already exists, without `--overwrite`;
  - two inputs that would write the same PDF.
- **A document that can't be converted is skipped, and the rest still run.** That
  covers corrupt files, password-protected files and a `.doc` renamed to `.docx`.
  The reason goes to stderr and the exit code is 1. A failed file never leaves a
  partial PDF.
- **Warnings.** If something in a document could not be drawn, the PDF is still
  written and stderr says what was missed, for example:

  ```text
  report.docx: warning, 2 character(s) drawn as '?' (no bundled font has them, or their script needs shaping): '王' x1, 'س' x1
  ```

| Exit code | Meaning |
|---|---|
| 0 | Every file converted |
| 1 | At least one file failed; the others were converted |
| 2 | Usage error; nothing was converted |

## Python API

```python
from docx_to_pdf import ConversionFailed, convert_file, render_docx

# A file in, a file out. The target defaults to the source with a .pdf suffix.
rendered = convert_file("in/letter.docx", "out/letter.pdf")
for warning in rendered.warnings:
    print("not drawn:", warning)

# Bytes in, bytes out: no files involved.
pdf_bytes = render_docx(docx_bytes).pdf

# Every per-document problem (corrupt, encrypted, legacy .doc) is one exception.
try:
    convert_file("broken.docx")
except ConversionFailed as error:
    print(error)
```

For batches, `docx_to_pdf.service` validates every input and works out every
target before converting anything. One failure doesn't stop the rest:

```python
from pathlib import Path
from docx_to_pdf import service

jobs = service.plan_conversions([Path("a.docx"), Path("b.docx")], output_dir=Path("pdfs"))
for result in service.convert(jobs):
    print(result.job.target, "ok" if result.ok else result.error, result.warnings)
```

Another [Typer] application can mount the CLI with
`app.add_typer(docx_to_pdf.cli.app, name="docx-to-pdf")`.

## What it renders, and what it doesn't

**Rendered:**
- **Pages:** size, margins and orientation per section; multi-column sections,
  column breaks and the separator line.
- **Headers and footers:** first-page and odd/even variants. `PAGE`, `NUMPAGES`
  and `SECTIONPAGES` are filled in. Page numbering can restart per section and
  use formats such as `i, ii`.
- **Paragraphs:** alignment, indents, spacing, line spacing, borders, shading,
  keep-with-next and widow/orphan control.
- **Tab stops:** left, centre, right and decimal, with dot, hyphen and underscore
  leaders.
- **Text:** bold, italic, underline, strike, size, colour, highlight,
  superscript and subscript, and caps.
- **Lists:** numbered and bulleted, with labels computed as Word does (`1.`,
  `a)`, `iv.`, `1.2.3`, custom text, and restarts).
- **Tables:** merged cells, nested tables, cell margins, shading, header rows
  repeated across pages, and rows that split across pages.
- **Pictures:** inline in the line of text, and floating. Formats: PNG, JPEG,
  GIF, BMP, TIFF and WebP.
- **Other:** text boxes, hyperlinks, content controls, the displayed text of
  fields, and footnotes and endnotes.
- **Tracked changes** are shown as accepted (insertions kept, deletions dropped),
  with a warning.

**Never lost silently.** A guard checks that every piece of text in the body,
the headers and footers and the notes reached the PDF, and reports any that
didn't. Charts, SmartArt, embedded documents and undrawable images show a
`[... not shown]` placeholder and are counted in a warning.

**Approximated or not rendered:**
- **Floating images and text boxes** are placed where they are anchored. Text
  doesn't wrap around them.
- **Image rotation and flip** are ignored.
- **SVG** draws its PNG fallback; EMF/WMF show a placeholder.
- **Footnotes** are collected after the body, not placed at the foot of each page.
- **Equations** are drawn as plain text.
- **Not drawn:** shapes, page borders, per-cell border styles, line numbers and
  comments. Word doesn't print comments by default either.
- **Tabs in a wrapping paragraph** get a fixed gap. Tabs are laid out at their
  stops only in single lines.
- **Page fields in body text:** `PAGE` and `PAGEREF` show the value cached in the
  file.
- **Not supported:** legacy `.doc` and password-protected documents. Both are
  binary formats that need Word to read.

### Fonts

The converter brings its own fonts and embeds them in each PDF, so the output is
the same on every machine. Where a free font has exactly the same character
widths as Word's (measured), lines break where Word breaks them:

| Word font | Drawn with | Match |
|---|---|---|
| Calibri, Calibri Light | Carlito | identical widths and line height (Light uses the regular weight) |
| Arial | Liberation Sans | identical |
| Times New Roman | Liberation Serif | identical |
| Courier New, Consolas | Liberation Mono | identical for Courier New |
| Cambria | Caladea | looks alike, but about 6% narrower |
| Aptos and any other font | Liberation Sans, Serif or Mono, by font class | approximate: text reflows |

- Theme fonts, which is how Word documents usually name their fonts, are resolved
  through the document's theme. Line height follows Word's own figure for the
  original font.
- A character the chosen font lacks is drawn in DejaVu Sans. That covers accented
  Latin, Greek, Cyrillic, Vietnamese, ☐ ☒ ✓ and most symbols.
- Shown as `?`, with a warning:
  - Chinese, Japanese and Korean, because no CJK font is bundled;
  - scripts that need shaping (Arabic, Hebrew, Indic, Thai and similar). Drawn
    without shaping they would come out reversed or broken, so they are
    reported instead.

## How it works

See [docs/DESIGN.md](docs/DESIGN.md) for the design and the measurements behind
it. In short:

| Module | Role |
|---|---|
| `cli.py` | Typer adapter: arguments in, messages and exit codes out. No logic. |
| `service.py` | Batch use case: plan jobs (validation, targets, conflicts), then convert them one by one. |
| `errors.py` | The errors a caller may handle. All are `ValueError`s under `DocxToPdfError`. |
| `rendering/__init__.py` | `convert_file`: one document in, one PDF out. |
| `rendering/document.py` | Sections → page templates, the body, headers and footers, notes, page-number fields, warnings. |
| `rendering/blocks.py` | Paragraphs, tables, content controls, placeholders; the `BLOCK_RENDERERS` registry. |
| `rendering/inline.py` | Runs and what wraps them: hyperlinks, insertions, fields, note marks, pictures, text boxes. |
| `rendering/styles.py` | Formatting inheritance: run → character style → paragraph style → `basedOn` chain → document defaults. |
| `rendering/numbering.py` | List labels from numbering.xml. |
| `rendering/tabs.py` | Tab stops and leaders. |
| `rendering/borders.py` | Paragraph borders and shading. |
| `rendering/paragraphs.py` | Lines with pictures or mixed font sizes are sized one by one, as in Word. |
| `rendering/keeping.py` | Keep-with-next: a paragraph stays with the first line or row of what follows. |
| `rendering/fonts.py` | Word font → bundled font, Word's line heights, fallback and shaping checks. |
| `rendering/context.py` | Shared per-document state, and the text guard. |

## Contributing

Bug reports with a sample `.docx` (one with no private content) are the most
useful contribution. See [CONTRIBUTING.md](CONTRIBUTING.md) for the development
setup.

## License

The code is under the [MIT License](LICENSE).

The bundled fonts keep their own licenses. The texts are in
`src/docx_to_pdf/fonts/licenses/`:

| Fonts | License |
|---|---|
| Carlito, Caladea, Liberation Sans/Serif/Mono | SIL Open Font License 1.1 |
| DejaVu Sans | Bitstream Vera license, with DejaVu changes in the public domain |

[python-docx]: https://github.com/python-openxml/python-docx
[reportlab]: https://www.reportlab.com/opensource/
[Typer]: https://typer.tiangolo.com/
