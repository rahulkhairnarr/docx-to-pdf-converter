# Contributing

Thank you for helping. Issues and pull requests are welcome.

## Reporting a rendering problem

The most useful report includes:
- a small `.docx` that shows the problem, **with no private content**;
- the PDF this package produced, and what you expected (a screenshot of Word's
  rendering is ideal);
- any warnings `docx-to-pdf` printed.

## Development setup

You need [uv](https://docs.astral.sh/uv/).

```sh
git clone https://github.com/OWNER/docx-to-pdf-converter
cd docx-to-pdf-converter
uv sync                     # creates .venv with the package and the dev tools
uv run pytest               # the test suite
uv run docx-to-pdf --help   # the CLI from your checkout
```

The tests build their documents with python-docx (and raw XML where it has no
API), convert them, and read the PDFs back with pypdfium2. Positions are measured
in points, so a test can check, for example, that a right tab stop puts text
against the margin.

To look at the layout by eye, render each page to a PNG:

```sh
uv run python scripts/preview_pages.py some.docx out/
uv run python scripts/preview_pages.py fixture.docx out/ --fixture   # the test suite's rich document
```

## Making a change

- **One fix, one test.** Add a test that fails without your change. Generate any
  document you need in the test itself; never commit real documents.
- **Supporting a new construct** (say, block-level equations): write
  `render_x(element, context) -> list[Flowable]` in `rendering/blocks.py` and add
  the element's tag to `BLOCK_RENDERERS`. Inline constructs go in
  `rendering/inline.py` (`_walk` / `_run`). Until a construct is supported, the
  text guard reports its text as not rendered. Keep it that way: never drop text
  silently.
- **Adding a font:**
  1. Put the TTF files and the font's license in `src/docx_to_pdf/fonts/`.
  2. Add the family to `_FAMILIES` in `rendering/fonts.py`, and map Word's font
     name to it in `_SUBSTITUTES`.
  3. If it stands in for a Microsoft font, add that font's line height to
     `_WORD_LINE_HEIGHTS`. The line height is (usWinAscent + usWinDescent) /
     unitsPerEm, read from the original font.

  Only fonts whose license allows redistribution can be bundled.
- **Design notes:** if a change rests on a measurement or a reportlab
  behaviour you had to discover, add it to [docs/DESIGN.md](docs/DESIGN.md).

## Releasing

1. Update the version in `pyproject.toml`, and add the release to `CHANGELOG.md`.
2. Run `uv build`, which writes the wheel and sdist to `dist/`.
3. Run `uv publish`, with a PyPI token or trusted publishing.
