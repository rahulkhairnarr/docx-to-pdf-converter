# Changelog

All notable changes to this project are listed here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[Semantic Versioning](https://semver.org/).

## [0.1.0] - 2026-11-27

First release.

- **Conversion:** `.docx` to PDF in pure Python, with python-docx and reportlab.
  No Word or LibreOffice is needed.
- **Interfaces:** the `docx-to-pdf convert` command, and a Python API
  (`convert_file`, `render_docx`, and `service` for batches).
- **Layout:** sections and columns; headers and footers with page-number fields;
  paragraph formatting, borders and shading; tab stops and leaders; lists;
  tables; inline and floating pictures; text boxes; footnotes and endnotes.
- **Fonts:** bundled metric-compatible substitutes for Calibri, Arial, Times New
  Roman, Courier New and Cambria, with a DejaVu Sans fallback.
- **Warnings:** every piece of text that could not be drawn is reported, never
  dropped silently.
