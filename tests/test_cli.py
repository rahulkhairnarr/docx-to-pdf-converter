from __future__ import annotations

from pathlib import Path

from conftest import write_docx
from typer.testing import CliRunner

from docx_to_pdf import cli

runner = CliRunner()


def test_convert_writes_the_pdf(docx_file: Path) -> None:
    result = runner.invoke(cli.app, ["convert", str(docx_file)])
    assert result.exit_code == 0, result.output
    assert docx_file.with_suffix(".pdf").read_bytes().startswith(b"%PDF-")
    assert "letter.docx -> " in result.output


def test_failures_are_reported_and_exit_1(tmp_path: Path) -> None:
    good = write_docx(tmp_path / "good.docx")
    corrupt = tmp_path / "bad.docx"
    corrupt.write_bytes(b"not a zip")
    locked = tmp_path / "locked.docx"
    locked.write_bytes(bytes.fromhex("d0cf11e0a1b11ae1") + b"\0" * 512)
    result = runner.invoke(cli.app, ["convert", str(good), str(corrupt), str(locked)])
    assert result.exit_code == 1
    assert "bad.docx: skipped," in result.output
    assert "locked.docx: skipped, password-protected" in result.output
    assert "Failed 2 of 3 file(s)" in result.output
    assert good.with_suffix(".pdf").exists()


def test_warnings_are_reported_and_the_pdf_still_written(tmp_path: Path) -> None:
    source = write_docx(tmp_path / "names.docx", "Wang 王")
    result = runner.invoke(cli.app, ["convert", str(source)])
    assert result.exit_code == 0
    assert "names.docx: warning, 1 character(s) drawn as '?'" in result.output
    assert source.with_suffix(".pdf").exists()


def test_usage_errors_convert_nothing(docx_file: Path, tmp_path: Path) -> None:
    text = tmp_path / "notes.txt"
    text.write_text("not a document")
    result = runner.invoke(cli.app, ["convert", str(docx_file), str(text)])
    assert result.exit_code == 2
    assert "unsupported file type '.txt'" in result.output
    assert not docx_file.with_suffix(".pdf").exists()


def test_existing_pdf_needs_overwrite(docx_file: Path) -> None:
    docx_file.with_suffix(".pdf").write_bytes(b"old")
    assert runner.invoke(cli.app, ["convert", str(docx_file)]).exit_code == 2
    assert runner.invoke(cli.app, ["convert", str(docx_file), "--overwrite"]).exit_code == 0
