from __future__ import annotations

from pathlib import Path

import pytest
from conftest import write_docx

from docx_to_pdf import service
from docx_to_pdf.errors import TargetConflict, UnsupportedFileType


def test_target_is_beside_the_source_by_default(docx_file: Path) -> None:
    [job] = service.plan_conversions([docx_file])
    assert job.target == docx_file.with_suffix(".pdf")


def test_target_goes_to_output_dir(docx_file: Path, tmp_path: Path) -> None:
    [job] = service.plan_conversions([docx_file], tmp_path / "out")
    assert job.target == tmp_path / "out" / "letter.pdf"


def test_paths_are_made_absolute(docx_file: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(docx_file.parent)
    [job] = service.plan_conversions([Path(docx_file.name)], Path("out"))
    assert job.source.is_absolute() and job.target.is_absolute()


def test_suffix_check_is_case_insensitive(tmp_path: Path) -> None:
    upper = write_docx(tmp_path / "LETTER.DOCX")
    [job] = service.plan_conversions([upper])
    assert job.target.name == "LETTER.pdf"


@pytest.mark.parametrize("name", ["notes.txt", "legacy.doc", "macros.docm"])
def test_rejects_unsupported_suffix(tmp_path: Path, name: str) -> None:
    with pytest.raises(UnsupportedFileType, match="supported: .docx"):
        service.plan_conversions([tmp_path / name])


def test_refuses_to_overwrite_an_existing_pdf(docx_file: Path) -> None:
    docx_file.with_suffix(".pdf").write_bytes(b"keep me")
    with pytest.raises(TargetConflict, match="already exists"):
        service.plan_conversions([docx_file])
    assert service.plan_conversions([docx_file], overwrite=True)


def test_refuses_two_inputs_writing_one_pdf(tmp_path: Path) -> None:
    (tmp_path / "a").mkdir()
    (tmp_path / "b").mkdir()
    first = write_docx(tmp_path / "a" / "same.docx")
    second = write_docx(tmp_path / "b" / "SAME.docx")
    with pytest.raises(TargetConflict, match="same PDF"):
        service.plan_conversions([first, second], tmp_path / "out")


def test_one_failure_does_not_stop_the_batch(tmp_path: Path) -> None:
    good = write_docx(tmp_path / "a.docx")
    corrupt = tmp_path / "bad.docx"
    corrupt.write_bytes(b"not a zip")
    last = write_docx(tmp_path / "c.docx")

    results = list(service.convert(service.plan_conversions([good, corrupt, last], tmp_path / "out")))

    assert [(result.job.source.name, result.ok) for result in results] == [
        ("a.docx", True), ("bad.docx", False), ("c.docx", True)
    ]
    assert sorted(path.name for path in (tmp_path / "out").iterdir()) == ["a.pdf", "c.pdf"]
