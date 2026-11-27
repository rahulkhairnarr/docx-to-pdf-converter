"""Service layer: the conversion use case, independent of how it is called.

The CLI (cli.py), and any application that embeds this package, are thin adapters
over `plan_conversions` + `convert`. All the checks live here once:
- every input is validated, and every target worked out, before anything is
  converted, so a typo in the last file costs nothing;
- one document failing does not stop the batch; its error is in its result.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from pathlib import Path

from .errors import ConversionFailed, TargetConflict, UnsupportedFileType
from .rendering import convert_file

# Only the Office Open XML format: legacy .doc is a binary format nothing pure-Python
# reads faithfully. .docm/.dotx are the same format with another content type; add
# them here once python-docx is checked to open them.
SUPPORTED_SUFFIXES = (".docx",)


@dataclass(frozen=True)
class Job:
    source: Path
    target: Path


@dataclass(frozen=True)
class ConversionResult:
    job: Job
    error: str | None = None
    # The PDF was written, but something in the document couldn't be drawn (and is named here).
    warnings: tuple[str, ...] = ()

    @property
    def ok(self) -> bool:
        return self.error is None


def plan_conversions(sources: Iterable[Path], output_dir: Path | None = None, overwrite: bool = False) -> list[Job]:
    """One job per source: `<stem>.pdf` in `output_dir`, or beside the source.
    Raises UnsupportedFileType or TargetConflict, before anything is converted."""
    jobs = [_job(Path(source).resolve(), output_dir.resolve() if output_dir else None) for source in sources]
    _check_targets(jobs, overwrite)
    return jobs


def convert(jobs: Iterable[Job]) -> Iterator[ConversionResult]:
    """Convert each job, yielding results as they finish so a caller can report progress."""
    for job in jobs:
        job.target.parent.mkdir(parents=True, exist_ok=True)
        try:
            rendered = convert_file(job.source, job.target)
        except ConversionFailed as error:
            yield ConversionResult(job, error=str(error))
        else:
            yield ConversionResult(job, warnings=rendered.warnings)


def _job(source: Path, output_dir: Path | None) -> Job:
    if source.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise UnsupportedFileType(
            f"unsupported file type {source.suffix!r} ({source.name}); supported: {', '.join(SUPPORTED_SUFFIXES)}"
        )
    return Job(source=source, target=(output_dir or source.parent) / f"{source.stem}.pdf")


def _check_targets(jobs: list[Job], overwrite: bool) -> None:
    # Compared case-insensitively: on Windows and macOS `A.pdf` and `a.pdf` are one file.
    counts = Counter(str(job.target).casefold() for job in jobs)
    if shared := sorted({job.target.name for job in jobs if counts[str(job.target).casefold()] > 1}):
        raise TargetConflict(f"several inputs would write the same PDF: {', '.join(shared)}")
    if not overwrite and (existing := [str(job.target) for job in jobs if job.target.exists()]):
        raise TargetConflict(f"already exists (pass --overwrite to replace): {', '.join(existing)}")
