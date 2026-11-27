"""`docx-to-pdf` CLI: a thin adapter over service.py. It only turns arguments into
service calls and results into output. To add a command, add one @app.command().

`app` is importable, so another Typer application can mount it:
`app.add_typer(docx_to_pdf.cli.app, name="docx-to-pdf")`.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Annotated

import typer

from . import service
from .errors import TargetConflict, UnsupportedFileType

app = typer.Typer(add_completion=False, help="Convert Word documents to PDF.", no_args_is_help=True)


@app.callback()
def _root() -> None:
    """Having a callback keeps `convert` a named subcommand while it is the only one."""


@app.command()
def convert(
    files: Annotated[
        list[Path],
        typer.Argument(exists=True, dir_okay=False, help=f"Word files ({', '.join(service.SUPPORTED_SUFFIXES)})"),
    ],
    output_dir: Annotated[
        Path | None,
        typer.Option("--output-dir", "-o", file_okay=False, help="Write the PDFs here (default: beside each input)"),
    ] = None,
    overwrite: Annotated[bool, typer.Option(help="Replace PDFs that already exist")] = False,
) -> None:
    """Convert each file to `<name>.pdf`. A file that fails is reported on stderr and
    skipped; the others are still converted, and the exit code is 1. Content that
    could not be drawn is reported on stderr as a warning; the PDF is still written."""
    try:
        jobs = service.plan_conversions(files, output_dir, overwrite)
    except (UnsupportedFileType, TargetConflict) as error:
        raise typer.BadParameter(str(error), param_hint="FILES") from error

    failed = 0
    for result in service.convert(jobs):
        name = result.job.source.name
        if not result.ok:
            failed += 1
            typer.echo(f"{name}: skipped, {result.error}", err=True)
            continue
        typer.echo(f"{name} -> {result.job.target}")
        for warning in result.warnings:
            typer.echo(f"{name}: warning, {warning}", err=True)
    if failed:
        typer.echo(f"Failed {failed} of {len(jobs)} file(s); see above", err=True)
        raise typer.Exit(1)


def main() -> int:
    """Dispatch to the matching command and return its exit code: the process boundary,
    where nothing may escape as a traceback."""
    try:
        app()
    except SystemExit as exit_request:
        return int(exit_request.code or 0)
    except Exception as error:  # noqa: BLE001 - CLI boundary, must not traceback
        print(f"error: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
