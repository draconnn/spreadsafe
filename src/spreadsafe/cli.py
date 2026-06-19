from __future__ import annotations

from pathlib import Path
import shutil

import typer

from spreadsafe.detectors import load_config
from spreadsafe.mapping import PseudonymMapper
from spreadsafe.reporter import write_reports
from spreadsafe.sanitizer import Sanitizer
from spreadsafe.scanner import scan_directory
from spreadsafe.validators import ValidationResult, validate_output

app = typer.Typer(help="Create Codex-safe sanitized spreadsheet packages.")


@app.command()
def scan(input_dir: Path, out: Path = typer.Option(..., "--out")) -> None:
    try:
        _ensure_input_directory(input_dir)
        _ensure_input_output_do_not_overlap(input_dir, out)
        config = load_config(input_dir / "spreadsafe.yml")
        sanitizer = Sanitizer(config, PseudonymMapper(seed=str(input_dir.resolve())))
        _ensure_output_is_owned_or_empty(out)
        sanitized_dir = out / "sanitized"
        reports_dir = out / "reports"
        _clear_generated_directory(sanitized_dir)
        _clear_generated_directory(reports_dir)
        sanitized_dir.mkdir(parents=True, exist_ok=True)
        reports_dir.mkdir(parents=True, exist_ok=True)
        sanitizer.sanitize_directory(input_dir, sanitized_dir)
        reports = scan_directory(
            sanitized_dir,
            path_prefix="sanitized",
            max_sample_rows_per_sheet=config.max_sample_rows_per_sheet,
        )
        write_reports(reports, sanitizer.risks, reports_dir, config)
        (out / ".spreadsafe-package").write_text("spreadsafe\n", encoding="utf-8")
        result = validate_output(out, config)
        if not result.passed:
            raise ValueError("; ".join(result.issues))
    except ValueError as exc:
        typer.echo(f"error: {exc}", err=True)
        raise typer.Exit(2) from exc
    typer.echo(f"Wrote reports to {reports_dir}")


@app.command()
def sanitize(input_dir: Path, out: Path = typer.Option(..., "--out")) -> None:
    try:
        _ensure_input_directory(input_dir)
        config = load_config(input_dir / "spreadsafe.yml")
        sanitizer = Sanitizer(config)
        sanitized_dir = out / "sanitized"
        _ensure_input_output_do_not_overlap(input_dir, out)
        _ensure_output_is_owned_or_empty(out)
        _clear_generated_directory(sanitized_dir)
        _clear_generated_directory(out / "reports")
        sanitized_dir.mkdir(parents=True, exist_ok=True)
        sanitizer.sanitize_directory(input_dir, sanitized_dir)
        (out / ".spreadsafe-package").write_text("spreadsafe\n", encoding="utf-8")
    except ValueError as exc:
        typer.echo(f"error: {exc}", err=True)
        raise typer.Exit(2) from exc
    typer.echo(f"Wrote sanitized files to {out / 'sanitized'}")


@app.command()
def validate(out: Path, config: Path | None = typer.Option(None, "--config")) -> None:
    try:
        loaded_config = load_config(config)
        result = validate_output(out, loaded_config)
    except ValueError as exc:
        typer.echo(f"error: {exc}", err=True)
        raise typer.Exit(2) from exc
    for warning in result.warnings:
        typer.echo(f"warning: {warning}", err=True)
    if not result.passed:
        for issue in result.issues:
            typer.echo(f"error: {issue}", err=True)
        raise typer.Exit(1)
    typer.echo("Validation passed")


@app.command(name="package")
def package_command(input_dir: Path, out: Path = typer.Option(..., "--out")) -> None:
    try:
        result = package_directory(input_dir, out)
    except ValueError as exc:
        typer.echo(f"error: {exc}", err=True)
        raise typer.Exit(2) from exc
    if not result.passed:
        for issue in result.issues:
            typer.echo(f"error: {issue}", err=True)
        raise typer.Exit(1)
    for warning in result.warnings:
        typer.echo(f"warning: {warning}", err=True)
    typer.echo(f"Wrote Codex-safe package to {out}")


def package_directory(input_dir: Path, output_dir: Path) -> ValidationResult:
    _ensure_input_directory(input_dir)
    _ensure_input_output_do_not_overlap(input_dir, output_dir)
    config = load_config(input_dir / "spreadsafe.yml")
    sanitizer = Sanitizer(config, PseudonymMapper(seed=str(input_dir.resolve())))
    sanitized_dir = output_dir / "sanitized"
    reports_dir = output_dir / "reports"
    _ensure_output_is_owned_or_empty(output_dir)
    _clear_generated_directory(sanitized_dir)
    _clear_generated_directory(reports_dir)
    sanitized_dir.mkdir(parents=True, exist_ok=True)
    reports_dir.mkdir(parents=True, exist_ok=True)

    sanitizer.sanitize_directory(input_dir, sanitized_dir)
    reports = scan_directory(
        sanitized_dir,
        path_prefix="sanitized",
        max_sample_rows_per_sheet=config.max_sample_rows_per_sheet,
    )
    write_reports(reports, sanitizer.risks, reports_dir, config)
    (output_dir / ".spreadsafe-package").write_text("spreadsafe\n", encoding="utf-8")
    return validate_output(output_dir, config)


def _ensure_input_directory(input_dir: Path) -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise ValueError(f"Input directory does not exist: {input_dir}")


def _ensure_input_output_do_not_overlap(input_dir: Path, output_dir: Path) -> None:
    resolved_input = input_dir.resolve()
    resolved_output = output_dir.resolve()
    if resolved_output == resolved_input or resolved_output.is_relative_to(resolved_input):
        raise ValueError("Output directory cannot be inside input directory")
    if resolved_input.is_relative_to(resolved_output):
        raise ValueError("Input directory cannot be inside output directory")


def _clear_generated_directory(path: Path) -> None:
    if not path.exists():
        return
    if path.is_symlink():
        raise ValueError(f"Cannot prepare output directory because {path} is a symlink")
    if not path.is_dir():
        raise ValueError(f"Cannot prepare output directory because {path} is not a directory")
    for child in path.rglob("*"):
        if child.is_symlink():
            raise ValueError(f"Cannot prepare output directory because {child} is a symlink")
    shutil.rmtree(path)


def _ensure_output_is_owned_or_empty(output_dir: Path) -> None:
    if not output_dir.exists():
        return
    if output_dir.is_symlink():
        raise ValueError(f"Cannot use output path because {output_dir} is a symlink")
    if not output_dir.is_dir():
        raise ValueError(f"Cannot use output path because {output_dir} is not a directory")
    marker = output_dir / ".spreadsafe-package"
    allowed_root_entries = {".gitignore", ".spreadsafe-package", "reports", "sanitized"}
    for child in output_dir.iterdir():
        if child.name not in allowed_root_entries:
            raise ValueError(f"Refusing to use existing output directory containing unmanaged files: {output_dir}")
    for child_name in ("sanitized", "reports"):
        child = output_dir / child_name
        if child.is_symlink():
            raise ValueError(f"Cannot prepare output directory because {child} is a symlink")
        if child.exists() and not child.is_dir():
            raise ValueError(f"Cannot prepare output directory because {child} is not a directory")
        if not marker.exists() and child.exists() and any(child.iterdir()):
            raise ValueError(
                f"Refusing to clear existing output directory without spreadsafe marker: {output_dir}"
            )


def main() -> None:
    app()


if __name__ == "__main__":
    main()
