from __future__ import annotations

from pathlib import Path
import shutil
from tempfile import TemporaryDirectory

import typer

from spreadsafe.detectors import Config, Detector, load_config
from spreadsafe.mapping import PseudonymMapper
from spreadsafe.reporter import write_reports
from spreadsafe.sanitizer import Sanitizer
from spreadsafe.scanner import scan_directory
from spreadsafe.validators import ValidationResult, _is_safe_generated_value, validate_output

app = typer.Typer(help="Create Codex-safe sanitized spreadsheet packages.")


@app.command()
def scan(input_dir: Path, out: Path = typer.Option(..., "--out")) -> None:
    try:
        result = package_directory(input_dir, out)
    except ValueError as exc:
        typer.echo(f"error: {exc}", err=True)
        raise typer.Exit(2) from exc
    _exit_if_validation_failed(result)
    typer.echo(f"Wrote reports to {out / 'reports'}")


@app.command()
def sanitize(input_dir: Path, out: Path = typer.Option(..., "--out")) -> None:
    try:
        _ensure_input_directory(input_dir)
        config = load_config(input_dir / "spreadsafe.yml")
        sanitizer = Sanitizer(config)
        _ensure_input_output_do_not_overlap(input_dir, out)
        _ensure_output_is_owned_or_empty(out, config)
        with _temporary_output(out) as staged_name:
            staged_out = Path(staged_name)
            sanitized_dir = staged_out / "sanitized"
            sanitized_dir.mkdir(parents=True, exist_ok=True)
            sanitizer.sanitize_directory(input_dir, sanitized_dir)
            (staged_out / ".spreadsafe-package").write_text("spreadsafe\n", encoding="utf-8")
            result = validate_output(staged_out, config)
            if result.passed:
                _replace_generated_output(staged_out, out, ("sanitized", "reports", ".spreadsafe-package"))
    except ValueError as exc:
        typer.echo(f"error: {exc}", err=True)
        raise typer.Exit(2) from exc
    _exit_if_validation_failed(result)
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
    _exit_if_validation_failed(result)
    typer.echo(f"Wrote Codex-safe package to {out}")


def package_directory(input_dir: Path, output_dir: Path) -> ValidationResult:
    _ensure_input_directory(input_dir)
    _ensure_input_output_do_not_overlap(input_dir, output_dir)
    config = load_config(input_dir / "spreadsafe.yml")
    sanitizer = Sanitizer(config, PseudonymMapper(seed=str(input_dir.resolve())))
    _ensure_output_is_owned_or_empty(output_dir, config)
    with _temporary_output(output_dir) as staged_name:
        staged_output = Path(staged_name)
        sanitized_dir = staged_output / "sanitized"
        reports_dir = staged_output / "reports"
        sanitized_dir.mkdir(parents=True, exist_ok=True)
        reports_dir.mkdir(parents=True, exist_ok=True)

        sanitizer.sanitize_directory(input_dir, sanitized_dir)
        reports = scan_directory(
            sanitized_dir,
            path_prefix="sanitized",
            max_sample_rows_per_sheet=config.max_sample_rows_per_sheet,
        )
        write_reports(reports, sanitizer.risks, reports_dir, config)
        (staged_output / ".spreadsafe-package").write_text("spreadsafe\n", encoding="utf-8")
        result = validate_output(staged_output, config)
        if result.passed:
            _replace_generated_output(
                staged_output,
                output_dir,
                ("sanitized", "reports", ".spreadsafe-package"),
            )
        return result


def _exit_if_validation_failed(result: ValidationResult) -> None:
    for warning in result.warnings:
        typer.echo(f"warning: {warning}", err=True)
    if not result.passed:
        for issue in result.issues:
            typer.echo(f"error: {issue}", err=True)
        raise typer.Exit(1)


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


def _temporary_output(output_dir: Path) -> TemporaryDirectory[str]:
    output_dir.parent.mkdir(parents=True, exist_ok=True)
    return TemporaryDirectory(prefix=f".{output_dir.name}-", dir=output_dir.parent)


def _replace_generated_output(
    staged_output: Path,
    output_dir: Path,
    child_names: tuple[str, ...],
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    backups: list[tuple[Path, Path]] = []
    installed: list[Path] = []
    try:
        for child_name in child_names:
            child = output_dir / child_name
            if child.exists():
                backup = _backup_path(child)
                _replace_path(child, backup)
                backups.append((backup, child))
        for child_name in child_names:
            child = staged_output / child_name
            if child.exists():
                destination = output_dir / child_name
                _replace_path(child, destination)
                installed.append(destination)
    except Exception:
        for child in installed:
            _remove_path(child)
        for backup, child in reversed(backups):
            if backup.exists():
                _remove_path(child)
                _replace_path(backup, child)
        raise
    for backup, _child in backups:
        _remove_path(backup)


def _backup_path(path: Path) -> Path:
    candidate = path.with_name(f".{path.name}.backup")
    counter = 2
    while candidate.exists():
        candidate = path.with_name(f".{path.name}.backup-{counter}")
        counter += 1
    return candidate


def _remove_path(path: Path) -> None:
    if path.is_dir() and not path.is_symlink():
        shutil.rmtree(path)
    elif path.exists():
        path.unlink()


def _replace_path(source: Path, destination: Path) -> None:
    shutil.move(str(source), destination)


def _ensure_output_is_owned_or_empty(output_dir: Path, config: Config | None = None) -> None:
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
        if child.name == ".gitignore":
            if child.is_symlink():
                raise ValueError(f"Cannot use output directory because {child} is a symlink")
            if not child.is_file():
                raise ValueError("Existing output .gitignore is not a file")
            if config is not None:
                detector = Detector(config)
                for detection in detector.detect_text(child.read_text(encoding="utf-8", errors="ignore")):
                    if not _is_safe_generated_value(detection.value):
                        raise ValueError("Existing output .gitignore contains sensitive data")
    for child_name in ("sanitized", "reports"):
        child = output_dir / child_name
        if child.is_symlink():
            raise ValueError(f"Cannot prepare output directory because {child} is a symlink")
        if child.exists() and not child.is_dir():
            raise ValueError(f"Cannot prepare output directory because {child} is not a directory")
        if child.exists():
            for nested in child.rglob("*"):
                if nested.is_symlink():
                    raise ValueError(f"Cannot prepare output directory because {nested} is a symlink")
        if not marker.exists() and child.exists() and any(child.iterdir()):
            raise ValueError(
                f"Refusing to clear existing output directory without spreadsafe marker: {output_dir}"
            )


def main() -> None:
    app()


if __name__ == "__main__":
    main()
