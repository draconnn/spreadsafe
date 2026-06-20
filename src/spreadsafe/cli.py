from __future__ import annotations

from argparse import ArgumentParser, Namespace
from pathlib import Path
import shutil
import sys
from tempfile import TemporaryDirectory
from collections.abc import Callable, Sequence
from typing import cast

from spreadsafe.detectors import Config, Detector, load_config
from spreadsafe.mapping import PseudonymMapper
from spreadsafe.reporter import write_reports
from spreadsafe.sanitizer import Sanitizer
from spreadsafe.scanner import scan_directory
from spreadsafe.validators import (
    ALLOWED_PACKAGE_ROOT_ENTRIES,
    ValidationResult,
    _is_safe_generated_value,
    validate_output,
)


def scan(input_dir: Path, out: Path) -> int:
    try:
        _ensure_input_directory(input_dir)
        _ensure_input_output_do_not_overlap(input_dir, out)
        config = load_config(input_dir / "spreadsafe.yml")
        sanitizer = Sanitizer(config, PseudonymMapper(seed=str(input_dir.resolve())))
        with _temporary_output(out) as staged_name:
            staged_out = Path(staged_name)
            sanitized_dir = staged_out / "sanitized"
            reports_dir = staged_out / "reports"
            sanitized_dir.mkdir(parents=True, exist_ok=True)
            reports_dir.mkdir(parents=True, exist_ok=True)
            sanitizer.sanitize_directory(input_dir, sanitized_dir)
            reports = scan_directory(
                sanitized_dir,
                max_sample_rows_per_sheet=config.max_sample_rows_per_sheet,
            )
            write_reports(reports, sanitizer.risks, reports_dir, config)
            (staged_out / ".spreadsafe-reports").write_text("spreadsafe-reports\n", encoding="utf-8")
            _remove_path(sanitized_dir)
            _replace_generated_output(
                staged_out,
                out,
                ("sanitized", "reports", ".spreadsafe-package", ".spreadsafe-reports"),
                config=config,
            )
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    print(f"Wrote reports to {out / 'reports'}")
    return 0


def sanitize(input_dir: Path, out: Path) -> int:
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
                _replace_generated_output(
                    staged_out,
                    out,
                    ("sanitized", "reports", ".spreadsafe-package", ".spreadsafe-reports"),
                )
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    if _exit_if_validation_failed(result):
        return 1
    print(f"Wrote sanitized files to {out / 'sanitized'}")
    return 0


def validate(out: Path, config: Path | None = None) -> int:
    try:
        loaded_config = load_config(config)
        result = validate_output(out, loaded_config)
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    if _exit_if_validation_failed(result):
        return 1
    print("Validation passed")
    return 0


def package_command(input_dir: Path, out: Path) -> int:
    try:
        result = package_directory(input_dir, out)
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    if _exit_if_validation_failed(result):
        return 1
    print(f"Wrote Codex-safe package to {out}")
    return 0


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
                ("sanitized", "reports", ".spreadsafe-package", ".spreadsafe-reports"),
            )
        return result


def _exit_if_validation_failed(result: ValidationResult) -> bool:
    for warning in result.warnings:
        print(f"warning: {warning}", file=sys.stderr)
    if not result.passed:
        for issue in result.issues:
            print(f"error: {issue}", file=sys.stderr)
        return True
    return False


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
    *,
    config: Config | None = None,
) -> None:
    _ensure_output_is_owned_or_empty(output_dir, config)
    output_dir.mkdir(parents=True, exist_ok=True)
    backups: list[tuple[Path, Path]] = []
    installed: list[Path] = []
    try:
        for child_name in child_names:
            child = output_dir / child_name
            _ensure_managed_output_child_is_safe(child)
            if child.exists():
                backup = _backup_path(child)
                _replace_path(child, backup)
                backups.append((backup, child))
        for child_name in child_names:
            child = staged_output / child_name
            if child.exists():
                destination = output_dir / child_name
                _ensure_managed_output_child_is_safe(destination)
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


def _ensure_managed_output_child_is_safe(path: Path) -> None:
    if path.is_symlink():
        raise ValueError(f"Cannot replace managed output path because {path} is a symlink")
    if not path.exists():
        return
    if path.name in {"sanitized", "reports"} and not path.is_dir():
        raise ValueError(f"Cannot replace managed output path because {path} is not a directory")
    if path.name == ".spreadsafe-package":
        if not path.is_file():
            raise ValueError("Existing output package marker is not a file")
        if path.read_text(encoding="utf-8", errors="ignore") != "spreadsafe\n":
            raise ValueError("Existing output package marker content is invalid")
    if path.name == ".spreadsafe-reports":
        if not path.is_file():
            raise ValueError("Existing output reports marker is not a file")
        if path.read_text(encoding="utf-8", errors="ignore") != "spreadsafe-reports\n":
            raise ValueError("Existing output reports marker content is invalid")


def _ensure_output_is_owned_or_empty(
    output_dir: Path,
    config: Config | None = None,
) -> None:
    if not output_dir.exists():
        return
    if output_dir.is_symlink():
        raise ValueError(f"Cannot use output path because {output_dir} is a symlink")
    if not output_dir.is_dir():
        raise ValueError(f"Cannot use output path because {output_dir} is not a directory")
    marker = output_dir / ".spreadsafe-package"
    if marker.is_symlink():
        raise ValueError(f"Cannot use output directory because {marker} is a symlink")
    if marker.exists():
        if not marker.is_file():
            raise ValueError("Existing output package marker is not a file")
        if marker.read_text(encoding="utf-8", errors="ignore") != "spreadsafe\n":
            raise ValueError("Existing output package marker content is invalid")
    reports_marker = output_dir / ".spreadsafe-reports"
    has_reports_marker = False
    if reports_marker.exists():
        if not reports_marker.is_file():
            raise ValueError("Existing output reports marker is not a file")
        if reports_marker.read_text(encoding="utf-8", errors="ignore") != "spreadsafe-reports\n":
            raise ValueError("Existing output reports marker content is invalid")
        has_reports_marker = True
    for child in output_dir.iterdir():
        if child.name not in ALLOWED_PACKAGE_ROOT_ENTRIES:
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
        if (
            not marker.exists()
            and child.exists()
            and any(child.iterdir())
            and (
                child.name != "reports"
                or not has_reports_marker
                or not _is_generated_reports_dir(child)
            )
        ):
            raise ValueError(
                f"Refusing to clear existing output directory without spreadsafe marker: {output_dir}"
            )


def _is_generated_reports_dir(path: Path) -> bool:
    report_names = {"workbook-report.md", "workbook-report.json", "risk-report.md"}
    children = list(path.iterdir())
    return {child.name for child in children} == report_names and all(
        child.is_file() for child in children
    )


def _build_parser() -> ArgumentParser:
    parser = ArgumentParser(description="Create Codex-safe sanitized spreadsheet packages.")
    subparsers = parser.add_subparsers(dest="command", required=True)
    for name in ("scan", "sanitize", "package"):
        command = subparsers.add_parser(name)
        command.add_argument("input_dir", type=Path)
        command.add_argument("--out", required=True, type=Path)
        command.set_defaults(func=_run_input_output_command)
    validate_parser = subparsers.add_parser("validate")
    validate_parser.add_argument("out", type=Path)
    validate_parser.add_argument("--config", type=Path)
    validate_parser.set_defaults(func=_run_validate_command)
    return parser


def _run_input_output_command(args: Namespace) -> int:
    if args.command == "scan":
        return scan(args.input_dir, args.out)
    if args.command == "sanitize":
        return sanitize(args.input_dir, args.out)
    return package_command(args.input_dir, args.out)


def _run_validate_command(args: Namespace) -> int:
    return validate(args.out, args.config)


def main(argv: Sequence[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    func = cast(Callable[[Namespace], int], args.func)
    return func(args)


if __name__ == "__main__":
    raise SystemExit(main())
