from __future__ import annotations

import os
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import TypeVar

from openpyxl.workbook.workbook import Workbook

from .exceptions import ConfigurationError


T = TypeVar("T")


def resolve_path(root: Path, value: str | Path) -> Path:
    """Resolve a config path without assuming POSIX-only path behavior."""

    path = Path(value).expanduser()
    if not path.is_absolute():
        path = root / path
    return path.resolve()


def validate_pdf_input_path(path: Path, allowed_root: Path | None = None) -> Path:
    resolved = path.expanduser().resolve()
    if allowed_root is not None and not _is_relative_to(resolved, allowed_root.resolve()):
        raise ConfigurationError(f"PDF path is outside the configured input folder: {resolved}")
    if resolved.suffix.lower() != ".pdf":
        raise ConfigurationError(f"Input file is not a PDF: {resolved}")
    if not resolved.exists():
        raise ConfigurationError(f"Input PDF does not exist: {resolved}")
    if not resolved.is_file():
        raise ConfigurationError(f"Input PDF path is not a file: {resolved}")
    return resolved


def validate_output_target(path: Path, protected_paths: list[Path] | None = None) -> Path:
    resolved = path.expanduser().resolve()
    for protected_path in protected_paths or []:
        if resolved == protected_path.expanduser().resolve():
            raise ConfigurationError(f"Refusing to overwrite protected source file: {resolved}")
    resolved.parent.mkdir(parents=True, exist_ok=True)
    return resolved


def atomic_save_workbook(workbook: Workbook, output_path: Path) -> None:
    """Save an Excel workbook atomically to avoid corrupting a prior output file."""

    target = validate_output_target(output_path)
    temp_path = _temporary_path(target)
    try:
        workbook.save(temp_path)
        os.replace(temp_path, target)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def atomic_write_path(output_path: Path, writer: Callable[[Path], T]) -> T:
    """Write to a temporary file in the target folder and replace on success."""

    target = validate_output_target(output_path)
    temp_path = _temporary_path(target)
    try:
        result = writer(temp_path)
        os.replace(temp_path, target)
        return result
    finally:
        if temp_path.exists():
            temp_path.unlink()


def _temporary_path(target: Path) -> Path:
    handle, raw_path = tempfile.mkstemp(prefix=f".{target.stem}.", suffix=target.suffix, dir=target.parent)
    os.close(handle)
    return Path(raw_path)


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False

