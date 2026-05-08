"""Path resolution helpers for the visuals tools."""

from __future__ import annotations

from pathlib import Path

from security import resolve_local_path

from ._base import LAYOUT_RELATIVE_PATH


def _resolve_pbix_path(pbix_path: str, *, must_exist: bool) -> Path:
    return resolve_local_path(pbix_path, must_exist=must_exist, allowed_extensions={".pbix"})


def _resolve_extract_folder(extract_folder: str, *, must_exist: bool) -> Path:
    return resolve_local_path(extract_folder, must_exist=must_exist)


def _resolve_theme_path(theme_json_path: str) -> Path:
    return resolve_local_path(theme_json_path, must_exist=True, allowed_extensions={".json"})


def _layout_path(extract_folder: Path) -> Path:
    return extract_folder / LAYOUT_RELATIVE_PATH
