"""Tests for the shared atomic-write helpers (Phase 1 safety)."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from atomic_io import atomic_write_bytes, atomic_write_text, snapshot_once


class TestAtomicWriteBytes:
    def test_creates_new_file_without_backup(self, tmp_path: Path) -> None:
        target = tmp_path / "sub" / "file.bin"
        bak = atomic_write_bytes(target, b"data")
        assert target.read_bytes() == b"data"
        assert bak is None
        assert not list(tmp_path.rglob("*.tmp.*"))

    def test_backs_up_existing_content(self, tmp_path: Path) -> None:
        target = tmp_path / "file.bin"
        target.write_bytes(b"old")
        bak = atomic_write_bytes(target, b"new")
        assert target.read_bytes() == b"new"
        assert bak == tmp_path / "file.bin.bak"
        assert bak.read_bytes() == b"old"

    def test_backup_disabled(self, tmp_path: Path) -> None:
        target = tmp_path / "file.bin"
        target.write_bytes(b"old")
        bak = atomic_write_bytes(target, b"new", backup=False)
        assert bak is None
        assert not (tmp_path / "file.bin.bak").exists()

    def test_no_temp_residue(self, tmp_path: Path) -> None:
        target = tmp_path / "file.bin"
        atomic_write_bytes(target, b"one")
        atomic_write_bytes(target, b"two")
        assert sorted(p.name for p in tmp_path.iterdir()) == ["file.bin", "file.bin.bak"]


class TestAtomicWriteText:
    def test_writes_utf8_without_newline_translation(self, tmp_path: Path) -> None:
        target = tmp_path / "file.tmdl"
        atomic_write_text(target, "line1\nline2\n")
        assert target.read_bytes() == b"line1\nline2\n"


class TestSnapshotOnce:
    def test_creates_snapshot_first_time(self, tmp_path: Path) -> None:
        source = tmp_path / "Layout"
        source.write_bytes(b"pristine")
        snap = snapshot_once(source)
        assert snap == tmp_path / "Layout.orig"
        assert snap.read_bytes() == b"pristine"

    def test_never_overwrites_existing_snapshot(self, tmp_path: Path) -> None:
        source = tmp_path / "Layout"
        source.write_bytes(b"pristine")
        snapshot_once(source)
        source.write_bytes(b"mutated")
        snap = snapshot_once(source)
        assert snap is not None
        assert snap.read_bytes() == b"pristine"

    def test_missing_source_returns_none(self, tmp_path: Path) -> None:
        assert snapshot_once(tmp_path / "absent") is None
