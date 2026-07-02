"""Atomic file-write helpers shared by every on-disk writer.

Pattern (same guarantees as the historical layout writer):

1. Serialise to a sibling temp file (``<name>.tmp.<pid>``).
2. Copy the existing target (if any) to ``<name>.bak`` so a previous
   known-good version is recoverable after a crash mid-write.
3. ``os.replace`` the temp onto the target — atomic on Windows + POSIX.

On any exception during steps 1/2, the original file is untouched.
"""

from __future__ import annotations

import logging
import os
from pathlib import Path

logger = logging.getLogger("atomic_io")

BACKUP_SUFFIX = ".bak"
SNAPSHOT_SUFFIX = ".orig"


def atomic_write_bytes(path: Path | str, data: bytes, *, backup: bool = True) -> Path | None:
    """Atomically replace ``path`` with ``data``.

    Returns the backup path when a backup of the previous content was
    written, else ``None``. A failed backup is logged and does not block
    the (still atomic) write.
    """
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = target.with_name(f"{target.name}.tmp.{os.getpid()}")
    bak_path: Path | None = None
    try:
        tmp_path.write_bytes(data)
        if backup and target.exists():
            candidate = target.with_name(f"{target.name}{BACKUP_SUFFIX}")
            try:
                candidate.write_bytes(target.read_bytes())
                bak_path = candidate
            except OSError as exc:
                logger.warning("Backup of %s failed (%s); proceeding with atomic write.", target, exc)
        os.replace(tmp_path, target)
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass
    return bak_path


def atomic_write_text(
    path: Path | str,
    text: str,
    *,
    encoding: str = "utf-8",
    backup: bool = True,
) -> Path | None:
    """Atomically replace ``path`` with ``text`` (no newline translation)."""
    return atomic_write_bytes(path, text.encode(encoding), backup=backup)


def snapshot_once(path: Path | str, *, suffix: str = SNAPSHOT_SUFFIX) -> Path | None:
    """Copy ``path`` to ``<name><suffix>`` unless that snapshot already exists.

    Unlike the rolling ``.bak`` written on every save, the snapshot
    preserves the pristine pre-operation content across multi-round
    loops (e.g. ``pbi_repair_loop``) so a clean rollback stays possible.
    Returns the snapshot path (existing or newly written), or ``None``
    when the source does not exist or the copy failed.
    """
    source = Path(path)
    if not source.exists():
        return None
    snapshot = source.with_name(f"{source.name}{suffix}")
    if snapshot.exists():
        return snapshot
    try:
        snapshot.write_bytes(source.read_bytes())
    except OSError as exc:
        logger.warning("Pristine snapshot of %s failed: %s", source, exc)
        return None
    return snapshot


__all__ = ["BACKUP_SUFFIX", "SNAPSHOT_SUFFIX", "atomic_write_bytes", "atomic_write_text", "snapshot_once"]
