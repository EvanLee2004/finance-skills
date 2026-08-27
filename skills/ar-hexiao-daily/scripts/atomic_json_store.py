"""Crash-safe JSON persistence for reconciliation ledgers.

The public interface deliberately stays small: ``load_json`` for validated
reads and ``update_json`` for a locked read-modify-write transaction.  A
damaged existing file is never treated as an empty ledger.
"""

from __future__ import annotations

import copy
import json
import os
import shutil
import tempfile
import time
from collections.abc import Callable
from contextlib import contextmanager
from pathlib import Path
from typing import TypeVar

T = TypeVar("T")
R = TypeVar("R")


class JsonStoreError(RuntimeError):
    """Base class for durable JSON store failures."""


class JsonStoreCorruptionError(JsonStoreError):
    """An existing JSON file could not be parsed or validated."""


class JsonStoreLockTimeout(JsonStoreError):
    """Another process held the ledger lock for too long."""


def _validate_or_raise(path: Path, value: T, validate: Callable[[T], None]) -> T:
    try:
        validate(value)
    except (TypeError, ValueError, KeyError) as exc:
        raise JsonStoreCorruptionError(
            f"台账文件格式无效，已停止处理且不会按空台账继续：{path.name}"
        ) from exc
    return value


def load_json(
    path: Path,
    *,
    default_factory: Callable[[], T],
    validate: Callable[[T], None],
) -> T:
    """Load and validate JSON, returning a fresh default only when absent."""
    path = Path(path)
    if not path.exists():
        return _validate_or_raise(path, default_factory(), validate)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise JsonStoreCorruptionError(
            f"台账文件无法读取，已停止处理且不会按空台账继续：{path.name}"
        ) from exc
    return _validate_or_raise(path, value, validate)


def _flush_directory(path: Path) -> None:
    """Best-effort directory flush; Windows does not expose a portable fd."""
    if os.name == "nt":
        return
    descriptor = os.open(str(path), os.O_RDONLY)
    try:
        os.fsync(descriptor)
    finally:
        os.close(descriptor)


def _copy_backup(path: Path) -> None:
    if not path.is_file():
        return
    backup = path.with_suffix(path.suffix + ".bak")
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{backup.name}.", suffix=".tmp", dir=path.parent
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        shutil.copy2(path, temporary)
        with temporary.open("r+b") as handle:
            os.fsync(handle.fileno())
        os.replace(temporary, backup)
    finally:
        temporary.unlink(missing_ok=True)


def _write_json_atomic(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        _copy_backup(path)
        os.replace(temporary, path)
        _flush_directory(path.parent)
    finally:
        temporary.unlink(missing_ok=True)


@contextmanager
def _exclusive_lock(path: Path, timeout_seconds: float = 10.0):
    lock_path = path.with_suffix(path.suffix + ".lock")
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    handle = lock_path.open("a+b")
    if handle.tell() == 0:
        handle.write(b"0")
        handle.flush()
    deadline = time.monotonic() + max(float(timeout_seconds), 0.0)
    acquired = False
    try:
        while not acquired:
            try:
                handle.seek(0)
                if os.name == "nt":
                    import msvcrt

                    msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl

                    fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                acquired = True
            except OSError as exc:
                if time.monotonic() >= deadline:
                    raise JsonStoreLockTimeout(
                        f"等待台账写锁超时，请确认没有其它核销任务正在更新：{path.name}"
                    ) from exc
                time.sleep(0.05)
        yield
    finally:
        if acquired:
            handle.seek(0)
            if os.name == "nt":
                import msvcrt

                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl

                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
        handle.close()


def update_json(
    path: Path,
    *,
    default_factory: Callable[[], T],
    validate: Callable[[T], None],
    mutate: Callable[[T], R],
    lock_timeout_seconds: float = 10.0,
) -> R:
    """Run one locked, validated and crash-safe read-modify-write transaction."""
    path = Path(path)
    with _exclusive_lock(path, lock_timeout_seconds):
        current = load_json(path, default_factory=default_factory, validate=validate)
        working = copy.deepcopy(current)
        result = mutate(working)
        _validate_or_raise(path, working, validate)
        _write_json_atomic(path, working)
        return result
