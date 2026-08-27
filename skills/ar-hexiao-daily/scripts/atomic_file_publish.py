"""Rollback-journalled publication of a verified set of files."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import uuid
from collections.abc import Mapping
from pathlib import Path


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _write_manifest(path: Path, payload: dict) -> None:
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        with temporary.open("w", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def recover(transaction_dir: Path) -> bool:
    """Roll back an interrupted publication; return whether recovery occurred."""
    transaction_dir = Path(transaction_dir).resolve()
    manifest_path = transaction_dir / "manifest.json"
    if not transaction_dir.exists():
        return False
    if not manifest_path.is_file():
        shutil.rmtree(transaction_dir)
        return False
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("state") == "completed":
        shutil.rmtree(transaction_dir)
        return False
    if manifest.get("state") != "prepared":
        raise RuntimeError("多年度发布事务状态无效，拒绝猜测恢复。")
    for entry in manifest.get("files") or []:
        target = Path(str(entry.get("target") or "")).resolve()
        if entry.get("existed"):
            backup = (transaction_dir / str(entry.get("backup") or "")).resolve()
            if not backup.is_file() or not backup.is_relative_to(transaction_dir):
                raise RuntimeError("多年度发布事务缺少恢复副本，拒绝继续。")
            temporary = target.with_name(f".{target.name}.rollback.tmp")
            shutil.copy2(backup, temporary)
            os.replace(temporary, target)
        else:
            target.unlink(missing_ok=True)
    shutil.rmtree(transaction_dir)
    return True


def publish(replacements: Mapping[Path, Path], transaction_dir: Path) -> None:
    """Publish all staged files or restore every original on failure."""
    transaction_dir = Path(transaction_dir).resolve()
    recover(transaction_dir)
    normalized = [
        (Path(target).resolve(), Path(staged).resolve())
        for target, staged in replacements.items()
    ]
    if not normalized:
        return
    for _target, staged in normalized:
        if not staged.is_file():
            raise FileNotFoundError(f"待发布文件不存在：{staged}")
    transaction_dir.mkdir(parents=True, exist_ok=False)
    entries = []
    try:
        for index, (target, staged) in enumerate(normalized):
            entry = {
                "target": str(target),
                "sha256": _sha256(staged),
                "existed": target.is_file(),
                "backup": "",
            }
            if target.is_file():
                backup_relative = Path("backups") / f"{index:05d}.bak"
                backup = transaction_dir / backup_relative
                backup.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(target, backup)
                entry["backup"] = str(backup_relative)
            entries.append(entry)
        manifest = {"state": "prepared", "files": entries}
        _write_manifest(transaction_dir / "manifest.json", manifest)
        for (target, staged), entry in zip(normalized, entries, strict=True):
            target.parent.mkdir(parents=True, exist_ok=True)
            temporary = target.with_name(f".{target.name}.publish.tmp")
            try:
                shutil.copy2(staged, temporary)
                if _sha256(temporary) != entry["sha256"]:
                    raise RuntimeError(f"多年度文件发布校验失败：{target.name}")
                os.replace(temporary, target)
            finally:
                temporary.unlink(missing_ok=True)
        manifest["state"] = "completed"
        _write_manifest(transaction_dir / "manifest.json", manifest)
    except Exception:
        recover(transaction_dir)
        raise
    else:
        shutil.rmtree(transaction_dir)
