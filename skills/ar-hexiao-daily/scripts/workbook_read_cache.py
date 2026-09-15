"""Lab-only row cache, authenticated by the fingerprint held by the Worker.

No pickle, no shared user cache, no cached reads in actual write verification.
Every hit also hashes the current workbook bytes, so post-write files miss.
"""
from __future__ import annotations

import copy
import datetime as dt
import gzip
import hashlib
import io
import json
from pathlib import Path

SCHEMA = "ar-read-rows-v1"
MAX_BYTES = 256 * 1024 * 1024
_entries = {}
_root = None


def digest(path):
    with Path(path).open("rb") as handle:
        return hashlib.file_digest(handle, "sha256").hexdigest()


def encode(value):
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return {"type": type(value).__name__, "value": value.isoformat()}
    if isinstance(value, dt.timedelta):
        return {"type": "timedelta", "days": value.days, "seconds": value.seconds, "microseconds": value.microseconds}
    raise TypeError("Unsupported workbook cell value")


def decode(value):
    kind = value.get("type")
    if kind in {"datetime", "date", "time"}:
        return getattr(dt, kind).fromisoformat(value["value"])
    if kind == "timedelta":
        return dt.timedelta(days=value["days"], seconds=value["seconds"], microseconds=value["microseconds"])
    return value


def configure(path: Path, fingerprint: str, root: Path):
    global _entries, _root
    root = root.resolve()
    if path.is_symlink() or not path.resolve().is_relative_to(root):
        raise ValueError("缓存不属于当前任务")
    with path.open("rb") as handle:
        raw = handle.read(MAX_BYTES + 1)
    if len(raw) > MAX_BYTES or hashlib.sha256(raw).hexdigest() != fingerprint:
        raise ValueError("缓存指纹与任务固定记录不一致")
    with gzip.GzipFile(fileobj=io.BytesIO(raw)) as stream:
        unpacked = stream.read(MAX_BYTES + 1)
    if len(unpacked) > MAX_BYTES:
        raise ValueError("缓存超过读取上限")
    payload = json.loads(unpacked, object_hook=decode)
    if payload.get("schema") != SCHEMA or payload.get("root") != str(root):
        raise ValueError("缓存版本或任务归属不一致")
    _entries, _root = payload["entries"], root


def read_rows(path: Path, sheet: str | None = None, *, worksheet=None):
    import openpyxl

    path = Path(path)
    if _root is not None and (path.is_symlink() or not path.resolve().is_relative_to(_root)):
        raise ValueError("工作簿不属于当前任务")
    sheet_key = sheet or "@active"
    has_candidate = _root is not None and any(key.endswith(f":{sheet_key}") for key in _entries)
    key = f"{digest(path)}:{sheet_key}" if has_candidate else ""
    if key in _entries:
        return [tuple(row) for row in copy.deepcopy(_entries[key])]
    if worksheet is not None:
        return list(worksheet.iter_rows(values_only=True))
    workbook = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    try:
        if sheet is not None and sheet not in workbook.sheetnames:
            raise ValueError(f"工作簿缺少 {sheet} 工作表")
        worksheet = workbook[sheet] if sheet else workbook.active
        return list(worksheet.iter_rows(values_only=True))
    finally:
        workbook.close()


def build(root: Path, files: list[tuple[Path, str | None]], output: Path):
    root = root.resolve()
    entries = {}
    for path, sheet in files:
        path = path.resolve()
        if not path.is_relative_to(root):
            raise ValueError("输入超出任务目录")
        before = digest(path)
        key = f"{before}:{sheet or '@active'}"
        if key not in entries:
            entries[key] = read_rows(path, sheet)
        if digest(path) != before:
            raise ValueError("读取索引期间工作簿发生变化")
    raw = json.dumps({"schema": SCHEMA, "root": str(root), "entries": entries}, default=encode,
                     ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    if len(raw) > MAX_BYTES:
        raise ValueError("工作簿索引超过大小上限")
    if output.is_symlink() or not output.resolve().is_relative_to(root):
        raise ValueError("缓存目标超出任务目录")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("xb") as handle:
        handle.write(gzip.compress(raw, compresslevel=1, mtime=0))


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--ledger", action="append", default=[])
    parser.add_argument("--export", action="append", default=[])
    args = parser.parse_args()
    build(Path(args.root), [(Path(path), "明细") for path in args.ledger] +
          [(Path(path), None) for path in args.export], Path(args.output))
    print("只读工作簿索引已生成")
