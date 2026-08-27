from __future__ import annotations

import datetime as dt
import json
from concurrent.futures import ThreadPoolExecutor

import atomic_file_publish as PUBLISH
import atomic_json_store as STORE
import batch_ledger
import fallback_allocation_ledger
import pytest


def _counter_default() -> dict:
    return {"count": 0}


def _validate_counter(value: dict) -> None:
    if not isinstance(value, dict) or not isinstance(value.get("count"), int):
        raise TypeError("invalid counter")


def test_corrupt_existing_ledgers_fail_closed(tmp_path) -> None:
    batch_ledger.ledger_path(tmp_path).write_text("{truncated", encoding="utf-8")
    with pytest.raises(STORE.JsonStoreCorruptionError, match="不会按空台账继续"):
        batch_ledger.load(tmp_path)

    allocation_path = fallback_allocation_ledger.ledger_path(tmp_path)
    allocation_path.write_text("[]", encoding="utf-8")
    with pytest.raises(STORE.JsonStoreCorruptionError, match="不会按空台账继续"):
        fallback_allocation_ledger.load(tmp_path)


def test_atomic_update_preserves_original_when_publish_fails(tmp_path, monkeypatch) -> None:
    path = tmp_path / "ledger.json"
    path.write_text('{"count": 1}', encoding="utf-8")
    original_replace = STORE.os.replace

    def fail_target_publish(source, target):
        if STORE.Path(target) == path:
            raise OSError("simulated publish interruption")
        return original_replace(source, target)

    monkeypatch.setattr(STORE.os, "replace", fail_target_publish)
    with pytest.raises(OSError, match="simulated publish interruption"):
        STORE.update_json(
            path,
            default_factory=_counter_default,
            validate=_validate_counter,
            mutate=lambda value: value.update(count=2),
        )

    assert json.loads(path.read_text(encoding="utf-8")) == {"count": 1}
    assert not list(tmp_path.glob(".ledger.json.*.tmp"))


def test_atomic_update_serializes_concurrent_writers_and_keeps_backup(tmp_path) -> None:
    path = tmp_path / "ledger.json"

    def increment(_index: int) -> None:
        STORE.update_json(
            path,
            default_factory=_counter_default,
            validate=_validate_counter,
            mutate=lambda value: value.update(count=value["count"] + 1),
        )

    with ThreadPoolExecutor(max_workers=6) as pool:
        list(pool.map(increment, range(18)))

    assert STORE.load_json(
        path,
        default_factory=_counter_default,
        validate=_validate_counter,
    ) == {"count": 18}
    backup = path.with_suffix(".json.bak")
    assert json.loads(backup.read_text(encoding="utf-8")) == {"count": 17}


def test_batch_record_remains_monotonic_under_atomic_store(tmp_path) -> None:
    day = dt.date(2026, 8, 17)
    batch_ledger.record(tmp_path, day, "applied", payments=2)
    batch_ledger.record(tmp_path, day, "classified", payments=2)

    state = batch_ledger.load(tmp_path)
    assert state["runs"][day.isoformat()]["stage"] == "applied"
    assert batch_ledger.ledger_path(tmp_path).with_suffix(".json.bak").is_file()


def test_multi_file_publish_rolls_back_when_second_replace_fails(tmp_path, monkeypatch) -> None:
    first = tmp_path / "first.xlsx"
    second = tmp_path / "second.xlsx"
    staged_first = tmp_path / "staged-first.xlsx"
    staged_second = tmp_path / "staged-second.xlsx"
    first.write_bytes(b"old-first")
    second.write_bytes(b"old-second")
    staged_first.write_bytes(b"new-first")
    staged_second.write_bytes(b"new-second")
    original_replace = PUBLISH.os.replace
    failed = False

    def fail_second(source, target):
        nonlocal failed
        if PUBLISH.Path(target) == second and not failed:
            failed = True
            raise OSError("simulated second publish failure")
        return original_replace(source, target)

    monkeypatch.setattr(PUBLISH.os, "replace", fail_second)
    with pytest.raises(OSError, match="second publish failure"):
        PUBLISH.publish(
            {first: staged_first, second: staged_second},
            tmp_path / "transaction",
        )

    assert first.read_bytes() == b"old-first"
    assert second.read_bytes() == b"old-second"
    assert not (tmp_path / "transaction").exists()
