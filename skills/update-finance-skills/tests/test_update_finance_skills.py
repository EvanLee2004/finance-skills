#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""更新财务skills：白名单覆盖、保留普通技能 config、核销跟仓库。"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
SCRIPTS = SKILL / "scripts"
sys.path.insert(0, str(SCRIPTS))

import update as upd  # noqa: E402


def _write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _fake_pack(tmp: Path) -> tuple[Path, dict, Path]:
    src = tmp / "src" / "skills"
    dest = tmp / "dest"
    src.mkdir(parents=True)
    dest.mkdir()
    pack = {
        "gitee": "https://gitee.com/Lee157/finance-skills.git",
        "github": "https://github.com/EvanLee2004/finance-skills.git",
        "branch": "main",
        "whitelist": [
            "receivables-merge",
            "ar-hexiao-daily",
            "update-finance-skills",
        ],
        "protected_from_overwrite": [],
        "overwrite_config": ["ar-hexiao-daily"],
        "root_files": ["财务技能包_来源与更新.md", "财务技能_说什么用哪个.md"],
    }
    for sid in pack["whitelist"]:
        _write(src / sid / "SKILL.md", f"cloud {sid}\n")
        _write(src / sid / "scripts" / "run.py", "print(1)\n")
        _write(src / sid / "config" / "rules.md", "cloud-config\n")
    _write(src / "财务技能包_来源与更新.md", "source note\n")
    _write(src / "财务技能_说什么用哪个.md", "router\n")
    return src, pack, dest


def test_pack_json_lists_gitee_and_hexiao_on_main():
    pack = upd.load_pack()
    assert pack["gitee"] == "https://gitee.com/Lee157/finance-skills.git"
    assert pack["branch"] == "main"
    assert "ar-hexiao-daily" not in pack.get("protected_from_overwrite", [])
    assert "ar-hexiao-daily" in pack.get("overwrite_config", [])
    assert "update-finance-skills" in pack["whitelist"]
    assert "ar-hexiao-daily" in pack["whitelist"]


def test_updates_existing_hexiao_including_config(tmp_path: Path):
    src, pack, dest = _fake_pack(tmp_path)
    _write(dest / "ar-hexiao-daily" / "SKILL.md", "old-hexiao\n")
    _write(dest / "ar-hexiao-daily" / "scripts" / "old.py", "old\n")
    _write(dest / "ar-hexiao-daily" / "config" / "rules.md", "old-config\n")

    report = upd.apply_update(src, dest, pack)

    assert "ar-hexiao-daily" in report["updated"]
    assert "ar-hexiao-daily" not in report["skipped_protected"]
    assert (dest / "ar-hexiao-daily" / "SKILL.md").read_text(encoding="utf-8") == "cloud ar-hexiao-daily\n"
    assert (dest / "ar-hexiao-daily" / "config" / "rules.md").read_text(encoding="utf-8") == "cloud-config\n"


def test_installs_hexiao_only_when_missing(tmp_path: Path):
    src, pack, dest = _fake_pack(tmp_path)

    report = upd.apply_update(src, dest, pack)

    assert "ar-hexiao-daily" in report["installed"]
    assert "ar-hexiao-daily" not in report["skipped_protected"]
    assert (dest / "ar-hexiao-daily" / "SKILL.md").read_text(encoding="utf-8") == "cloud ar-hexiao-daily\n"


def test_updates_whitelist_and_keeps_local_config(tmp_path: Path):
    src, pack, dest = _fake_pack(tmp_path)
    _write(dest / "receivables-merge" / "SKILL.md", "old-skill\n")
    _write(dest / "receivables-merge" / "config" / "rules.md", "local-config\n")

    report = upd.apply_update(src, dest, pack)

    assert "receivables-merge" in report["updated"]
    assert "receivables-merge" in report["config_kept"]
    assert (dest / "receivables-merge" / "SKILL.md").read_text(encoding="utf-8") == "cloud receivables-merge\n"
    assert (dest / "receivables-merge" / "config" / "rules.md").read_text(encoding="utf-8") == "local-config\n"


def test_does_not_touch_colleague_extra_skill(tmp_path: Path):
    src, pack, dest = _fake_pack(tmp_path)
    _write(dest / "my-own-skill" / "SKILL.md", "keep-me\n")

    upd.apply_update(src, dest, pack)

    assert (dest / "my-own-skill" / "SKILL.md").read_text(encoding="utf-8") == "keep-me\n"


def test_copies_router_files_to_skills_root(tmp_path: Path):
    src, pack, dest = _fake_pack(tmp_path)

    report = upd.apply_update(src, dest, pack)

    assert "财务技能_说什么用哪个.md" in report["root_files"]
    assert (dest / "财务技能_说什么用哪个.md").read_text(encoding="utf-8") == "router\n"


def test_prefers_gitee_url():
    pack = upd.load_pack()
    assert upd.primary_remote(pack).startswith("https://gitee.com/")


def test_skill_description_mentions_update_phrases():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    head = text.split("---", 2)[1]
    for phrase in ("更新财务skills", "更新财务技能", "安装财务skills"):
        assert phrase in head, phrase


def test_router_card_covers_spoken_jobs():
    card = (SKILL.parent / "财务技能_说什么用哪个.md").read_text(encoding="utf-8")
    for phrase in ("更新财务skills", "跑本周应收", "跑昨天的核销", "琪哥发票入金蝶", "九点下单"):
        assert phrase in card, phrase
