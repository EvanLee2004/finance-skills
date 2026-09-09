#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 Gitee finance-skills 白名单装进本机 opencode skills。核销跟仓库 main。"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

SKIP_SUFFIXES = {".pyc"}


def load_pack(path: Path | None = None) -> dict[str, Any]:
    cfg = path or Path(__file__).resolve().parent.parent / "config" / "pack.json"
    return json.loads(cfg.read_text(encoding="utf-8"))


def primary_remote(pack: dict[str, Any]) -> str:
    return str(pack["gitee"])


def default_dest() -> Path:
    override = os.environ.get("OPENCODE_SKILLS_DIR", "").strip()
    if override:
        return Path(override).expanduser()
    return Path.home() / ".config" / "opencode" / "skills"


def _should_skip(name: str, pack: dict[str, Any]) -> bool:
    if name in pack.get("skip_names", []):
        return True
    return Path(name).suffix in SKIP_SUFFIXES


def _copy_tree(src: Path, dest: Path, pack: dict[str, Any], *, keep_config: bool) -> None:
    dest.mkdir(parents=True, exist_ok=True)
    for item in src.iterdir():
        if _should_skip(item.name, pack):
            continue
        if keep_config and item.name == "config":
            continue
        target = dest / item.name
        if item.is_dir():
            if target.exists():
                shutil.rmtree(target)
            shutil.copytree(
                item,
                target,
                ignore=shutil.ignore_patterns("__pycache__", ".pyc", "*.pyc"),
            )
        else:
            shutil.copy2(item, target)


def apply_update(src_skills: Path, dest_skills: Path, pack: dict[str, Any]) -> dict[str, list[str]]:
    report: dict[str, list[str]] = {
        "updated": [],
        "installed": [],
        "skipped_protected": [],
        "missing_in_source": [],
        "config_kept": [],
        "root_files": [],
        "removed": [],
    }
    dest_skills.mkdir(parents=True, exist_ok=True)
    protected = set(pack.get("protected_from_overwrite") or [])
    whitelist = set(pack["whitelist"])
    for skill_id in pack.get("remove_retired") or []:
        if not skill_id or skill_id in whitelist:
            continue
        dest_skill = dest_skills / skill_id
        if dest_skill.is_dir():
            shutil.rmtree(dest_skill)
            report["removed"].append(skill_id)

    for skill_id in pack["whitelist"]:
        src_skill = src_skills / skill_id
        dest_skill = dest_skills / skill_id
        if not src_skill.is_dir():
            report["missing_in_source"].append(skill_id)
            continue
        if skill_id in protected and dest_skill.exists():
            report["skipped_protected"].append(skill_id)
            continue
        existed = dest_skill.exists()
        overwrite_config = set(pack.get("overwrite_config") or [])
        keep_config = dest_skill.joinpath("config").exists() and skill_id not in overwrite_config
        _copy_tree(src_skill, dest_skill, pack, keep_config=keep_config)
        if keep_config:
            report["config_kept"].append(skill_id)
        if existed:
            report["updated"].append(skill_id)
        else:
            report["installed"].append(skill_id)

    for name in pack.get("root_files") or []:
        src_file = src_skills / name
        if src_file.is_file():
            shutil.copy2(src_file, dest_skills / name)
            report["root_files"].append(name)
    return report


def fetch_source(cache_dir: Path, pack: dict[str, Any]) -> tuple[Path, str, str]:
    """拉 Gitee main；不通再试 GitHub。返回 (skills 目录, short SHA, remote 名)。"""
    cache_dir.mkdir(parents=True, exist_ok=True)
    repo = cache_dir / "finance-skills"
    branch = str(pack.get("branch") or "main")
    remotes = [("gitee", primary_remote(pack)), ("github", str(pack["github"]))]
    last_err = ""
    for label, url in remotes:
        try:
            if (repo / ".git").exists():
                subprocess.run(
                    ["git", "-C", str(repo), "remote", "set-url", "origin", url],
                    check=True,
                    capture_output=True,
                    text=True,
                )
                subprocess.run(
                    ["git", "-C", str(repo), "fetch", "--depth", "1", "origin", branch],
                    check=True,
                    capture_output=True,
                    text=True,
                )
                subprocess.run(
                    ["git", "-C", str(repo), "checkout", "-f", f"origin/{branch}"],
                    check=True,
                    capture_output=True,
                    text=True,
                )
            else:
                if repo.exists():
                    shutil.rmtree(repo)
                subprocess.run(
                    ["git", "clone", "--depth", "1", "-b", branch, url, str(repo)],
                    check=True,
                    capture_output=True,
                    text=True,
                )
            sha = subprocess.check_output(
                ["git", "-C", str(repo), "rev-parse", "--short", "HEAD"],
                text=True,
            ).strip()
            return repo / "skills", sha, label
        except (subprocess.CalledProcessError, FileNotFoundError) as exc:
            last_err = str(exc)
            continue
    raise RuntimeError(f"Gitee 和 GitHub 都拉不下来：{last_err}")


def format_report(report: dict[str, list[str]], sha: str, remote: str, dest: Path) -> str:
    lines = [
        f"更新到 {remote} {sha}",
        f"装到 {dest}",
        f"更新：{', '.join(report['updated']) or '无'}",
        f"新装：{', '.join(report['installed']) or '无'}",
        f"跳过（保护名单）：{', '.join(report['skipped_protected']) or '无'}",
        f"本地 config 保留：{', '.join(report['config_kept']) or '无'}",
        f"已下线并删除：{', '.join(report.get('removed') or []) or '无'}",
        "白名单外其他技能：未动",
        "请重启 opencode。",
    ]
    if report["missing_in_source"]:
        lines.append(f"仓里没有：{', '.join(report['missing_in_source'])}")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="更新财务部官方 skills（Gitee 优先；核销跟 main）")
    parser.add_argument("--dest", type=Path, default=None, help="opencode skills 目录")
    parser.add_argument("--source", type=Path, default=None, help="已有仓内 skills/ 目录；不传则拉 Gitee")
    parser.add_argument("--cache", type=Path, default=None, help="浅克隆缓存目录")
    args = parser.parse_args(argv)

    pack = load_pack()
    dest = args.dest or default_dest()
    if args.source:
        src = args.source
        sha = "local"
        remote = "local"
    else:
        cache = args.cache or Path.home() / ".cache" / "finance-skills-update"
        src, sha, remote = fetch_source(cache, pack)
    report = apply_update(src, dest, pack)
    print(format_report(report, sha, remote, dest))
    return 0


if __name__ == "__main__":
    sys.exit(main())
