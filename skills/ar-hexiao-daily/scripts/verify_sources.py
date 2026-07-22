#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
源文件只读保证（sha256 清单 + 前后比对）。

思路借鉴李尚 SKILL 仓 `scripts/verify_source_hashes.py`（同项目同需求的另一实现），
本版按本技能的工作区结构与中文交付习惯改写。

为什么要有它：我们对明妹的承诺是"只读不改你的表"。**光在文档里写承诺没用**，
跑完要能拿出机器可验的证据。用法：
    跑之前： python3 scripts/verify_sources.py snapshot
    跑之后： python3 scripts/verify_sources.py verify     # 有任何改动 → 非 0 退出
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import sys
from pathlib import Path
from typing import Dict, List

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common  # noqa: E402

MANIFEST_NAME = "源文件清单.json"
# 需要保证只读的目录：她给的智云导出 + 她的表副本 + 台账（台账我们自己写，不纳入）
WATCH_DIRS = ("01_智云导出", "02_我的表副本")


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def collect(workspace: Path) -> List[dict]:
    out: List[dict] = []
    for d in WATCH_DIRS:
        base = workspace / d
        if not base.is_dir():
            continue
        for p in sorted(base.rglob("*")):
            if p.is_file() and not p.name.startswith("~$"):
                out.append(
                    {
                        "path": str(p),
                        "dir": d,
                        "size": p.stat().st_size,
                        "sha256": sha256(p),
                    }
                )
    return out


def manifest_path(workspace: Path) -> Path:
    d = workspace / "04_产出"
    d.mkdir(parents=True, exist_ok=True)
    return d / MANIFEST_NAME


def do_snapshot(workspace: Path) -> int:
    files = collect(workspace)
    mp = manifest_path(workspace)
    mp.write_text(
        json.dumps(
            {
                "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
                "workspace": str(workspace),
                "files": files,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"已记录 {len(files)} 个源文件指纹 → {mp}")
    if not files:
        print("WARN: 监视目录里没有任何文件（还没放输入？）", file=sys.stderr)
    return 0


def do_verify(workspace: Path) -> int:
    mp = manifest_path(workspace)
    if not mp.is_file():
        print(f"ERROR: 找不到清单 {mp}；请先跑 snapshot", file=sys.stderr)
        return 2
    payload = json.loads(mp.read_text(encoding="utf-8"))
    recorded: Dict[str, dict] = {f["path"]: f for f in payload.get("files", [])}
    problems: List[str] = []
    for path_s, item in recorded.items():
        p = Path(path_s)
        if not p.is_file():
            problems.append(f"不见了: {p.name}")
            continue
        if sha256(p) != item.get("sha256"):
            problems.append(f"被改动: {p.name}")
    now = {f["path"] for f in collect(workspace)}
    for extra in sorted(now - set(recorded)):
        problems.append(f"新出现（未在清单内，可能是运行时误写）: {Path(extra).name}")

    if problems:
        print("源文件只读校验 未通过：", file=sys.stderr)
        for x in problems:
            print(f"  - {x}", file=sys.stderr)
        return 1
    print(f"源文件只读校验通过：{len(recorded)} 个文件与运行前完全一致（未被改动）")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="源文件只读保证（sha256）")
    ap.add_argument("action", choices=["snapshot", "verify"])
    ap.add_argument("--workspace", default=str(common.WORK))
    args = ap.parse_args(argv)
    ws = Path(args.workspace)
    return do_snapshot(ws) if args.action == "snapshot" else do_verify(ws)


if __name__ == "__main__":
    sys.exit(main())
