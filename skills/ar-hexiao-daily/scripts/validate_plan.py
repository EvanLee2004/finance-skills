#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第 7 步 a：**写入前校验计划**（官方 plan → validate → execute 模式的中间环）。

为什么要有这一步：判定结果是"打算怎么填"，但从判定到写入之间，她的表可能已经变了
（月初贴交付会插行、部分核销会在上方插行 → **行号立刻失效**）。
直接按行号写 = 把值写到别人家的行上。所以写之前逐条复核，过不了的不写、并说清为什么。

用法：
    python3 scripts/validate_plan.py --plan 判定结果.json --ledger 盈亏副本.xlsx \
        --out 04_产出/写入计划_校验后.json
退出码：0=有可写的（或全跳过）；2=输入不可用；**1=存在冲突**（有笔需要人看）
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import common  # noqa: E402
import writeoff_duplicate_audit as WDA  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

FIVE = ["计提", "回款明细", "是否结账", "收款时间", "收款方式"]
DERIVED = ["差异"]
VALID_JIEZHANG = {"是", "否"}
VALID_WAY = {"汇", "冲预收", "支", "现"}


def duplicate_audit_error(plan: dict) -> str:
    """防止重复审计或逻辑记录在判定JSON到校验之间被手工改坏。"""
    if "duplicate_writeoff_audits" not in plan:
        return ""  # 兼容纯单测/旧夹具；当前版本四件套会始终带该字段
    audits = plan.get("duplicate_writeoff_audits") or {}
    expected = str(plan.get("duplicate_writeoff_audit_sha256") or "")
    actual = WDA.audit_fingerprint(audits)
    if not expected or expected != actual:
        return "系统重复核销审计指纹不一致，计划可能被修改，必须重新判定"
    for item in plan.get("auto") or []:
        ar = str(item.get("ar") or "")
        audit = audits.get(ar) or {}
        status = audit.get("status")
        warnings = set(item.get("warning_codes") or [])
        if status == "unresolved":
            return f"父回款 {ar} 的重复审计未解决，却进入了auto"
        if status == "recovered":
            if "W_SYSTEM_DUPLICATE_WRITEOFF_COLLAPSED" not in warnings:
                return f"父回款 {ar} 已做系统重复纠正，但auto缺少警告码"
            if item.get("duplicate_writeoff_audit") != audit:
                return f"父回款 {ar} 的auto行与顶层重复审计不一致"
    return ""


def _norm(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    # Excel 里 300 与 300.0 是同一个数
    try:
        f = float(s)
        return f"{f:.2f}"
    except (TypeError, ValueError):
        return s


def read_ledger_rows(path: Path) -> Dict[int, dict]:
    """把盈亏『明细』整表读成 {行号: {列名: 值}}，用于逐条复核。"""
    import openpyxl

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    if "明细" not in wb.sheetnames:
        wb.close()
        raise ValueError(f"盈亏表无『明细』sheet：{wb.sheetnames}")
    ws = wb["明细"]
    aliases = common.load_aliases()
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hrow, headers = common.find_header_row(
        all_rows, "盈亏明细", ["SO", "SOD", "计提", "回款明细", "是否结账"], aliases
    )
    cols = common.resolve_columns(
        headers,
        "盈亏明细",
        ["SO", "SOD", "计提", "回款明细", "是否结账", "收款时间", "收款方式"],
        aliases,
    )
    diff_idx = common.fuzzy_find_col(
        headers, (aliases.get("盈亏明细", {}) or {}).get("差异", ["差异"])
    )
    if diff_idx is not None:
        cols["差异"] = diff_idx
    yidx = common.fuzzy_find_col(
        headers, (aliases.get("盈亏明细", {}) or {}).get("应收", ["应收金额", "应收"])
    )
    if yidx is None:
        raise ValueError("盈亏『明细』找不到应收金额列，无法校验部分回款拆行")
    cols["应收"] = yidx
    out: Dict[int, dict] = {}
    for i, row in enumerate(all_rows, start=1):
        if i <= hrow + 1:
            continue
        vals = list(row)

        def cell(key):
            idx = cols.get(key)
            return vals[idx] if idx is not None and idx < len(vals) else None

        out[i] = {
            "SO": str(cell("SO") or "").strip(),
            "SOD": str(cell("SOD") or "").strip(),
            "计提": cell("计提"),
            "回款明细": cell("回款明细"),
            "差异": cell("差异"),
            "_差异列存在": "差异" in cols,
            "是否结账": cell("是否结账"),
            "收款时间": cell("收款时间"),
            "收款方式": cell("收款方式"),
            "应收金额": cell("应收"),
        }
    return out


def _matches_identity(row: Optional[dict], so: str, sod: str) -> bool:
    if row is None:
        return False
    if so and row.get("SO") != so:
        return False
    if sod and row.get("SOD") != sod:
        return False
    return bool(so or sod)


def _matches_planned_fields(row: dict, expected: dict) -> bool:
    """现有业务行是否已等于计划目标；用于插行后的唯一重定位与幂等复核。"""
    for key in FIVE:
        if _norm(row.get(key)) != _norm(expected.get(key)):
            return False
    expected_sod = str(expected.get("实收SOD") or "").strip()
    if expected_sod and row.get("SOD") != expected_sod:
        return False
    return True


def resolve_item_row(item: dict, rows: Dict[int, dict]) -> tuple[Optional[int], str]:
    """
    优先使用判定时行号；受控插行使行号失效时，按 SO/SOD 业务身份唯一重定位。

    不允许仅按 SO 猜行。同一 SO/SOD 有多行时，只有其中恰有一行已等于计划目标，
    才可用于写后幂等复核；否则继续冲突。
    """
    ref = item.get("ledger_row_ref")
    if not ref:
        return None, "判定结果里没有行号，无法定位"
    ref = int(ref)
    so = str(item.get("so") or "").strip()
    sod = str(item.get("sod") or "").strip()
    if _matches_identity(rows.get(ref), so, sod):
        return ref, ""

    candidates = [
        row_no for row_no, row in rows.items()
        if _matches_identity(row, so, sod)
    ]
    if len(candidates) == 1:
        return candidates[0], ""
    target_matches = [
        row_no for row_no in candidates
        if _matches_planned_fields(rows[row_no], item.get("five_cols") or {})
    ]
    if len(target_matches) == 1:
        return target_matches[0], ""
    if not candidates:
        return None, f"按 SO/SOD 找不到原计划第 {ref} 行对应的业务行"
    return None, (
        f"按 SO/SOD 找到 {len(candidates)} 行，无法唯一重定位"
        f"（候选行={candidates[:8]}）"
    )


def check_one(item: dict, rows: Dict[int, dict]) -> dict:
    """
    单条复核 → {verdict: write|skip|conflict, reason}
    - write    ：行号对得上、目标格是空的、值合法 → 可以写
    - skip     ：已经填过且与计划一致 → 幂等跳过（重复跑不重复写）
    - conflict ：行号对不上 / 已填但不一致 / 值不合法 → 不写，交给人看
    """
    ref = item.get("ledger_row_ref")
    five = item.get("five_cols") or {}
    derived = item.get("derived_cols") or {}
    so, sod = (item.get("so") or "").strip(), (item.get("sod") or "").strip()

    if not ref:
        return {"verdict": "conflict", "reason": "判定结果里没有行号，无法定位"}
    row = rows.get(int(ref))
    if row is None:
        return {"verdict": "conflict", "reason": f"第 {ref} 行在表里不存在了（表被删过行？）"}

    # ① 行号还指着同一单吗——她插过行的话这里必然对不上
    if sod and row["SOD"] and row["SOD"] != sod:
        return {
            "verdict": "conflict",
            "reason": f"第 {ref} 行现在是 {row['SOD']}，不是计划里的 {sod}（表在判定之后被插过行）",
        }
    if so and row["SO"] and row["SO"] != so:
        return {
            "verdict": "conflict",
            "reason": f"第 {ref} 行现在是 {row['SO']}，不是计划里的 {so}（表被改过）",
        }

    # ② 值本身合法吗
    if five.get("是否结账") not in VALID_JIEZHANG:
        return {"verdict": "conflict", "reason": f"是否结账取值异常：{five.get('是否结账')!r}"}
    if five.get("收款方式") not in VALID_WAY:
        return {"verdict": "conflict", "reason": f"收款方式取值异常：{five.get('收款方式')!r}"}
    for k in ("计提", "回款明细"):
        v = five.get(k)
        if v is None:
            continue  # 部分核销时计提本就留空
        try:
            float(v)
        except (TypeError, ValueError):
            return {"verdict": "conflict", "reason": f"{k} 不是数字：{v!r}"}
    for k in DERIVED:
        if k not in derived:
            continue
        if not row.get("_差异列存在"):
            return {"verdict": "conflict", "reason": "本次需要写差异，但盈亏明细没有“差异”列"}
        try:
            float(derived[k])
        except (TypeError, ValueError):
            return {"verdict": "conflict", "reason": f"{k} 不是数字：{derived[k]!r}"}

    op = item.get("row_operation") or {}
    if op:
        if op.get("type") != "split_below":
            return {"verdict": "conflict", "reason": f"未知行操作：{op.get('type')!r}"}
        if row.get("_差异列存在") and _norm(row.get("差异")) not in ("", "None"):
            return {
                "verdict": "conflict",
                "reason": (
                    "部分回款阶段计提和业务值差异都必须留空，"
                    f"但当前差异={_norm(row.get('差异'))!r}；禁止覆盖"
                ),
            }
        try:
            source = round(float(op["source_receivable"]), 2)
            paid = round(float(op["paid_receivable"]), 2)
            unpaid = round(float(op["unpaid_receivable"]), 2)
            latest = round(float(op["latest_delivery"]), 2)
            cumulative = round(float(op["cumulative_received"]), 2)
        except (KeyError, TypeError, ValueError):
            return {"verdict": "conflict", "reason": "部分回款拆行参数缺失或不是数字"}
        if paid < 0 or unpaid <= 0:
            return {"verdict": "conflict", "reason": f"拆行应收异常：已收侧={paid} 未收侧={unpaid}"}
        if abs((paid + unpaid) - source) > 0.011:
            return {"verdict": "conflict", "reason": f"拆行不守恒：{paid}+{unpaid}!={source}"}
        if abs((latest - cumulative) - unpaid) > 0.011:
            return {"verdict": "conflict", "reason": f"未回款公式不成立：{latest}-{cumulative}!={unpaid}"}
        inserted = op.get("inserted_five_cols") or {}
        if inserted.get("是否结账") != "否":
            return {"verdict": "conflict", "reason": "拆出的未回款行必须是否结账=否"}
        for key in ("计提", "回款明细", "收款时间", "收款方式"):
            if inserted.get(key) is not None:
                return {"verdict": "conflict", "reason": f"未回款行 {key} 必须留空"}

        # 写后重跑：原行已变成已收侧、下一行已是未收侧时，识别为完整幂等状态。
        # 不能继续拿拆前 source_receivable 要求当前行，否则受控拆行必然假报冲突。
        current_receivable = common.to_number(row.get("应收金额"))
        if (
            current_receivable is not None
            and abs(float(current_receivable) - paid) <= 0.011
        ):
            next_row = rows.get(int(ref) + 1)
            next_receivable = (
                common.to_number(next_row.get("应收金额"))
                if next_row is not None else None
            )
            same_next_identity = _matches_identity(next_row, so, sod)
            paid_matches = _matches_planned_fields(row, five)
            unpaid_matches = (
                next_row is not None
                and _matches_planned_fields(next_row, inserted)
                and (
                    not next_row.get("_差异列存在")
                    or _norm(next_row.get("差异")) in ("", "None")
                )
            )
            if (
                same_next_identity
                and next_receivable is not None
                and abs(float(next_receivable) - unpaid) <= 0.011
                and paid_matches
                and unpaid_matches
            ):
                return {
                    "verdict": "skip",
                    "reason": "部分回款拆行已完整写入（已收行+紧邻未收行），幂等跳过",
                }
            return {
                "verdict": "conflict",
                "reason": "当前行看似已拆分，但已收行或紧邻未收行与计划不一致",
            }
        if current_receivable is None or abs(float(current_receivable) - source) > 0.011:
            return {
                "verdict": "conflict",
                "reason": f"当前行应收已变化：表里={current_receivable} 计划基线={source}",
            }

    # ③ 目标格现在是什么
    # ⚠「是否结账」是她盈亏表**预置的默认值**：单子交付了、钱还没到的行默认就是「否」
    #   （2026-07-24 真实全表统计：「否」+回款空 = 1973 行，全是还没收款的挂账行）。
    #   所以它**不能当"这一行她已经填过"的证据**——否则每一笔新到账（表里那行本就带着「否」）
    #   都会被误判成"已填过、且和计划不一致"→ 冲突，导致自动回填对真实新数据全线失效
    #   （2026-07-24 opencode 真实 24 号数据实测：4 笔可填全被「否」挡成冲突、可写=0）。
    #   判"填没填过"只看**回款证据列**（计提/回款明细/收款时间/收款方式）；是否结账留给下面 same/diff
    #   去比对（她填了钱却没翻「是」这种真不一致，仍会照常报冲突）。
    evidence_cols = [k for k in FIVE if k != "是否结账"]
    filled = [k for k in evidence_cols if _norm(row.get(k)) not in ("", "None")]
    if not filled:
        for k in DERIVED:
            if k in derived and _norm(row.get(k)) not in ("", "None", _norm(derived[k])):
                return {
                    "verdict": "conflict",
                    "reason": f"{k}: 表里={_norm(row.get(k))!r} 计划={_norm(derived[k])!r}",
                }
        return {"verdict": "write", "reason": "回款列为空，可写（是否结账为她表预置默认，不计入已填证据）"}

    # 这行她已经填过了 → 逐列比对。**两种不一致都算不一致**：
    #   ① 计划有值 ≠ 表里的值
    #   ② 计划算的是「留空」，但表里填了值 —— 2026-07-25 修：旧版 `if five.get(k) is not None`
    #      把这种整个跳过了，结果「表里计提=1000、本次算的是留空」会被判成"已填过且一致·跳过"，
    #      静默放过。而「计提到底该不该填」恰恰是明妹口径里最容易出错的一条
    #      （回款明细合计 = 交付额才可填计提），漏报等于把最该她看的那行藏起来。
    diff: List[str] = []
    for k in FIVE:
        want, got = five.get(k), _norm(row.get(k))
        if want is None:
            if got not in ("", "None"):
                diff.append(f"{k}: 表里={got!r} 本次算的是**留空**")
            continue
        if got != _norm(want):
            diff.append(f"{k}: 表里={got!r} 计划={_norm(want)!r}")
    derived_missing: List[str] = []
    for k in DERIVED:
        if k not in derived:
            continue
        want, got = derived[k], _norm(row.get(k))
        if got in ("", "None"):
            derived_missing.append(k)
        elif got != _norm(want):
            diff.append(f"{k}: 表里={got!r} 计划={_norm(want)!r}")
    if diff:
        return {
            "verdict": "conflict",
            "reason": "这行已经填过，且和本次算的不一样 → " + "；".join(diff),
        }
    if derived_missing:
        return {
            "verdict": "write",
            "reason": "五项回款字段已一致，仅补业务值差异公式：" + "、".join(derived_missing),
        }
    if not diff:
        return {"verdict": "skip", "reason": "已经填过且与本次一致（幂等跳过）"}
    raise AssertionError("不可达")


def validate(
    plan: dict,
    rows: Dict[int, dict],
    ledger_path: Optional[Path] = None,
) -> dict:
    items = [dict(it) for it in (plan.get("auto") or [])]
    checked: List[dict] = []
    seen_rows: Dict[int, str] = {}
    audit_error = duplicate_audit_error(plan)
    for it in items:
        original_ref = it.get("ledger_row_ref")
        if audit_error:
            res = {"verdict": "conflict", "reason": audit_error}
        else:
            resolved_ref, locate_error = resolve_item_row(it, rows)
            if locate_error:
                res = {"verdict": "conflict", "reason": locate_error}
            else:
                if int(resolved_ref) != int(original_ref):
                    it["_relocated_from"] = int(original_ref)
                    it["ledger_row_ref"] = int(resolved_ref)
                res = check_one(it, rows)
        ref = it.get("ledger_row_ref")
        # 同一行被两条计划命中 → 都不写（谁对谁错要人定）
        if res["verdict"] == "write" and ref in seen_rows:
            res = {
                "verdict": "conflict",
                "reason": f"第 {ref} 行被两笔计划同时命中（另一笔 {seen_rows[ref]}），需人工指定",
            }
        elif res["verdict"] == "write":
            seen_rows[ref] = it.get("case_id") or ""
        item = {**it, "_check": res}
        # 校验这一刻该行的身份，写进计划带给 apply。
        # apply 拿它跟**写入那一刻**的表再对一次：对不上说明中间被插过行/删过行。
        cur = rows.get(int(ref)) if ref else None
        if cur is not None:
            item["_identity"] = {"row": int(ref), "SO": cur["SO"], "SOD": cur["SOD"]}
        checked.append(item)
    buckets = {"write": [], "skip": [], "conflict": []}
    for c in checked:
        buckets[c["_check"]["verdict"]].append(c)
    out = {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "hexiao_date": plan.get("hexiao_date") or "",
        "counts": {k: len(v) for k, v in buckets.items()},
        "selection": {
            "mode": "all_auto",
            "selected": len(items),
            "total_auto": len(items),
        },
        "duplicate_writeoff_audits": plan.get("duplicate_writeoff_audits") or {},
        "duplicate_writeoff_audit_sha256": plan.get("duplicate_writeoff_audit_sha256") or "",
        **buckets,
    }
    if ledger_path is not None:
        # 盈亏副本在「校验」这一刻的指纹。apply 前会再算一次比对：
        # 不一致 = 她在"看清单 → 说确认"这段时间动过表 → 拒写，让她重跑一遍（几十秒的事）。
        out["ledger_path"] = str(ledger_path)
        out["ledger_sha256"] = common.sha256_file(ledger_path)
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="写入前校验计划（plan→validate→execute）")
    # --plan / --ledger 都可不给：不给就去工作区自己找（脏活归程序，别让 AI/她填路径）
    ap.add_argument("--plan", default="", help="判定结果 json；不给则取 04_产出 最新")
    ap.add_argument("--ledger", default="", help="盈亏核算表副本（只读）；不给则取 02_我的表副本/*盈亏*")
    ap.add_argument("--out", default="", help="校验后计划 json")
    # 防呆：同上。--workspace 还用于在没给 --out 时把结果落进正确的 04_产出/
    ap.add_argument("--workspace", default="", help="工作区根（没给 --out 时用它定产出位置）")
    ap.add_argument("--hexiao-date", default="", help="（校验日期以判定结果为准，收下防止链路中断）")
    args = ap.parse_args(argv)

    ws = common.resolve_workspace(args.workspace or None)
    out_dir = ws / "04_产出"

    def _latest(pattern: str):
        c = [p for p in sorted(out_dir.glob(pattern)) if not p.name.startswith("~$")]
        return c[-1] if c else None

    plan_p = Path(args.plan) if args.plan else (_latest("判定结果_*.json") or Path(""))
    if args.ledger:
        ledger_p = Path(args.ledger)
    else:
        cand = [
            p for p in sorted((ws / "02_我的表副本").glob("*盈亏*"))
            if not p.name.startswith(("~$", "."))
        ] if (ws / "02_我的表副本").is_dir() else []
        ledger_p = cand[0] if cand else Path("")

    if not plan_p.is_file():
        print(
            f"ERROR: 找不到判定结果{f' {plan_p}' if args.plan else f'（{out_dir} 里没有 判定结果_*.json）'}"
            "\n  先跑 classify_hexiao.py",
            file=sys.stderr,
        )
        return 2
    if not ledger_p.is_file():
        print(
            f"ERROR: 找不到盈亏表{f' {ledger_p}' if args.ledger else f'（{ws}/02_我的表副本/ 里没有 *盈亏* 文件）'}",
            file=sys.stderr,
        )
        return 2

    plan = json.loads(plan_p.read_text(encoding="utf-8"))
    try:
        rows = read_ledger_rows(ledger_p)
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    result = validate(plan, rows, ledger_path=ledger_p)
    # 没给 --out 就落进**解析后的工作区**的 04_产出/，别落到 plan 旁边（会跟日清分家）
    out_p = Path(args.out) if args.out else (out_dir / "写入计划_校验后.json")
    out_p.parent.mkdir(parents=True, exist_ok=True)
    out_p.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

    c = result["counts"]
    if result.get("hexiao_date"):
        print(f"核销日期：{common.date_cn(result['hexiao_date'])}")
    print(f"校验完成 可写={c['write']} 幂等跳过={c['skip']} 冲突={c['conflict']}")
    for x in result["conflict"][:10]:
        print(f"  ⚠ {x.get('case_id')}: {x['_check']['reason']}")
    if c["conflict"] > 10:
        print(f"  …另有 {c['conflict']-10} 条冲突，详见 {out_p.name}")
    print(f"计划: {out_p}")
    return 1 if c["conflict"] else 0


if __name__ == "__main__":
    sys.exit(main())
