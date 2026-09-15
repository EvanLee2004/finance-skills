"""整 SO 结账状态；调用方先按交付年度选定盈亏表。"""
from __future__ import annotations

SO_ALREADY_SETTLED = "OK_SO_ALREADY_SETTLED"


def inspect_so(so: str, rows) -> dict:
    """rows 使用 row/so/sod/settled 四个字段，包含拆分业务行。"""
    so = str(so or "").strip()
    matches = sorted(
        (dict(row) for row in rows if so and str(row.get("so") or "").strip() == so),
        key=lambda row: row["row"],
    )
    open_rows = [row["row"] for row in matches if str(row.get("settled") or "").strip() != "是"]
    sods = sorted({str(row.get("sod") or "").strip() for row in matches if str(row.get("sod") or "").strip()})
    complete = bool(matches) and not open_rows
    refs = [row["row"] for row in matches]
    if complete:
        reason = (
            f"SO={so} 在对应年度盈亏表共有 {len(matches)} 条业务行（含拆分行），"
            f"{len(sods)} 个已登记 SOD，是否结账全部为是；整单跳过，保留历史金额、计提、日期和方式。"
            f"核查行：{','.join(map(str, refs))}。此结论依据整单结账状态，不表示本批回款已逐笔写入。"
        )
    elif matches:
        reason = f"SO={so} 尚有是否结账非是的业务行：{','.join(map(str, open_rows))}；不能整单跳过。"
    else:
        reason = f"SO={so} 在对应年度盈亏表没有业务行；不能以空集合认定已结账。"
    return {"all_settled": complete, "so": so, "rows": matches, "open_rows": open_rows, "sods": sods, "reason": reason}
