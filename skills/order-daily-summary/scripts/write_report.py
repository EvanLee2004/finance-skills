"""写出亮晶版式 xlsx：sheet「下单数据」+「处理日志」；可选明细。"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any

from summarize import DEPT_DISPLAY_ORDER, SummaryResult, dept_totals, row_totals


def write_report(
    result: SummaryResult,
    out_path: Path | str,
    *,
    window_start: date | str | None = None,
    window_end: date | str | None = None,
    api_row_count: int | None = None,
    include_detail: bool = False,
    data_asof: datetime | None = None,
    late_warning: str | None = None,
    extra_log: dict[str, Any] | None = None,
) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    asof = data_asof or datetime.now()
    asof_text = (
        f"数据截至 {asof.strftime('%Y-%m-%d %H:%M:%S')}"
        "（智云为实时数据，昨日订单当天可能被改期，本表为该时刻快照）"
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "下单数据"

    headers = ["日期", "总计", *DEPT_DISPLAY_ORDER]
    ws.append(headers)
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill

    dates = sorted(result.by_date.keys())
    totals = row_totals(result)
    for d in dates:
        depts = result.by_date[d]
        row = [d, totals.get(d, 0.0)]
        for col in DEPT_DISPLAY_ORDER:
            row.append(depts.get(col, 0.0))
        ws.append(row)

    # number format
    for r in range(2, ws.max_row + 1):
        for c in range(2, len(headers) + 1):
            ws.cell(row=r, column=c).number_format = "#,##0.00"

    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws.freeze_panes = "A2"

    # 数据截至 + 晚跑提示：直接落在数据下方，亮晶打开就看得到
    ws.append([])
    asof_row = ws.max_row + 1
    ws.cell(row=asof_row, column=1, value=asof_text).font = Font(italic=True, color="808080")
    if late_warning:
        warn_row = asof_row + 1
        wc = ws.cell(row=warn_row, column=1, value=f"⚠ {late_warning}")
        wc.font = Font(bold=True, color="C00000")
        wc.alignment = Alignment(wrap_text=False, vertical="center")

    # 处理日志
    log = wb.create_sheet("处理日志")
    log.append(["项目", "内容"])
    log["A1"].font = Font(bold=True)
    log["B1"].font = Font(bold=True)
    log.append(["处理时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    log.append(["数据截至", asof.strftime("%Y-%m-%d %H:%M:%S")])
    log.append(["快照说明", "智云为实时数据，昨日订单当天可能被改期，本表为该时刻快照"])
    if late_warning:
        log.append(["⚠晚跑提示", late_warning])
    log.append(
        [
            "日期窗口",
            f"{window_start or ''}～{window_end or ''}".strip("～") or "（未指定）",
        ]
    )
    log.append(["接口行数", api_row_count if api_row_count is not None else result.detail_row_count])
    log.append(["明细行数", result.detail_row_count])
    log.append(["总计万元", result.grand_total_wan])
    log.cell(row=log.max_row, column=2).number_format = "#,##0.00"  # 不依赖行号，插行也不错位
    log.append(["金额字段", result.amount_field_used or "（未解析到）"])
    dtot = dept_totals(result)
    for name in DEPT_DISPLAY_ORDER:
        log.append([f"分部门万元·{name}", dtot.get(name, 0.0)])
        log.cell(row=log.max_row, column=2).number_format = "#,##0.00"
    unmatched = "、".join(result.unmatched_sales) if result.unmatched_sales else "无"
    log.append(["未匹配销售数量", len(result.unmatched_sales)])
    log.append(["未匹配销售名单", unmatched])
    if extra_log:
        for k, v in extra_log.items():
            log.append([str(k), v])
    log.column_dimensions["A"].width = 22
    log.column_dimensions["B"].width = 60

    if include_detail and result.detail_rows:
        detail = wb.create_sheet("明细")
        dheaders = ["销售", "下单日期", "金额本币", "金额万元", "归类", "金额字段"]
        detail.append(dheaders)
        for cell in detail[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
        for r in result.detail_rows:
            detail.append([r.get(h, "") for h in dheaders])
        for col in range(1, len(dheaders) + 1):
            detail.column_dimensions[get_column_letter(col)].width = 14

    wb.save(out_path)
    return out_path
