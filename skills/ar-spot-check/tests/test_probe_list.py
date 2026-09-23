import sys
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import probe_list  # noqa: E402


def test_probe_list_cleans_the_ministry_name(tmp_path):
    sales = tmp_path / "sales.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "销售反馈"
    ws.append(["销售人员", "客户名称", "新智云单号", "交付月份", "账龄", "结算阶段"])
    ws.append(["甲", "中华人民共和国公安部（SO10000001）", "SO10000001", 202508, 13, "未对账"])
    ws.append(["甲", "示例客户", "SO20000001", 202509, 1, "未对账"])
    extra = wb.create_sheet("不是这张")
    extra.append(["无关"])
    wb.save(sales)
    out = tmp_path / "customers.tsv"
    assert probe_list.main(["--sales", str(sales), "--out", str(out)]) == 0
    lines = out.read_text(encoding="utf-8").splitlines()
    assert lines == ["公安部\tSO10000001\t202508", "示例客户\tSO20000001\t202509"]
