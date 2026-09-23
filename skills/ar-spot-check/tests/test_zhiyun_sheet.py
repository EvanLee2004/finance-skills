import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import zhiyun_probe  # noqa: E402


def test_finance_sheet_is_one_row_per_order():
    customers = [{"name": "甲公司", "sos": ["SO1", "XR2"]}]
    contracts = [
        {"id": "c1", "archive": "20260001", "customers": ["甲公司"]},
        {"id": "c2", "archive": "20260002", "customers": ["甲公司"]},
    ]
    orders = [
        {"so": "SO1", "status": "OP4/项目已交付", "contract_ids": ["c1"], "customers": ["甲公司"]},
        {"so": "XR2", "status": "OP1/项目确认中", "contract_ids": [], "customers": ["甲公司"]},
    ]
    rows = zhiyun_probe.finance_rows(customers, contracts, orders)
    assert [row["订单号"] for row in rows] == ["SO1", "XR2"]
    assert rows[0]["合同归档号"] == "20260001"
    assert rows[0]["订单状态"] == "OP4/项目已交付"
    assert rows[0]["说明"] == ""
    assert rows[1]["合同归档号"] == "未找到"
    assert rows[1]["订单状态"] == "OP1/项目确认中"
    assert "看不出来" not in str(rows)
    assert "20260002" not in str(rows)


def test_missing_order_is_explicit():
    rows = zhiyun_probe.finance_rows([{"name": "乙", "sos": ["SO9"]}], [], [])
    assert rows == [
        {
            "销售": "",
            "客户": "乙",
            "订单号": "SO9",
            "合同归档号": "未找到",
            "订单状态": "未找到",
            "说明": "下单里没有这个单号。",
        }
    ]


def test_order_archive_used_when_contract_has_none():
    rows = zhiyun_probe.finance_rows(
        [{"name": "甲", "sos": ["SO1"]}],
        [],
        [{"so": "SO1", "status": "OP1/项目确认中", "contract_ids": [], "archive": "20261111"}],
    )
    assert rows[0]["合同归档号"] == "20261111"
    assert rows[0]["订单状态"] == "OP1/项目确认中"
    assert rows[0]["说明"] == ""


def test_detail_only_does_not_invent_a_status():
    rows = zhiyun_probe.finance_rows([{"name": "甲", "sos": ["SO8"]}], [], [], {"so8"})
    assert rows[0]["订单状态"] == "未找到"
    assert rows[0]["合同归档号"] == "未找到"
    assert rows[0]["说明"] == "下单里没有这个单号，订单明细里有。"


def test_junk_labels_become_missing():
    rows = zhiyun_probe.finance_rows(
        [{"name": "甲", "sos": ["SO1"]}],
        [{"id": "c1", "archive": "None"}],
        [{"so": "SO1", "status": "看不出来", "contract_ids": ["c1"], "archive": "[]"}],
    )
    assert rows[0]["合同归档号"] == "未找到"
    assert rows[0]["订单状态"] == "未找到"
    assert "看不出来" not in str(rows)
    assert "没挂在这份合同" not in str(rows)


def test_one_order_can_name_two_archives():
    rows = zhiyun_probe.finance_rows(
        [{"name": "甲", "sos": ["SO1"]}],
        [
            {"id": "c1", "archive": "20260001"},
            {"id": "c2", "archive": "20260002"},
            {"id": "c3", "archive": "20269999"},
        ],
        [{"so": "SO1", "status": "OP4/项目已交付", "contract_ids": ["c1", "c2"], "archive": ""}],
    )
    assert rows[0]["订单号"] == "SO1"
    assert rows[0]["合同归档号"] == "20260001、20260002"
    assert "20269999" not in rows[0]["合同归档号"]
    assert rows[0]["说明"] == "这张单挂了多份合同。"
