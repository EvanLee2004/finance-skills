"""Fixed-order parent allocations through the existing reconciliation interfaces."""
import copy
import os
import tempfile
import openpyxl
import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import classify_hexiao as C
import validate_plan as V
import fallback_allocation_ledger as FAL
import apply_to_copy as A


def payment(ar, amount, day="2026-09-04"):
    return dict(ar=ar, amount_orig=amount, amount_local=amount, currency="人民币CNY",
        arrival_date=day, hexiao_date="2026-09-07", huikuan_type="分笔回款",
        orders=[dict(so="SO_TEST", deliver=100, currency="人民币CNY", delivery_date="2026-05-01")],
        sod_lines={"SO_TEST": [dict(sod="SOD_TEST", deliver=100)]})


def plan(payments):
    records = C.expand_payments(payments)
    ledger = C.LedgerIndex(synthetic={"so": {"SO_TEST": [5]}, "sod": {"SOD_TEST": [5]},
        "rows": {5: dict(so="SO_TEST", sod="SOD_TEST", yingshou=100, huikuan=None, jiezhang="否")}})
    result = C.classify_records_by_year(records, {2026: ledger})
    result["hexiao_date"] = "2026-09-07"
    result["parent_fallback_allocations"] = {p["ar"]: p["_parent_fallback_allocation"] for p in payments if p.get("_parent_fallback_allocation")}
    return result


def rows():
    return {5: {"SO": "SO_TEST", "SOD": "SOD_TEST", "应收金额": 100, "计提": None,
        "回款明细": None, "是否结账": "否", "收款时间": None, "收款方式": None,
        "实收SOD": "SOD_TEST", "差异": None}}


class SequenceTests(unittest.TestCase):
    def test_shared_outstanding_is_consumed_in_fixed_order(self):
        payments = [payment("AR_B", 80), payment("AR_A", 60)]
        C.expand_payments(payments)
        by_ar = {p["ar"]: p["_parent_fallback_allocation"] for p in payments}
        self.assertEqual(by_ar["AR_A"]["allocations"][0]["allocated_local"], 60)
        self.assertEqual(by_ar["AR_B"]["allocations"][0]["allocated_local"], 40)
        self.assertEqual(by_ar["AR_B"]["unallocated_parent_amount"], 40)

    def test_chain_classifies_and_validates(self):
        result = plan([payment("AR_B", 80), payment("AR_A", 60)])
        self.assertEqual(len(result["auto"]), 2, result)
        checked = V.validate_by_year(result, {2026: rows()})
        self.assertEqual(checked["counts"], {"write": 2, "skip": 0, "conflict": 0}, checked)
        self.assertEqual([s["current_received"] for s in checked["write"][0]["row_operation"]["steps"]], [60, 40])

    def test_zero_allocation_is_checked_and_preserved(self):
        ps = [payment("AR_B", 80), payment("AR_A", 100)]
        result = plan(ps)
        checked = V.validate_by_year(result, {2026: rows()})
        self.assertEqual(checked["counts"], {"write": 1, "skip": 1, "conflict": 0}, checked)
        saved = FAL.eligible_entries(checked)
        self.assertIn("AR_B", saved)
        self.assertEqual(saved["AR_B"]["unallocated_parent_amount"], 80)

    def test_validation_failure_blocks_dependent_zero(self):
        result = plan([payment("AR_B", 80), payment("AR_A", 100)])
        bad = rows(); bad[5]["SOD"] = "SOD_CHANGED"
        checked = V.validate_by_year(result, {2026: bad})
        self.assertEqual(checked["counts"], {"write": 0, "skip": 0, "conflict": 2}, checked)

    def test_arrival_date_precedes_ar_and_input_order(self):
        ps = [payment("AR_A", 80, "2026-09-05"), payment("AR_Z", 60, "2026-09-04")]
        C.expand_payments(ps)
        self.assertEqual(ps[0]["_parent_fallback_allocation"]["allocations"][0]["allocated_local"], 40)
        original = copy.deepcopy(ps[0]["_parent_fallback_allocation"])
        C.expand_payments(list(reversed(ps)))
        self.assertEqual(ps[0]["_parent_fallback_allocation"], original)

    def test_missing_arrival_date_cannot_choose_ar_only(self):
        with self.assertRaises(ValueError):
            C.expand_payments([payment("AR_A", 60, ""), payment("AR_B", 80)])

    def test_explicit_order_amounts_are_not_reallocated(self):
        p = payment("AR_A", 80)
        p.update(writeoffs={"SO_TEST": 35}, writeoffs_local={"SO_TEST": 35})
        records = C.expand_payments([p])
        self.assertEqual(records[0]["amount_local"], 35)
        self.assertNotIn("_parent_fallback_allocation", p)

    def test_multiple_sods_share_capacity_and_validate_as_one_component(self):
        ps = [payment("AR_A", 60), payment("AR_B", 40)]
        for p in ps:
            p["sod_lines"] = {"SO_TEST": [dict(sod="SOD_A", deliver=50), dict(sod="SOD_B", deliver=50)]}
        records = C.expand_payments(ps)
        ledger = C.LedgerIndex(synthetic={"so": {"SO_TEST": [5, 6]}, "sod": {"SOD_A": [5], "SOD_B": [6]},
            "rows": {ref: dict(so="SO_TEST", sod=sod, yingshou=50, huikuan=None, jiezhang="否") for ref, sod in [(5, "SOD_A"), (6, "SOD_B")]}})
        result = C.classify_records_by_year(records, {2026: ledger})
        self.assertEqual(result["hold"], [], result["hold"])
        self.assertEqual(sorted((r["ar"], r["sod"], r["five_cols"]["回款明细"]) for r in result["auto"]),
            [("AR_A", "SOD_A", 50), ("AR_A", "SOD_B", 10), ("AR_B", "SOD_B", 40)])
        actual = {ref: {**rows()[5], "SOD": sod, "应收金额": 50} for ref, sod in [(5, "SOD_A"), (6, "SOD_B")]}
        checked = V.validate_by_year(result, {2026: actual})
        self.assertEqual(checked["counts"], {"write": 3, "skip": 0, "conflict": 0}, checked["conflict"])
        actual[5]["SOD"] = "CHANGED"
        checked = V.validate_by_year(result, {2026: actual})
        self.assertEqual(checked["counts"], {"write": 0, "skip": 0, "conflict": 3})

    def test_voided_receipt_cannot_be_accepted_as_zero(self):
        second = payment("AR_B", 80); second["status"] = "已作废"
        result = plan([payment("AR_A", 100), second])
        self.assertEqual(result["auto"], [])
        self.assertTrue(any(r["ar"] == "AR_B" and r["code"] == "E7" for r in result["exception"]))

    def test_saved_zero_does_not_accept_replaced_or_missing_order(self):
        checked = V.validate_by_year(plan([payment("AR_A", 100), payment("AR_B", 80)]), {2026: rows()})
        saved = FAL.eligible_entries(checked)
        p = payment("AR_B", 80); p["_fallback_allocation_state"] = {"parents": saved}
        rs = C.expand_payments([p])
        for ledger in [C.LedgerIndex(synthetic={}), C.LedgerIndex(synthetic={"so": {"SO_TEST": [5]},
            "sod": {"SOD_TEST": [5]}, "rows": {5: dict(so="SO_TEST", sod="SOD_TEST", yingshou=100, huikuan=None, jiezhang="否")}})]:
            result = C.classify_records_by_year(rs, {2026: ledger})
            self.assertEqual(result["auto"], [])
            self.assertEqual(result["hold"][0]["code"], "E_FALLBACK_ZERO_BASELINE")

    def test_missing_related_year_blocks_the_component(self):
        first = payment("AR_A", 60)
        first["orders"].append(dict(so="SO_OLD", deliver=200, currency="人民币CNY", delivery_date="2025-05-01"))
        first["amount_orig"] = first["amount_local"] = 110
        first["sod_lines"]["SO_OLD"] = [dict(sod="SOD_OLD", deliver=200)]
        result = plan([first, payment("AR_B", 80)])
        self.assertEqual(result["auto"], [])

    def test_saved_allocation_cannot_move_on_rerun(self):
        checked = V.validate_by_year(plan([payment("AR_A", 60), payment("AR_B", 80)]), {2026: rows()})
        saved = FAL.eligible_entries(checked)
        ps = [payment("AR_B", 80), payment("AR_A", 60)]
        for p in ps:
            p["_fallback_allocation_state"] = {"parents": saved}
        C.expand_payments(ps)
        self.assertEqual(ps[0]["_parent_fallback_allocation"]["allocations"][0]["allocated_local"], 40)
        self.assertEqual(ps[1]["_parent_fallback_allocation"]["allocations"][0]["allocated_local"], 60)
        self.assertTrue(all(p["_parent_fallback_allocation"]["reused_successful_allocation"] for p in ps))

    def test_prior_classification_failure_blocks_entire_component(self):
        first = payment("AR_A", 60)
        first["sod_lines"] = {}
        result = plan([first, payment("AR_B", 80)])
        self.assertEqual(result["auto"], [])

    @unittest.skipUnless(os.environ.get("LOCAL_SYNTHETIC_TEST_DIR"), "workbook artifacts are local only")
    def test_write_readback_and_replay(self):
        with tempfile.TemporaryDirectory(dir=os.environ["LOCAL_SYNTHETIC_TEST_DIR"]) as folder:
            root = Path(folder)
            src, out = root / "source.xlsx", root / "result.xlsx"
            wb = openpyxl.Workbook(); sheet = wb.active; sheet.title = "明细"
            sheet.append(["部门", "销售人员", "客户名称", "单号", "新智云单号", "应收金额", "计提金额", "回款明细", "是否结账（是/否）", "收款时间", "收款方式(支/汇/现)", "实收金额", "差异"])
            sheet.append(["测试", "测试", "测试", "", "SO_TEST", 100, None, None, "否", None, None, "SOD_TEST", None])
            wb.save(src); wb.close()
            for amounts in [(60, 80), (100, 80)]:
                ps = [payment("AR_A", amounts[0]), payment("AR_B", amounts[1])]
                rs = C.expand_payments(ps)
                result = C.classify_records_by_year(rs, {2026: C.LedgerIndex(src)}, ledger_paths={2026: src})
                result.update(hexiao_date="2026-09-07", parent_fallback_allocations={p["ar"]: p["_parent_fallback_allocation"] for p in ps})
                checked = V.validate_by_year(result, {2026: V.read_ledger_rows(src)}, {2026: src})
                self.assertEqual(checked["conflict"], [])
                self.assertEqual(A.precheck_before_write(checked, checked["write"], src), [])
                A.write_plan(src, out, checked["write"])
                self.assertEqual(A.verify_written(out, checked["write"]), [])
                state_dir = root / str(amounts[0]); state_dir.mkdir()
                FAL.commit(state_dir, checked)
                saved = FAL.load(state_dir)
                replay = [payment("AR_B", amounts[1]), payment("AR_A", amounts[0])]
                for p in replay:
                    p["_fallback_allocation_state"] = saved
                rs2 = C.expand_payments(replay)
                result2 = C.classify_records_by_year(rs2, {2026: C.LedgerIndex(out)}, ledger_paths={2026: out})
                result2.update(hexiao_date="2026-09-07", parent_fallback_allocations={p["ar"]: p["_parent_fallback_allocation"] for p in replay})
                checked2 = V.validate_by_year(result2, {2026: V.read_ledger_rows(out)}, {2026: out})
                self.assertEqual(checked2["counts"], {"write": 0, "skip": 2, "conflict": 0}, checked2["counts"])
                FAL.commit(state_dir, checked2)

if __name__ == "__main__":
    unittest.main()
