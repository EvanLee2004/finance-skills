"""Zero-allocation settled orders retain delivery for the flow plan."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "vendor" / "scripts"))
import classify_hexiao as classify
import build_flow_plan as flow


class SettledDeliveryTests(unittest.TestCase):
    def expand(self, currency="人民币CNY", rate=1, amount=334.53):
        payment = dict(
            ar="AR_TEST", currency=currency, amount_orig=100, amount_local=100,
            hexiao_date="2026-09-07",
            orders=[dict(so="SO_TEST", deliver=amount, deliver_local=None,
                         currency=currency, rate=rate)],
            _ledger_settled_sos=["SO_TEST"],
        )
        return classify.expand_payment(payment, {})

    def test_cny_delivery_reaches_flow_without_allocating_receipt(self):
        record, = self.expand()
        self.assertEqual(record["forced_code"], "E_SETTLED_SO_RECHECK")
        self.assertEqual(record["amount_orig"], 0)
        self.assertEqual(record["amount_local"], 0)
        self.assertEqual(flow._delivery_amount(record), 334.53)
        self.assertEqual(record["deliver_local"], 334.53)

    def test_foreign_delivery_uses_order_rate(self):
        record, = self.expand("美元USD", 7)
        self.assertEqual(flow._delivery_amount(record), 2341.71)
        self.assertEqual(record["amount_local"], 0)

    def test_missing_foreign_rate_still_blocks_flow(self):
        record, = self.expand("美元USD", None)
        self.assertIsNone(flow._delivery_amount(record))
        self.assertEqual(record["amount_orig"], 0)

    def test_missing_delivery_stays_missing(self):
        records = self.expand(amount=None)
        self.assertTrue(records)
        self.assertTrue(all(flow._delivery_amount(r) is None for r in records))


if __name__ == "__main__":
    unittest.main()
