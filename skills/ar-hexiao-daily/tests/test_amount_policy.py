from decimal import Decimal

import amount_policy as P


def test_technical_and_business_tolerances_are_separate():
    assert P.technical_equal(100, 100.005)
    assert not P.technical_equal(100, 100.006)
    assert P.within_business_tolerance(100, 101.00)
    assert not P.within_business_tolerance(100, 101.01)


def test_business_comparison_keeps_exact_delta():
    result = P.business_comparison("100.00", "99.72")
    assert result["delta"] == Decimal("0.28")
    assert result["technical_equal"] is False
    assert result["business_equal"] is True
