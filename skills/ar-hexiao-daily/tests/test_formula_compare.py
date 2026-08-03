from formula_compare import classify_formula_relation, translate_formula


def test_row_shift_formula_is_coordinate_equivalent():
    relation = classify_formula_relation(
        "=VLOOKUP(X5001,$AF$2:$AG$83,2,0)",
        "=VLOOKUP(X5002,$AF$2:$AG$83,2,0)",
        left_coordinate="A5001",
        right_coordinate="A5002",
    )
    assert relation == "coordinate_shift_equivalent"


def test_same_coordinate_wrong_reference_is_logic_difference():
    relation = classify_formula_relation(
        "=N2838+N2838",
        "=N2838+N2839",
        left_coordinate="M2838",
        right_coordinate="M2838",
    )
    assert relation == "logic_difference"


def test_formula_and_constant_are_not_silently_equivalent():
    relation = classify_formula_relation(
        "=N2+N3",
        100,
        left_coordinate="M2",
        right_coordinate="M2",
    )
    assert relation == "formula_value_difference"


def test_translation_uses_real_column_and_preserves_absolute_references():
    formula = "=H7+$H$2+H$3+$H7+SUM(H7:I9)"
    translated = translate_formula(formula, "G7", "H8")
    assert translated == "=I8+$H$2+I$3+$H8+SUM(I8:J10)"


def test_identical_text_at_different_rows_can_still_be_wrong():
    relation = classify_formula_relation(
        "=N8+N9",
        "=N8+N9",
        left_coordinate="M8",
        right_coordinate="M9",
    )
    assert relation == "logic_difference"
