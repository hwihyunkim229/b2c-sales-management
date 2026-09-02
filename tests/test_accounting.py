from app.main import build_accounting_summary


def accounting_row(**overrides):
    row = {
        "total_sales": 0,
        "ancillary_sales": 0,
        "delivery_sales": 0,
        "unspecified_sales": 0,
        "product_quantity": 0,
        "confirmed_quantity": 0,
        "cancel_quantity": 0,
        "credit_card_amount": 0,
        "cash_receipt_amount": 0,
        "tax_invoice_amount": 0,
        "other_amount": 0,
    }
    row.update(overrides)
    return row


def test_product_sale_balances_supply_price_and_vat():
    result = build_accounting_summary(accounting_row(
        total_sales=550_000,
        product_quantity=1,
        credit_card_amount=550_000,
    ))
    assert result["debit_total"] == 550_000
    assert result["credit_total"] == 550_000
    assert result["difference"] == 0


def test_ancillary_and_delivery_are_negative_debits():
    result = build_accounting_summary(accounting_row(
        total_sales=22_000,
        ancillary_sales=11_000,
        delivery_sales=11_000,
        other_amount=22_000,
    ))
    assert result["debit_rows"][1]["amount"] == -10_000
    assert result["debit_rows"][2]["amount"] == -10_000
    assert result["difference"] == 0


def test_confirmed_product_posts_to_both_advance_and_product_sales():
    result = build_accounting_summary(accounting_row(confirmed_quantity=1))
    assert result["debit_rows"][3]["amount"] == 500_000
    assert result["credit_rows"][1]["amount"] == 500_000
    assert result["difference"] == 0


def test_cancelled_product_keeps_negative_quantity_and_vat():
    result = build_accounting_summary(accounting_row(
        total_sales=-550_000,
        cancel_quantity=-1,
        credit_card_amount=-550_000,
    ))
    assert result["credit_rows"][2]["amount"] == -500_000
    assert result["payment_rows"][0]["vat_amount"] == -50_000
    assert result["difference"] == 0


def test_negative_unspecified_sale_becomes_positive_advance():
    result = build_accounting_summary(accounting_row(
        total_sales=-550_000,
        unspecified_sales=-550_000,
        cash_receipt_amount=-550_000,
    ))
    assert result["debit_rows"][4]["amount"] == 500_000
    assert result["difference"] == 0
