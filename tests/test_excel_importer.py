from datetime import date

from app.services.excel_importer import (
    _date_value,
    _item_type,
    _number,
    _sales_month,
)


def test_number_normalization():
    assert _number("1,250,000") == 1_250_000
    assert _number(None) == 0


def test_month_normalization():
    assert _sales_month("7월", None) == 7
    assert _sales_month("", date(2026, 8, 14)) == 8


def test_excel_date_normalization():
    assert _date_value(46207) == "2026-07-04"


def test_item_type_classification():
    assert _item_type("오므론", 1) == "부가상품"
    assert _item_type("배송비", 1) == "배송비"
    assert _item_type("본품취소", -1) == "취소"

