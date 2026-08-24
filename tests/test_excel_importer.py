from datetime import date

from app.services.excel_importer import (
    _date_value,
    _integrated_item_type,
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
    assert _item_type("오므론", 1) == "부가제품"
    assert _item_type("오므론", -1) == "부가제품"
    assert _item_type("배송비", 1) == "배송비"
    assert _item_type("배송비", -1) == "배송비"
    assert _item_type("사이즈미정", -1) == "사이즈미정"
    assert _item_type("본품취소", -1) == "제품취소"
    assert _item_type("rev2_10", 1) == "제품"


def test_integrated_item_type_prioritizes_special_product_options():
    assert _integrated_item_type("제품취소", "배송비", 1) == "배송비"
    assert _integrated_item_type("제품취소", "오므론", -1) == "부가제품"
    assert _integrated_item_type("제품취소", "사이즈미정", -1) == "사이즈미정"
    assert _integrated_item_type("제품취소", "rev2_10", -1) == "제품취소"
    assert _integrated_item_type("제품", "rev1_12", -1) == "제품취소"
    assert _integrated_item_type("제품확정", "rev2_10", -1) == "제품취소"
    assert _integrated_item_type("제품확정", "rev2_10", 1) == "제품확정"
