from dataclasses import dataclass

@dataclass(frozen=True)
class ChannelAdapter:
    platform: str
    sheet_names: tuple[str, ...]
    header_aliases: dict[str, tuple[str, ...]]

COMMON_ALIASES = {
    "order_number": ("주문번호", "상품주문번호", "주문 번호"),
    "tax_date": ("결제일", "주문일", "매출일", "구매확정일"),
    "product_option": ("옵션정보", "옵션명", "상품옵션", "반지사이즈"),
    "quantity": ("수량", "판매수량", "구매수량"),
    "buyer_name": ("구매자", "수취인명", "주문자명"),
}

ADAPTERS = (
    ChannelAdapter("네이버", ("네이버", "Naver"), COMMON_ALIASES),
    ChannelAdapter("쿠팡", ("쿠팡", "Coupang"), COMMON_ALIASES),
    ChannelAdapter("카페24", ("카페24", "Cafe24"), COMMON_ALIASES),
    ChannelAdapter("11번가", ("11번가",), COMMON_ALIASES),
    ChannelAdapter("SSG", ("SSG",), COMMON_ALIASES),
    ChannelAdapter("지마켓", ("지마켓", "G마켓"), COMMON_ALIASES),
    ChannelAdapter("옥션", ("옥션",), COMMON_ALIASES),
    ChannelAdapter("베네피아", ("베네피아",), COMMON_ALIASES),
    ChannelAdapter("현대홈쇼핑", ("현대홈쇼핑",), COMMON_ALIASES),
    ChannelAdapter("현대백화점", ("현대백화점",), COMMON_ALIASES),
)

def detect_adapter(sheet_names: list[str]) -> ChannelAdapter | None:
    normalized = {name.strip().lower() for name in sheet_names}

    for adapter in ADAPTERS:
        if any(name.lower() in normalized for name in adapter.sheet_names):
            return adapter

    return None