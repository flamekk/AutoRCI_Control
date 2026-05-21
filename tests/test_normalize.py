from __future__ import annotations

from datetime import date, datetime

from src.normalize import (
    detect_document_type,
    normalize_date,
    normalize_invoice_number,
    parse_french_amount,
)


def test_normalize_invoice_number_removes_spaces_and_fvf_prefix() -> None:
    assert normalize_invoice_number(" FVF 385380 ") == "VF385380"
    assert normalize_invoice_number("VF385380") == "VF385380"
    assert normalize_invoice_number("aaf31700") == "AAF31700"


def test_normalize_invoice_number_rejects_empty_or_invalid_values() -> None:
    assert normalize_invoice_number(None) is None
    assert normalize_invoice_number("") is None
    assert normalize_invoice_number("ABC123") is None
    assert normalize_invoice_number("VF 38/5380") is None


def test_detect_document_type() -> None:
    assert detect_document_type("VF385380") == "FACTURE"
    assert detect_document_type("FVF385380") == "FACTURE"
    assert detect_document_type("AAF31700") == "AVOIR"
    assert detect_document_type("ZZ123") == "UNKNOWN"


def test_parse_french_amount() -> None:
    assert parse_french_amount("1.189.358,56") == 1189358.56
    assert parse_french_amount("865,79-") == -865.79
    assert parse_french_amount("-865,79") == -865.79
    assert parse_french_amount(" 1 189 358,56 ") == 1189358.56


def test_parse_french_amount_rejects_invalid_values() -> None:
    assert parse_french_amount(None) is None
    assert parse_french_amount("") is None
    assert parse_french_amount("abc") is None
    assert parse_french_amount("12,34,56") is None


def test_normalize_date_from_common_values() -> None:
    assert normalize_date("15/05/26") == "2026-05-15"
    assert normalize_date("15/05/2026") == "2026-05-15"
    assert normalize_date(datetime(2026, 5, 15, 10, 30)) == "2026-05-15"
    assert normalize_date(date(2026, 5, 15)) == "2026-05-15"


def test_normalize_date_from_excel_serial_and_timestamp_like_object() -> None:
    assert normalize_date(46157) == "2026-05-15"

    class TimestampLike:
        def to_pydatetime(self) -> datetime:
            return datetime(2026, 5, 15, 0, 0)

    assert normalize_date(TimestampLike()) == "2026-05-15"


def test_normalize_date_rejects_invalid_values() -> None:
    assert normalize_date(None) is None
    assert normalize_date("") is None
    assert normalize_date("pas une date") is None
