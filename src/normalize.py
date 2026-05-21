from __future__ import annotations

import math
import re
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any


EMPTY_MARKERS = {"", "NONE", "NULL", "NAN", "NAT"}
INVOICE_PATTERN = re.compile(r"^(VF|AAF)\d+$")
RAW_FACTURE_PATTERN = re.compile(r"^(FVF|VF)\d+$")
EXCEL_EPOCH = datetime(1899, 12, 30)


def file_record(path: Path, source_type: str) -> dict[str, Any]:
    stats = path.stat()
    return {
        "source_type": source_type,
        "file_name": path.name,
        "file_path": str(path),
        "extension": path.suffix.lower(),
        "size_bytes": stats.st_size,
        "modified_at": datetime.fromtimestamp(stats.st_mtime, tz=timezone.utc).isoformat(),
        "status": "detected",
    }


def normalize_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    ingested_at = datetime.now(tz=timezone.utc).isoformat()
    return [{**record, "ingested_at": ingested_at} for record in records]


def normalize_invoice_number(value: Any) -> str | None:
    if _is_empty(value):
        return None

    normalized = re.sub(r"\s+", "", str(value)).upper()
    if not re.fullmatch(r"[A-Z0-9]+", normalized):
        return None

    if normalized.startswith("FVF"):
        normalized = f"VF{normalized[3:]}"

    if not INVOICE_PATTERN.fullmatch(normalized):
        return None
    return normalized


def detect_document_type(invoice_number: Any) -> str:
    if _is_empty(invoice_number):
        return "UNKNOWN"

    raw_value = re.sub(r"\s+", "", str(invoice_number)).upper()
    if RAW_FACTURE_PATTERN.fullmatch(raw_value):
        return "FACTURE"

    normalized = normalize_invoice_number(invoice_number)
    if normalized and normalized.startswith("VF"):
        return "FACTURE"
    if normalized and normalized.startswith("AAF"):
        return "AVOIR"
    return "UNKNOWN"


def parse_french_amount(value: Any) -> float | None:
    if _is_empty(value) or isinstance(value, bool):
        return None

    if isinstance(value, Decimal):
        if value.is_nan():
            return None
        return float(value)

    if isinstance(value, int):
        return float(value)

    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        return float(value)

    cleaned = str(value).strip()
    if _is_empty(cleaned):
        return None

    cleaned = re.sub(r"[\s\u00a0\u202f]+", "", cleaned)
    cleaned = re.sub(r"(?i)(mad|dhs?|eur|euro|\$)", "", cleaned)

    negative = False
    if cleaned.startswith("(") and cleaned.endswith(")"):
        negative = True
        cleaned = cleaned[1:-1]

    if cleaned.endswith("-"):
        negative = True
        cleaned = cleaned[:-1]

    if cleaned.startswith("-"):
        negative = True
        cleaned = cleaned[1:]
    elif cleaned.startswith("+"):
        cleaned = cleaned[1:]

    if not cleaned or not re.fullmatch(r"\d+(?:[.,]\d+)*", cleaned):
        return None

    numeric_text = _french_number_to_decimal_text(cleaned)
    if numeric_text is None:
        return None

    try:
        amount = Decimal(numeric_text)
    except InvalidOperation:
        return None

    if negative:
        amount = -amount
    return float(amount)


def normalize_date(value: Any) -> str | None:
    if _is_empty(value) or isinstance(value, bool):
        return None

    if hasattr(value, "to_pydatetime"):
        try:
            value = value.to_pydatetime()
        except (TypeError, ValueError, AttributeError):
            return None

    if isinstance(value, datetime):
        return value.date().isoformat()

    if isinstance(value, date):
        return value.isoformat()

    if isinstance(value, (int, float, Decimal)):
        return _normalize_excel_serial(value)

    text = " ".join(str(value).strip().split())
    if _is_empty(text):
        return None

    if re.fullmatch(r"\d+(?:[.,]\d+)?", text):
        serial_text = text.replace(",", ".")
        try:
            return _normalize_excel_serial(Decimal(serial_text))
        except InvalidOperation:
            return None

    parsed = _parse_text_date(text)
    if parsed is None:
        return None
    return parsed.isoformat()


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split()).upper()


def normalize_amount(value: Any) -> Decimal | None:
    amount = parse_french_amount(value)
    if amount is None:
        return None
    try:
        return Decimal(str(amount))
    except InvalidOperation:
        return None


def _is_empty(value: Any) -> bool:
    if value is None:
        return True

    if isinstance(value, Decimal):
        return value.is_nan()

    if isinstance(value, float):
        return math.isnan(value)

    return str(value).strip().upper() in EMPTY_MARKERS


def _french_number_to_decimal_text(cleaned: str) -> str | None:
    if "," in cleaned:
        if cleaned.count(",") > 1:
            return None
        whole_part, decimal_part = cleaned.rsplit(",", 1)
        if not whole_part or not decimal_part:
            return None
        return f"{whole_part.replace('.', '')}.{decimal_part}"

    if "." not in cleaned:
        return cleaned

    parts = cleaned.split(".")
    if len(parts) > 2:
        if all(len(part) == 3 for part in parts[1:]):
            return "".join(parts)
        return None

    whole_part, decimal_or_group = parts
    if len(decimal_or_group) == 3 and len(whole_part) <= 3:
        return f"{whole_part}{decimal_or_group}"
    return cleaned


def _normalize_excel_serial(value: int | float | Decimal) -> str | None:
    try:
        serial = float(value)
    except (TypeError, ValueError, InvalidOperation):
        return None

    if math.isnan(serial) or math.isinf(serial) or serial <= 0:
        return None

    try:
        parsed = EXCEL_EPOCH + timedelta(days=serial)
    except OverflowError:
        return None
    return parsed.date().isoformat()


def _parse_text_date(text: str) -> date | None:
    iso_candidate = text.replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(iso_candidate).date()
    except ValueError:
        pass

    formats = (
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d-%m-%Y",
        "%d-%m-%y",
        "%d.%m.%Y",
        "%d.%m.%y",
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
    )
    for date_format in formats:
        try:
            return datetime.strptime(text, date_format).date()
        except ValueError:
            continue
    return None
