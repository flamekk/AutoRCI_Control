from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any

try:
    import pandas as pd
except ModuleNotFoundError:  # pragma: no cover - exercised only before dependency install.
    pd = None  # type: ignore[assignment]

try:
    from src.normalize import normalize_date
except ModuleNotFoundError:  # pragma: no cover - used when running python src/main.py.
    from normalize import normalize_date


LOGGER = logging.getLogger(__name__)


@dataclass(frozen=True)
class DateFilterResult:
    erp_records: "pd.DataFrame"
    summary: dict[str, Any]


def apply_reconciliation_date_filter(
    erp_records: Any,
    rci_records: Any,
    pdf_records: Any,
    date_filter_config: dict[str, Any] | None = None,
    *,
    disable: bool = False,
    date_from: str | None = None,
    date_to: str | None = None,
) -> DateFilterResult:
    pandas = _require_pandas()
    erp = _to_dataframe(erp_records)
    rci = _to_dataframe(rci_records)
    pdf = _to_dataframe(pdf_records)
    config = date_filter_config or {}

    before_count = len(erp)
    if disable or not _as_bool(config.get("enabled", True)):
        summary = _summary(
            enabled=False,
            mode="disabled",
            start_date=None,
            end_date=None,
            before_count=before_count,
            after_count=before_count,
            source="disabled",
        )
        LOGGER.info("Filtre date desactive: %s ligne(s) ERP conservee(s).", before_count)
        return DateFilterResult(erp.copy(), summary)

    start_date, end_date, mode, source = _resolve_period(
        erp,
        rci,
        pdf,
        config,
        date_from=date_from,
        date_to=date_to,
    )

    erp_dates = _normalized_date_series(erp, "erp_date")
    mask = erp_dates.map(lambda value: value is not None and start_date <= value <= end_date)
    filtered = erp[mask].copy() if not erp.empty else pandas.DataFrame(columns=erp.columns)
    after_count = len(filtered)
    excluded_count = before_count - after_count

    summary = _summary(
        enabled=True,
        mode=mode,
        start_date=start_date,
        end_date=end_date,
        before_count=before_count,
        after_count=after_count,
        source=source,
    )
    LOGGER.info(
        "Periode de rapprochement utilisee: %s -> %s (mode=%s, source=%s).",
        start_date.isoformat(),
        end_date.isoformat(),
        mode,
        source,
    )
    LOGGER.info("ERP avant filtre date: %s ligne(s).", before_count)
    LOGGER.info("ERP apres filtre date: %s ligne(s).", after_count)
    LOGGER.info("ERP exclu par filtre date: %s ligne(s).", excluded_count)
    return DateFilterResult(filtered, summary)


def _resolve_period(
    erp: "pd.DataFrame",
    rci: "pd.DataFrame",
    pdf: "pd.DataFrame",
    config: dict[str, Any],
    *,
    date_from: str | None,
    date_to: str | None,
) -> tuple[date, date, str, str]:
    if date_from or date_to:
        if not date_from or not date_to:
            raise ValueError("Les options --date-from et --date-to doivent etre fournies ensemble.")
        start_date = _parse_date_or_fail(date_from, "--date-from")
        end_date = _parse_date_or_fail(date_to, "--date-to")
        if start_date > end_date:
            raise ValueError("--date-from doit etre inferieur ou egal a --date-to.")
        return start_date, end_date, "manual", "cli"

    mode = str(config.get("mode", "auto") or "auto").strip().lower()
    if mode != "auto":
        raise ValueError(f"Mode de filtre date non supporte: {mode}")

    source_dates = [
        *_dates_from_column(rci, "rci_date"),
        *_dates_from_column(pdf, "pdf_invoice_date"),
    ]
    if source_dates:
        days_before = _int_config(config, "days_before", 3)
        days_after = _int_config(config, "days_after", 1)
        return (
            min(source_dates) - timedelta(days=days_before),
            max(source_dates) + timedelta(days=days_after),
            "auto",
            "rci_pdf",
        )

    erp_dates = _dates_from_column(erp, "erp_date")
    if not erp_dates:
        raise ValueError("Impossible de definir la periode: aucune date ERP/RCI/PDF exploitable.")

    fallback_days = max(_int_config(config, "fallback_days", 7), 1)
    end_date = max(erp_dates)
    start_date = end_date - timedelta(days=fallback_days - 1)
    LOGGER.warning(
        "Aucune date fiable detectee cote RCI/PDF. Utilisation des %s dernier(s) jour(s) ERP: %s -> %s.",
        fallback_days,
        start_date.isoformat(),
        end_date.isoformat(),
    )
    return start_date, end_date, "fallback", "erp"


def _summary(
    *,
    enabled: bool,
    mode: str,
    start_date: date | None,
    end_date: date | None,
    before_count: int,
    after_count: int,
    source: str,
) -> dict[str, Any]:
    period = (
        f"{start_date.isoformat()} -> {end_date.isoformat()}"
        if start_date is not None and end_date is not None
        else "filtre desactive"
    )
    return {
        "date_filter_enabled": enabled,
        "date_filter_mode": mode,
        "date_filter_source": source,
        "reconciliation_start_date": start_date.isoformat() if start_date else None,
        "reconciliation_end_date": end_date.isoformat() if end_date else None,
        "reconciliation_period": period,
        "erp_rows_before_date_filter": before_count,
        "erp_rows_after_date_filter": after_count,
        "erp_rows_excluded_by_date": before_count - after_count,
    }


def _normalized_date_series(frame: "pd.DataFrame", column: str) -> "pd.Series":
    pandas = _require_pandas()
    if frame.empty or column not in frame.columns:
        return pandas.Series([None] * len(frame), index=frame.index, dtype=object)
    return frame[column].map(_parse_date)


def _dates_from_column(frame: "pd.DataFrame", column: str) -> list[date]:
    if frame.empty or column not in frame.columns:
        return []
    dates = []
    for value in frame[column]:
        parsed = _parse_date(value)
        if parsed is not None:
            dates.append(parsed)
    return dates


def _parse_date(value: Any) -> date | None:
    normalized = normalize_date(value)
    if normalized is None:
        return None
    try:
        return date.fromisoformat(normalized)
    except ValueError:
        return None


def _parse_date_or_fail(value: str, label: str) -> date:
    parsed = _parse_date(value)
    if parsed is None:
        raise ValueError(f"Date invalide pour {label}: {value}. Format attendu YYYY-MM-DD.")
    return parsed


def _int_config(config: dict[str, Any], key: str, default: int) -> int:
    try:
        return int(config.get(key, default))
    except (TypeError, ValueError):
        return default


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on", "oui"}


def _to_dataframe(records: Any) -> "pd.DataFrame":
    pandas = _require_pandas()
    if records is None:
        return pandas.DataFrame()
    if hasattr(records, "copy") and hasattr(records, "columns"):
        return records.copy()
    if isinstance(records, list):
        return pandas.DataFrame(records)
    return pandas.DataFrame(list(records or []))


def _require_pandas() -> Any:
    if pd is None:
        raise RuntimeError(
            "Le module pandas est requis pour filtrer la periode de rapprochement. "
            "Installez les dependances avec: python -m pip install -r requirements.txt"
        )
    return pd
