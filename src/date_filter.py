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
    rci_records: "pd.DataFrame"
    pdf_records: "pd.DataFrame"
    rci_out_of_period_records: "pd.DataFrame"
    pdf_out_of_period_records: "pd.DataFrame"
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

    erp_before_count = len(erp)
    rci_before_count = len(rci)
    pdf_before_count = len(pdf)
    if disable or not _as_bool(config.get("enabled", True)):
        summary = _summary(
            enabled=False,
            mode="disabled",
            start_date=None,
            end_date=None,
            erp_before_count=erp_before_count,
            erp_after_count=erp_before_count,
            rci_before_count=rci_before_count,
            rci_after_count=rci_before_count,
            pdf_before_count=pdf_before_count,
            pdf_after_count=pdf_before_count,
            source="disabled",
        )
        LOGGER.info("Filtre date desactive: %s ligne(s) ERP conservee(s).", erp_before_count)
        empty_rci = _empty_like(rci)
        empty_pdf = _empty_like(pdf)
        return DateFilterResult(erp.copy(), rci.copy(), pdf.copy(), empty_rci, empty_pdf, summary)

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
    filtered_erp = erp[mask].copy() if not erp.empty else pandas.DataFrame(columns=erp.columns)
    erp_after_count = len(filtered_erp)
    erp_excluded_count = erp_before_count - erp_after_count

    filtered_rci, rci_out_of_period, rci_missing_date_count = _filter_source_by_date(
        rci,
        "rci_date",
        "RCI",
        start_date,
        end_date,
    )
    filtered_pdf, pdf_out_of_period, pdf_missing_date_count = _filter_source_by_date(
        pdf,
        "pdf_invoice_date",
        "PDF",
        start_date,
        end_date,
    )
    rci_after_count = len(filtered_rci)
    pdf_after_count = len(filtered_pdf)

    summary = _summary(
        enabled=True,
        mode=mode,
        start_date=start_date,
        end_date=end_date,
        erp_before_count=erp_before_count,
        erp_after_count=erp_after_count,
        rci_before_count=rci_before_count,
        rci_after_count=rci_after_count,
        pdf_before_count=pdf_before_count,
        pdf_after_count=pdf_after_count,
        source=source,
        rci_missing_date_count=rci_missing_date_count,
        pdf_missing_date_count=pdf_missing_date_count,
    )
    LOGGER.info(
        "Periode de rapprochement utilisee: %s -> %s (mode=%s, source=%s).",
        start_date.isoformat(),
        end_date.isoformat(),
        mode,
        source,
    )
    LOGGER.info("ERP avant filtre date: %s ligne(s).", erp_before_count)
    LOGGER.info("ERP apres filtre date: %s ligne(s).", erp_after_count)
    LOGGER.info("ERP exclu par filtre date: %s ligne(s).", erp_excluded_count)
    LOGGER.info("RCI avant filtre date: %s ligne(s).", rci_before_count)
    LOGGER.info("RCI apres filtre date: %s ligne(s).", rci_after_count)
    LOGGER.info("RCI exclu par periode: %s ligne(s).", len(rci_out_of_period))
    LOGGER.info("PDF avant filtre date: %s ligne(s).", pdf_before_count)
    LOGGER.info("PDF apres filtre date: %s ligne(s).", pdf_after_count)
    LOGGER.info("PDF exclu par periode: %s ligne(s).", len(pdf_out_of_period))
    if summary["no_rci_flux_in_period_alert"]:
        LOGGER.warning("Attention : aucun flux RCI dans la période de rapprochement.")
    if summary["pdf_period_mismatch_alert"]:
        LOGGER.warning("Attention : les PDF chargés ne correspondent pas à la période du flux RCI.")
    return DateFilterResult(
        filtered_erp,
        filtered_rci,
        filtered_pdf,
        rci_out_of_period,
        pdf_out_of_period,
        summary,
    )


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

    rci_dates = _dates_from_column(rci, "rci_date")
    if rci_dates:
        days_before = _int_config(config, "days_before", 3)
        days_after = _int_config(config, "days_after", 1)
        return (
            min(rci_dates) - timedelta(days=days_before),
            max(rci_dates) + timedelta(days=days_after),
            "auto",
            "rci_txt",
        )

    pdf_dates = _dates_from_column(pdf, "pdf_invoice_date")
    if pdf_dates and rci.empty:
        days_before = _int_config(config, "days_before", 3)
        days_after = _int_config(config, "days_after", 1)
        return (
            min(pdf_dates) - timedelta(days=days_before),
            max(pdf_dates) + timedelta(days=days_after),
            "auto",
            "pdf",
        )

    if pdf_dates and not rci.empty:
        LOGGER.warning(
            "Dates RCI non fiables ou absentes: les PDF ne sont pas utilises pour elargir automatiquement la periode RCI."
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
    erp_before_count: int,
    erp_after_count: int,
    rci_before_count: int,
    rci_after_count: int,
    pdf_before_count: int,
    pdf_after_count: int,
    source: str,
    rci_missing_date_count: int = 0,
    pdf_missing_date_count: int = 0,
) -> dict[str, Any]:
    period = (
        f"{start_date.isoformat()} -> {end_date.isoformat()}"
        if start_date is not None and end_date is not None
        else "filtre desactive"
    )
    rci_pdf_before_count = rci_before_count + pdf_before_count
    rci_pdf_after_count = rci_after_count + pdf_after_count
    return {
        "date_filter_enabled": enabled,
        "date_filter_mode": mode,
        "date_filter_source": source,
        "reconciliation_start_date": start_date.isoformat() if start_date else None,
        "reconciliation_end_date": end_date.isoformat() if end_date else None,
        "reconciliation_period": period,
        "erp_rows_before_date_filter": erp_before_count,
        "erp_rows_after_date_filter": erp_after_count,
        "erp_rows_excluded_by_date": erp_before_count - erp_after_count,
        "rci_rows_before_date_filter": rci_before_count,
        "rci_rows_after_date_filter": rci_after_count,
        "rci_rows_excluded_by_date": rci_before_count - rci_after_count,
        "pdf_rows_before_date_filter": pdf_before_count,
        "pdf_rows_after_date_filter": pdf_after_count,
        "pdf_rows_excluded_by_date": pdf_before_count - pdf_after_count,
        "rci_pdf_rows_before_date_filter": rci_pdf_before_count,
        "rci_pdf_rows_after_date_filter": rci_pdf_after_count,
        "rci_pdf_rows_excluded_by_date": rci_pdf_before_count - rci_pdf_after_count,
        "rci_rows_missing_date_kept": rci_missing_date_count,
        "pdf_rows_missing_date_kept": pdf_missing_date_count,
        "no_rci_flux_in_period_alert": bool(enabled and rci_before_count > 0 and rci_after_count == 0),
        "pdf_period_mismatch_alert": bool(
            enabled
            and source == "rci_txt"
            and pdf_before_count > 0
            and pdf_before_count > pdf_after_count
        ),
    }


def _filter_source_by_date(
    frame: "pd.DataFrame",
    date_column: str,
    source_label: str,
    start_date: date,
    end_date: date,
) -> tuple["pd.DataFrame", "pd.DataFrame", int]:
    pandas = _require_pandas()
    if frame.empty:
        return frame.copy(), pandas.DataFrame(columns=frame.columns), 0

    if date_column not in frame.columns:
        LOGGER.warning(
            "Filtre date %s: colonne %s absente, %s ligne(s) conservee(s) sans controle de periode.",
            source_label,
            date_column,
            len(frame),
        )
        kept = frame.copy()
        kept["date_filter_status"] = "date_missing_kept"
        return kept, pandas.DataFrame(columns=frame.columns), len(frame)

    parsed_dates = _normalized_date_series(frame, date_column)
    in_period_mask = parsed_dates.map(lambda value: value is not None and start_date <= value <= end_date)
    missing_date_mask = parsed_dates.map(lambda value: value is None)
    out_of_period_mask = ~(in_period_mask | missing_date_mask)

    kept = frame[in_period_mask | missing_date_mask].copy()
    out_of_period = frame[out_of_period_mask].copy()

    if not kept.empty:
        kept["date_filter_status"] = [
            "date_missing_kept" if parsed_dates.loc[index] is None else "in_period"
            for index in kept.index
        ]
    if not out_of_period.empty:
        out_of_period["date_filter_status"] = "out_of_period"
        out_of_period["commentaire_audit"] = "Hors période de rapprochement"

    missing_date_count = int(missing_date_mask.sum())
    if missing_date_count:
        LOGGER.warning(
            "Filtre date %s: %s ligne(s) sans %s conservee(s); periode impossible a confirmer.",
            source_label,
            missing_date_count,
            date_column,
        )

    return kept.reset_index(drop=True), out_of_period.reset_index(drop=True), missing_date_count


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


def _empty_like(frame: "pd.DataFrame") -> "pd.DataFrame":
    return _require_pandas().DataFrame(columns=frame.columns)


def _require_pandas() -> Any:
    if pd is None:
        raise RuntimeError(
            "Le module pandas est requis pour filtrer la periode de rapprochement. "
            "Installez les dependances avec: python -m pip install -r requirements.txt"
        )
    return pd
