from __future__ import annotations

import logging
import math
from datetime import datetime, timezone
from typing import Any

try:
    import pandas as pd
except ModuleNotFoundError:  # pragma: no cover - exercised only before dependency install.
    pd = None  # type: ignore[assignment]

try:
    from normalize import detect_document_type, normalize_date, normalize_invoice_number, parse_french_amount
except ModuleNotFoundError:  # pragma: no cover - useful when imported as src.reconcile in tests.
    from src.normalize import (
        detect_document_type,
        normalize_date,
        normalize_invoice_number,
        parse_french_amount,
    )


LOGGER = logging.getLogger(__name__)

DEFAULT_AMOUNT_TOLERANCE_MAD = 1.0

RECONCILIATION_COLUMNS = [
    "invoice_number",
    "document_type",
    "customer_code",
    "customer_name",
    "is_rci_covered",
    "amount_erp",
    "amount_rci",
    "amount_pdf",
    "amount_gap",
    "montant_impacte",
    "erp_date",
    "rci_date",
    "pdf_invoice_date",
    "due_date",
    "origin",
    "status",
    "priority",
    "action_recommandee",
    "source_erp",
    "source_rci",
    "source_pdf",
]

STATUS_ACTIONS = {
    "OK": "Aucune action",
    "MANQUANTE_RCI": (
        "Verifier l'integration dans le prochain flux ou preparer un renvoi "
        "(ecart de transmission, d'integration ou de reception du flux)"
    ),
    "ANOMALIE_MONTANT": "Controler montant ERP vs montant finance RCI",
    "ANOMALIE_DATE": "Verifier date facture/date echeance",
    "DOUBLON": "Verifier duplication de facture",
    "RCI_SEULEMENT": "Verifier origine cote RCI ou historique ERP",
    "HORS_SCOPE_RCI": "Aucune action: facture ERP hors perimetre de couverture RCI",
}

STATUS_PRIORITIES = {
    "OK": "BASSE",
    "ANOMALIE_DATE": "MOYENNE",
    "RCI_SEULEMENT": "MOYENNE",
    "MANQUANTE_RCI": "HAUTE",
    "ANOMALIE_MONTANT": "HAUTE",
    "DOUBLON": "HAUTE",
    "HORS_SCOPE_RCI": "BASSE",
}


def reconcile(
    erp_records: Any,
    rci_records: Any,
    pdf_records: Any,
    missing_required: dict[str, int] | None = None,
    amount_tolerance: float = DEFAULT_AMOUNT_TOLERANCE_MAD,
) -> dict[str, Any]:
    missing_required = missing_required or {}
    erp_frame = _to_dataframe(erp_records)
    rci_frame = _to_dataframe(rci_records)
    pdf_frame = _to_dataframe(pdf_records)

    reconciliation = reconcile_dataframes(
        erp_frame,
        rci_frame,
        pdf_frame,
        amount_tolerance=amount_tolerance,
    )

    erp_rows = _to_records(erp_frame, "erp")
    rci_rows = _to_records(rci_frame, "rci")
    pdf_rows = _to_records(pdf_frame, "pdf")

    blocking_anomalies = _missing_required_anomalies(missing_required)
    reconciliation_anomalies = _reconciliation_anomalies(reconciliation)
    anomalies = [*blocking_anomalies, *reconciliation_anomalies]

    status_counts = _status_counts(reconciliation)
    total_controlled_amount = _sum_abs_column(reconciliation, "amount_erp")
    total_impacted_amount = _sum_abs_column(reconciliation, "montant_impacte")
    missing_rci_amount = _sum_abs_column(
        reconciliation[reconciliation["status"] == "MANQUANTE_RCI"]
        if not reconciliation.empty
        else reconciliation,
        "montant_impacte",
    )
    erp_matchable_invoices = _erp_matchable_count(status_counts)
    matching_rate = (
        round(status_counts.get("OK", 0) / erp_matchable_invoices, 4)
        if erp_matchable_invoices
        else 0.0
    )
    gaps_detected = sum(
        status_counts.get(status, 0)
        for status in {
            "MANQUANTE_RCI",
            "ANOMALIE_MONTANT",
            "ANOMALIE_DATE",
            "DOUBLON",
            "RCI_SEULEMENT",
        }
    )
    for status, count in status_counts.items():
        LOGGER.info("Rapprochement %s: %s ligne(s)", status, count)

    return {
        "generated_at": datetime.now(tz=timezone.utc).isoformat(),
        "status": "completed_with_anomalies" if anomalies else "ok",
        "summary": {
            "erp_files": _count_distinct_files(erp_rows),
            "erp_rows": len(erp_rows),
            "rci_files": _count_distinct_files(rci_rows),
            "rci_rows": len(rci_rows),
            "pdf_files": _count_distinct_files(pdf_rows),
            "pdf_rows": len(pdf_rows),
            "reconciled_invoices": len(reconciliation),
            "matched_invoices": status_counts.get("OK", 0),
            "unmatched_erp": status_counts.get("MANQUANTE_RCI", 0),
            "unmatched_rci": status_counts.get("RCI_SEULEMENT", 0),
            "out_of_scope_rci": status_counts.get("HORS_SCOPE_RCI", 0),
            "rci_covered_invoices": _covered_invoice_count(reconciliation),
            "amount_anomalies": status_counts.get("ANOMALIE_MONTANT", 0),
            "date_anomalies": status_counts.get("ANOMALIE_DATE", 0),
            "duplicates": status_counts.get("DOUBLON", 0),
            "total_controlled_amount": total_controlled_amount,
            "total_amount_gap": total_impacted_amount,
            "total_impacted_amount": total_impacted_amount,
            "missing_rci_amount": missing_rci_amount,
            "erp_matchable_invoices": erp_matchable_invoices,
            "matching_rate": matching_rate,
            "gaps_detected": gaps_detected,
            "anomalies": len(anomalies),
            "amount_tolerance_mad": amount_tolerance,
            **{f"status_{status.lower()}": count for status, count in status_counts.items()},
        },
        "source_files": [*erp_rows, *rci_rows, *pdf_rows],
        "reconciliation": _dataframe_records(reconciliation),
        "anomalies": anomalies,
        "note": (
            "Rapprochement effectue sur invoice_number normalise. Les ecarts de presence sont "
            "formules comme des ecarts de transmission, d'integration ou de reception du flux."
        ),
    }


def reconcile_dataframes(
    erp_records: Any,
    rci_records: Any,
    pdf_records: Any,
    amount_tolerance: float = DEFAULT_AMOUNT_TOLERANCE_MAD,
) -> "pd.DataFrame":
    pandas = _require_pandas()
    erp = _aggregate_erp(_to_dataframe(erp_records))
    rci = _aggregate_rci(_to_dataframe(rci_records))
    pdf = _aggregate_pdf(_to_dataframe(pdf_records))
    consolidated_rci = _consolidate_rci_pdf(rci, pdf)

    if erp.empty and consolidated_rci.empty:
        return pandas.DataFrame(columns=RECONCILIATION_COLUMNS)

    merged = erp.merge(consolidated_rci, on="invoice_number", how="outer")
    records = []
    for _, row in merged.iterrows():
        records.append(_build_reconciliation_row(row, amount_tolerance))

    result = pandas.DataFrame(records).reindex(columns=RECONCILIATION_COLUMNS)
    return result.sort_values(["status", "invoice_number"], kind="stable").reset_index(drop=True)


def _aggregate_erp(frame: "pd.DataFrame") -> "pd.DataFrame":
    prepared = _prepare_source_frame(frame)
    if prepared.empty:
        return _empty_frame(
            [
                "invoice_number",
                "document_type_erp",
                "customer_code",
                "customer_name",
                "amount_erp",
                "erp_date",
                "is_rci_covered",
                "source_erp",
                "erp_count",
                "erp_duplicate_conflict",
            ]
        )

    records = []
    for invoice_number, group in prepared.groupby("invoice_number", sort=True):
        document_type = _first_non_empty(group.get("document_type")) or detect_document_type(invoice_number)
        signed_amounts = _distinct_signed_amounts(group.get("amount_erp"), document_type)
        document_types = _distinct_values(group.get("document_type"))
        dates = _distinct_normalized_dates(group.get("erp_date"))
        records.append(
            {
                "invoice_number": invoice_number,
                "document_type_erp": document_type,
                "customer_code": _first_non_empty(group.get("customer_code")),
                "customer_name": _first_non_empty(group.get("customer_name")),
                "is_rci_covered": _group_is_rci_covered(group),
                "amount_erp": _first_from_list(signed_amounts),
                "erp_date": _first_normalized_date(group.get("erp_date")),
                "source_erp": _join_unique(_source_labels(group, "erp")),
                "erp_count": len(group),
                "erp_duplicate_conflict": _has_duplicate_conflict(group, signed_amounts, document_types, dates),
            }
        )

    return _require_pandas().DataFrame(records)


def _aggregate_rci(frame: "pd.DataFrame") -> "pd.DataFrame":
    prepared = _prepare_source_frame(frame)
    if prepared.empty:
        return _empty_frame(
            [
                "invoice_number",
                "document_type_rci",
                "amount_rci",
                "rci_date",
                "source_rci",
                "rci_count",
                "rci_duplicate_conflict",
            ]
        )

    records = []
    for invoice_number, group in prepared.groupby("invoice_number", sort=True):
        document_type = _first_non_empty(group.get("document_type")) or detect_document_type(invoice_number)
        signed_amounts = _distinct_signed_amounts(group.get("amount_rci"), document_type)
        document_types = _distinct_values(group.get("document_type"))
        dates = _distinct_normalized_dates(group.get("rci_date"))
        records.append(
            {
                "invoice_number": invoice_number,
                "document_type_rci": document_type,
                "amount_rci": _first_from_list(signed_amounts),
                "rci_date": _first_normalized_date(group.get("rci_date")),
                "source_rci": _join_unique(_source_labels(group, "rci")),
                "rci_count": len(group),
                "rci_duplicate_conflict": _has_duplicate_conflict(group, signed_amounts, document_types, dates),
            }
        )

    return _require_pandas().DataFrame(records)


def _aggregate_pdf(frame: "pd.DataFrame") -> "pd.DataFrame":
    prepared = _prepare_source_frame(frame)
    if prepared.empty:
        return _empty_frame(
            [
                "invoice_number",
                "document_type_pdf",
                "amount_pdf",
                "pdf_invoice_date",
                "due_date",
                "origin",
                "source_pdf",
                "pdf_count",
                "pdf_duplicate_conflict",
            ]
        )

    records = []
    for invoice_number, group in prepared.groupby("invoice_number", sort=True):
        document_type = _first_non_empty(group.get("document_type")) or detect_document_type(invoice_number)
        signed_amounts = _distinct_signed_amounts(group.get("amount_pdf"), document_type)
        document_types = _distinct_values(group.get("document_type"))
        dates = _distinct_normalized_dates(group.get("pdf_invoice_date"))
        records.append(
            {
                "invoice_number": invoice_number,
                "document_type_pdf": document_type,
                "amount_pdf": _first_from_list(signed_amounts),
                "pdf_invoice_date": _first_normalized_date(group.get("pdf_invoice_date")),
                "due_date": _first_normalized_date(group.get("due_date")),
                "origin": _first_non_empty(group.get("origin")),
                "source_pdf": _join_unique(_source_labels(group, "pdf")),
                "pdf_count": len(group),
                "pdf_duplicate_conflict": _has_duplicate_conflict(group, signed_amounts, document_types, dates),
            }
        )

    return _require_pandas().DataFrame(records)


def _prepare_source_frame(frame: "pd.DataFrame") -> "pd.DataFrame":
    pandas = _require_pandas()
    if frame.empty or "invoice_number" not in frame.columns:
        return pandas.DataFrame(columns=list(frame.columns) + ["invoice_number"])

    prepared = frame.copy()
    prepared["invoice_number"] = prepared["invoice_number"].map(normalize_invoice_number)
    prepared = prepared[prepared["invoice_number"].notna()].copy()
    if "document_type" not in prepared.columns:
        prepared["document_type"] = prepared["invoice_number"].map(detect_document_type)
    return prepared


def _consolidate_rci_pdf(rci: "pd.DataFrame", pdf: "pd.DataFrame") -> "pd.DataFrame":
    pandas = _require_pandas()
    if rci.empty and pdf.empty:
        return _empty_frame(
            [
                "invoice_number",
                "document_type_consolidated",
                "amount_rci",
                "amount_pdf",
                "rci_date",
                "pdf_invoice_date",
                "due_date",
                "origin",
                "source_rci",
                "source_pdf",
                "rci_count",
                "pdf_count",
                "rci_duplicate_conflict",
                "pdf_duplicate_conflict",
                "consolidated_duplicate_conflict",
            ]
        )

    consolidated = rci.merge(pdf, on="invoice_number", how="outer")
    if "document_type_rci" not in consolidated.columns:
        consolidated["document_type_rci"] = None
    if "document_type_pdf" not in consolidated.columns:
        consolidated["document_type_pdf"] = None

    consolidated["document_type_consolidated"] = consolidated.apply(
        lambda row: _first_value(
            row.get("document_type_pdf"),
            row.get("document_type_rci"),
            detect_document_type(row.get("invoice_number")),
        ),
        axis=1,
    )

    for column in ["rci_count", "pdf_count"]:
        if column not in consolidated.columns:
            consolidated[column] = 0
        consolidated[column] = consolidated[column].fillna(0).astype(int)

    for column in ["rci_duplicate_conflict", "pdf_duplicate_conflict"]:
        if column not in consolidated.columns:
            consolidated[column] = False
        consolidated[column] = consolidated[column].fillna(False).astype(bool)

    consolidated["consolidated_duplicate_conflict"] = consolidated.apply(
        lambda row: _consolidated_conflict(row),
        axis=1,
    )

    expected = [
        "invoice_number",
        "document_type_consolidated",
        "amount_rci",
        "amount_pdf",
        "rci_date",
        "pdf_invoice_date",
        "due_date",
        "origin",
        "source_rci",
        "source_pdf",
        "rci_count",
        "pdf_count",
        "rci_duplicate_conflict",
        "pdf_duplicate_conflict",
        "consolidated_duplicate_conflict",
    ]
    for column in expected:
        if column not in consolidated.columns:
            consolidated[column] = None
    return consolidated.reindex(columns=expected)


def _build_reconciliation_row(row: Any, amount_tolerance: float) -> dict[str, Any]:
    document_type = _first_value(
        row.get("document_type_erp"),
        row.get("document_type_consolidated"),
        detect_document_type(row.get("invoice_number")),
    )
    amount_erp = _clean_number(row.get("amount_erp"))
    amount_rci = _clean_number(row.get("amount_rci"))
    amount_pdf = _clean_number(row.get("amount_pdf"))
    reference_amount = amount_pdf if amount_pdf is not None else amount_rci
    amount_gap = _amount_gap(amount_erp, reference_amount)
    duplicate_conflict = bool(
        row.get("erp_duplicate_conflict")
        or row.get("rci_duplicate_conflict")
        or row.get("pdf_duplicate_conflict")
        or row.get("consolidated_duplicate_conflict")
    )

    erp_count = _safe_int(row.get("erp_count"))
    rci_count = _safe_int(row.get("rci_count"))
    pdf_count = _safe_int(row.get("pdf_count"))
    erp_present = erp_count > 0
    consolidated_present = rci_count > 0 or pdf_count > 0
    is_rci_covered = _as_bool_or_default(row.get("is_rci_covered"), True)

    status = _determine_status(
        erp_present=erp_present,
        consolidated_present=consolidated_present,
        is_rci_covered=is_rci_covered,
        duplicate_conflict=duplicate_conflict,
        amount_gap=amount_gap,
        amount_erp=amount_erp,
        reference_amount=reference_amount,
        erp_date=row.get("erp_date"),
        rci_date=row.get("rci_date"),
        pdf_invoice_date=row.get("pdf_invoice_date"),
        due_date=row.get("due_date"),
        pdf_present=pdf_count > 0,
        amount_tolerance=amount_tolerance,
    )

    return {
        "invoice_number": row.get("invoice_number"),
        "document_type": document_type,
        "customer_code": _none_if_missing(row.get("customer_code")),
        "customer_name": _none_if_missing(row.get("customer_name")),
        "is_rci_covered": is_rci_covered if erp_present else None,
        "amount_erp": amount_erp,
        "amount_rci": amount_rci,
        "amount_pdf": amount_pdf,
        "amount_gap": amount_gap,
        "montant_impacte": _impacted_amount(status, amount_erp, amount_rci, amount_pdf, amount_gap),
        "erp_date": _none_if_missing(row.get("erp_date")),
        "rci_date": _none_if_missing(row.get("rci_date")),
        "pdf_invoice_date": _none_if_missing(row.get("pdf_invoice_date")),
        "due_date": _none_if_missing(row.get("due_date")),
        "origin": _none_if_missing(row.get("origin")),
        "status": status,
        "priority": STATUS_PRIORITIES[status],
        "action_recommandee": STATUS_ACTIONS[status],
        "source_erp": _none_if_missing(row.get("source_erp")),
        "source_rci": _none_if_missing(row.get("source_rci")),
        "source_pdf": _none_if_missing(row.get("source_pdf")),
    }


def _determine_status(
    *,
    erp_present: bool,
    consolidated_present: bool,
    is_rci_covered: bool,
    duplicate_conflict: bool,
    amount_gap: float | None,
    amount_erp: float | None,
    reference_amount: float | None,
    erp_date: Any,
    rci_date: Any,
    pdf_invoice_date: Any,
    due_date: Any,
    pdf_present: bool,
    amount_tolerance: float,
) -> str:
    if erp_present and not consolidated_present:
        if not is_rci_covered:
            return "HORS_SCOPE_RCI"
        return "MANQUANTE_RCI"
    if consolidated_present and not erp_present:
        return "RCI_SEULEMENT"
    if duplicate_conflict:
        return "DOUBLON"
    if amount_erp is None or reference_amount is None or amount_gap is None:
        return "ANOMALIE_MONTANT"
    if abs(amount_gap) > amount_tolerance:
        return "ANOMALIE_MONTANT"
    if _has_date_anomaly(erp_date, rci_date, pdf_invoice_date, due_date, pdf_present):
        return "ANOMALIE_DATE"
    return "OK"


def _has_date_anomaly(
    erp_date: Any,
    rci_date: Any,
    pdf_invoice_date: Any,
    due_date: Any,
    pdf_present: bool,
) -> bool:
    normalized_erp_date = normalize_date(erp_date)
    normalized_rci_date = normalize_date(rci_date)
    normalized_pdf_date = normalize_date(pdf_invoice_date)
    normalized_due_date = normalize_date(due_date)
    consolidated_invoice_date = normalized_pdf_date or normalized_rci_date

    if normalized_erp_date is None or consolidated_invoice_date is None:
        return True
    if normalized_erp_date != consolidated_invoice_date:
        return True
    if pdf_present and normalized_due_date is None:
        return True
    return False


def _missing_required_anomalies(missing_required: dict[str, int]) -> list[dict[str, Any]]:
    anomalies = []
    for source_type, missing_count in missing_required.items():
        anomalies.append(
            {
                "severity": "blocking",
                "source_type": source_type,
                "message": f"{missing_count} fichier(s) requis manquant(s) dans la source {source_type}.",
            }
        )
    return anomalies


def _reconciliation_anomalies(reconciliation: "pd.DataFrame") -> list[dict[str, Any]]:
    if reconciliation.empty:
        return []

    anomalies = []
    anomaly_statuses = reconciliation[~reconciliation["status"].isin({"OK", "HORS_SCOPE_RCI"})]
    for record in anomaly_statuses.to_dict("records"):
        anomalies.append(
            {
                "severity": _severity_for_status(record["status"]),
                "source_type": "reconciliation",
                "invoice_number": record["invoice_number"],
                "status": record["status"],
                "message": f"{record['status']} - {record['action_recommandee']}",
            }
        )
    return anomalies


def _severity_for_status(status: str) -> str:
    if status in {"MANQUANTE_RCI", "ANOMALIE_MONTANT", "DOUBLON"}:
        return "high"
    if status in {"ANOMALIE_DATE", "RCI_SEULEMENT"}:
        return "medium"
    return "low"


def _status_counts(reconciliation: "pd.DataFrame") -> dict[str, int]:
    if reconciliation.empty:
        return {}
    return {
        str(status): int(count)
        for status, count in reconciliation["status"].value_counts(sort=False).items()
    }


def _erp_matchable_count(status_counts: dict[str, int]) -> int:
    return sum(
        status_counts.get(status, 0)
        for status in {"OK", "MANQUANTE_RCI", "ANOMALIE_MONTANT", "ANOMALIE_DATE", "DOUBLON"}
    )


def _covered_invoice_count(reconciliation: "pd.DataFrame") -> int:
    if reconciliation.empty or "is_rci_covered" not in reconciliation.columns:
        return 0
    covered = reconciliation[
        (reconciliation["is_rci_covered"] == True)  # noqa: E712 - pandas boolean mask.
        & (reconciliation["status"] != "RCI_SEULEMENT")
    ]
    return int(len(covered))


def _sum_abs_column(frame: "pd.DataFrame", column: str) -> float:
    if frame.empty or column not in frame.columns:
        return 0.0
    total = 0.0
    for value in frame[column]:
        number = _clean_number(value)
        if number is not None:
            total += abs(number)
    return round(total, 2)


def _impacted_amount(
    status: str,
    amount_erp: float | None,
    amount_rci: float | None,
    amount_pdf: float | None,
    amount_gap: float | None,
) -> float:
    if status == "OK":
        return 0.0
    if status == "MANQUANTE_RCI":
        return _abs_or_zero(amount_erp)
    if status == "RCI_SEULEMENT":
        return _abs_or_zero(amount_pdf if amount_pdf is not None else amount_rci)
    if status == "HORS_SCOPE_RCI":
        return 0.0
    if status == "DOUBLON":
        return _abs_or_zero(_first_value(amount_erp, amount_pdf, amount_rci))
    if status == "ANOMALIE_MONTANT":
        return _abs_or_zero(amount_gap)
    if status == "ANOMALIE_DATE":
        return _abs_or_zero(amount_erp)
    return 0.0


def _abs_or_zero(value: Any) -> float:
    number = _clean_number(value)
    return round(abs(number), 2) if number is not None else 0.0


def _consolidated_conflict(row: Any) -> bool:
    if row.get("rci_duplicate_conflict") or row.get("pdf_duplicate_conflict"):
        return True

    document_type_rci = _none_if_missing(row.get("document_type_rci"))
    document_type_pdf = _none_if_missing(row.get("document_type_pdf"))
    if document_type_rci and document_type_pdf and document_type_rci != document_type_pdf:
        return True

    return False


def _to_dataframe(records: Any) -> "pd.DataFrame":
    pandas = _require_pandas()
    if records is None:
        return pandas.DataFrame()
    if hasattr(records, "copy") and hasattr(records, "columns"):
        return records.copy()
    if isinstance(records, list):
        return pandas.DataFrame(records)
    return pandas.DataFrame(list(records or []))


def _to_records(records: Any, source_type: str) -> list[dict[str, Any]]:
    rows = _dataframe_records(_to_dataframe(records))
    normalized_rows: list[dict[str, Any]] = []
    for row in rows:
        enriched = dict(row)
        enriched.setdefault("source_type", source_type)
        enriched.setdefault("file_name", enriched.get("source_file", ""))
        enriched.setdefault(
            "status",
            "extracted" if source_type in {"erp", "rci", "pdf"} else enriched.get("status", "detected"),
        )
        normalized_rows.append(enriched)
    return normalized_rows


def _dataframe_records(frame: "pd.DataFrame") -> list[dict[str, Any]]:
    records = []
    for record in frame.to_dict("records"):
        records.append({key: _none_if_missing(value) for key, value in record.items()})
    return records


def _empty_frame(columns: list[str]) -> "pd.DataFrame":
    return _require_pandas().DataFrame(columns=columns)


def _first_amount(series: Any, document_type: str | None) -> float | None:
    if series is None:
        return None
    for value in series:
        amount = _signed_amount(value, document_type)
        if amount is not None:
            return amount
    return None


def _group_is_rci_covered(group: "pd.DataFrame") -> bool:
    if "is_rci_covered" not in group.columns:
        return True
    values = [_as_bool_or_default(value, None) for value in group["is_rci_covered"]]
    values = [value for value in values if value is not None]
    if not values:
        return True
    return any(values)


def _as_bool_or_default(value: Any, default: bool | None) -> bool | None:
    if _is_missing(value):
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y", "on", "oui"}:
        return True
    if text in {"0", "false", "no", "n", "off", "non"}:
        return False
    return default


def _first_from_list(values: list[Any]) -> Any:
    return values[0] if values else None


def _distinct_signed_amounts(series: Any, document_type: str | None) -> list[float]:
    if series is None:
        return []

    amounts: list[float] = []
    for value in series:
        amount = _signed_amount(value, document_type)
        if amount is None or amount in amounts:
            continue
        amounts.append(amount)
    return amounts


def _distinct_values(series: Any) -> list[Any]:
    if series is None:
        return []

    values: list[Any] = []
    for value in series:
        if _is_missing(value):
            continue
        normalized = str(value).strip().upper()
        if normalized not in values:
            values.append(normalized)
    return values


def _distinct_normalized_dates(series: Any) -> list[str]:
    if series is None:
        return []

    values: list[str] = []
    for value in series:
        normalized = normalize_date(value)
        if normalized is None or normalized in values:
            continue
        values.append(normalized)
    return values


def _has_duplicate_conflict(
    group: "pd.DataFrame",
    signed_amounts: list[float],
    document_types: list[Any],
    dates: list[str],
) -> bool:
    if len(group) <= 1:
        return False
    if len(signed_amounts) > 1:
        return True
    if len(document_types) > 1:
        return True
    if len(dates) > 1:
        return True
    return False


def _signed_amount(value: Any, document_type: str | None) -> float | None:
    amount = _clean_number(value)
    if amount is None:
        return None
    if document_type == "AVOIR":
        return round(-abs(amount), 2)
    if document_type == "FACTURE":
        return round(abs(amount), 2)
    return round(amount, 2)


def _clean_number(value: Any) -> float | None:
    if _is_missing(value):
        return None
    amount = parse_french_amount(value)
    if amount is None or math.isnan(amount) or math.isinf(amount):
        return None
    return round(float(amount), 2)


def _amount_gap(amount_erp: float | None, reference_amount: float | None) -> float | None:
    if amount_erp is None or reference_amount is None:
        return None
    return round(amount_erp - reference_amount, 2)


def _first_normalized_date(series: Any) -> str | None:
    if series is None:
        return None
    for value in series:
        normalized = normalize_date(value)
        if normalized is not None:
            return normalized
    return None


def _first_non_empty(series: Any) -> Any:
    if series is None:
        return None
    for value in series:
        if not _is_missing(value):
            return value
    return None


def _first_value(*values: Any) -> Any:
    for value in values:
        if not _is_missing(value):
            return value
    return None


def _source_labels(group: "pd.DataFrame", source_type: str) -> list[str]:
    labels = []
    for _, row in group.iterrows():
        source_file = _first_value(row.get("source_file"), row.get("file_name"))
        if _is_missing(source_file):
            continue

        if source_type == "erp":
            source_sheet = row.get("source_sheet")
            labels.append(f"{source_file}:{source_sheet}" if not _is_missing(source_sheet) else str(source_file))
        elif source_type == "pdf":
            page = row.get("page")
            labels.append(f"{source_file}:p{page}" if not _is_missing(page) else str(source_file))
        else:
            labels.append(str(source_file))
    return labels


def _join_unique(values: list[Any]) -> str | None:
    unique_values = []
    for value in values:
        if _is_missing(value):
            continue
        text = str(value)
        if text not in unique_values:
            unique_values.append(text)
    return "; ".join(unique_values) if unique_values else None


def _count_distinct_files(records: list[dict[str, Any]]) -> int:
    names = {
        record.get("source_file") or record.get("file_name")
        for record in records
        if record.get("source_file") or record.get("file_name")
    }
    return len(names) if names else len(records)


def _safe_int(value: Any) -> int:
    if _is_missing(value):
        return 0
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _none_if_missing(value: Any) -> Any:
    if _is_missing(value):
        return None
    return value


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd is not None and pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _require_pandas() -> Any:
    if pd is None:
        raise RuntimeError(
            "Le module pandas est requis pour rapprocher les donnees. "
            "Installez les dependances avec: python -m pip install -r requirements.txt"
        )
    return pd
