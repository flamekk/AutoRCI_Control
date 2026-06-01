from __future__ import annotations

import logging
import math
from collections import Counter, defaultdict
from datetime import date
from difflib import SequenceMatcher
from typing import Any, Iterable

try:
    from src.normalize import normalize_date
    from src.reference_loader import normalize_reference_name
except ModuleNotFoundError:  # pragma: no cover - used when running python src/main.py.
    from normalize import normalize_date
    from reference_loader import normalize_reference_name


LOGGER = logging.getLogger(__name__)

IN_SCOPE_ERP_STATUSES = {
    "OK",
    "MANQUANTE_RCI",
    "ANOMALIE_MONTANT",
    "ANOMALIE_DATE",
    "DOUBLON",
}
ERP_ANALYSIS_STATUSES = {*IN_SCOPE_ERP_STATUSES, "HORS_SCOPE_RCI"}


def enrich_report_with_audits(
    report: dict[str, Any],
    reference_names: Iterable[Any] | None = None,
    *,
    log: bool = True,
) -> dict[str, Any]:
    """Attach business audit views and alert indicators to a reconciliation report."""

    reconciliation = [dict(record) for record in report.get("reconciliation", [])]
    summary = report.setdefault("summary", {})
    start_date, end_date = _period_bounds(summary)
    normalized_references = _normalized_reference_names(reference_names)

    audits = {
        "dates": build_audit_dates(reconciliation),
        "missing_rci": build_audit_missing_rci(reconciliation, start_date, end_date),
        "out_of_scope_rci": build_audit_out_of_scope_rci(reconciliation, normalized_references),
    }
    report["audits"] = audits
    _update_summary_alerts(report, audits)

    if log:
        log_audit_findings(report)
    return report


def build_audit_dates(records: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[str, dict[str, Any]] = defaultdict(_empty_date_group)

    for record in records:
        status = str(record.get("status") or "")
        if status in ERP_ANALYSIS_STATUSES:
            erp_date = normalize_date(record.get("erp_date")) or "DATE_NON_RENSEIGNEE"
            group = groups[erp_date]
            group["erp_date"] = erp_date
            group["nombre_total_factures_erp"] += 1

            if status == "OK":
                group["nombre_ok"] += 1
            elif status == "MANQUANTE_RCI":
                group["nombre_manquante_rci"] += 1
                group["montant_manquant_rci"] += _abs_number(record.get("montant_impacte") or record.get("amount_erp"))
            elif status == "HORS_SCOPE_RCI":
                group["nombre_hors_scope_rci"] += 1

        rci_date = normalize_date(record.get("rci_date") or record.get("pdf_invoice_date"))
        if rci_date and status == "RCI_HORS_PERIODE":
            group = groups[rci_date]
            group["erp_date"] = rci_date
            group["rci_hors_periode_count"] += 1
        elif rci_date and (record.get("source_rci") or record.get("amount_rci") is not None):
            group = groups[rci_date]
            group["erp_date"] = rci_date
            group["rci_count"] += 1

    rows = []
    for erp_date, group in sorted(groups.items(), key=lambda item: item[0]):
        in_scope_total = max(group["nombre_total_factures_erp"] - group["nombre_hors_scope_rci"], 0)
        group["montant_manquant_rci"] = round(group["montant_manquant_rci"], 2)
        group["taux_rapprochement_date"] = (
            round(group["nombre_ok"] / in_scope_total, 4) if in_scope_total else 0.0
        )
        rows.append(dict(group))
    return rows


def build_audit_missing_rci(
    records: Iterable[dict[str, Any]],
    start_date: date | None,
    end_date: date | None,
) -> list[dict[str, Any]]:
    rows = []
    for record in records:
        if record.get("status") != "MANQUANTE_RCI":
            continue

        erp_date = normalize_date(record.get("erp_date"))
        date_in_filter = _date_in_filter(erp_date, start_date, end_date)
        rows.append(
            {
                "invoice_number": record.get("invoice_number"),
                "erp_date": erp_date,
                "customer_name": record.get("customer_name"),
                "customer_code": record.get("customer_code"),
                "amount_erp": _number(record.get("amount_erp")),
                "is_rci_covered": record.get("is_rci_covered"),
                "status": record.get("status"),
                "montant_impacte": _number(record.get("montant_impacte")),
                "source_erp": record.get("source_erp"),
                "date_in_filter": date_in_filter,
                "commentaire_audit": (
                    "Dans la période filtrée"
                    if date_in_filter
                    else "Hors période filtrée - vérifier filtre date"
                ),
            }
        )
    return rows


def build_audit_out_of_scope_rci(
    records: Iterable[dict[str, Any]],
    reference_names: Iterable[str] | None,
) -> list[dict[str, Any]]:
    normalized_references = list(reference_names or [])
    rows = []
    for record in records:
        if record.get("status") != "HORS_SCOPE_RCI":
            continue

        normalized_customer_name = normalize_reference_name(record.get("customer_name"))
        closest_name, similarity = _closest_reference_name(normalized_customer_name, normalized_references)
        rows.append(
            {
                "invoice_number": record.get("invoice_number"),
                "erp_date": normalize_date(record.get("erp_date")),
                "customer_name": record.get("customer_name"),
                "normalized_customer_name": normalized_customer_name,
                "amount_erp": _number(record.get("amount_erp")),
                "status": record.get("status"),
                "closest_reference_name": closest_name,
                "closest_reference_similarity": similarity,
                "commentaire_audit": _out_of_scope_comment(closest_name, similarity),
            }
        )
    return rows


def log_audit_findings(report: dict[str, Any]) -> None:
    audits = report.get("audits", {})
    dates = audits.get("dates", [])
    missing_rows = audits.get("missing_rci", [])
    out_scope_rows = audits.get("out_of_scope_rci", [])
    summary = report.get("summary", {})

    LOGGER.info(
        "Audit dates - repartition OK par date: %s",
        _format_count_mapping({row["erp_date"]: row["nombre_ok"] for row in dates if row["nombre_ok"]}),
    )
    LOGGER.info(
        "Audit dates - repartition MANQUANTE_RCI par date: %s",
        _format_count_mapping(
            {row["erp_date"]: row["nombre_manquante_rci"] for row in dates if row["nombre_manquante_rci"]}
        ),
    )
    LOGGER.info(
        "Audit dates - repartition RCI_HORS_PERIODE par date: %s",
        _format_count_mapping(
            {row["erp_date"]: row["rci_hors_periode_count"] for row in dates if row["rci_hors_periode_count"]}
        ),
    )

    out_scope_by_client = Counter(str(row.get("customer_name") or "NON RENSEIGNE") for row in out_scope_rows)
    LOGGER.info(
        "Audit hors scope RCI - repartition par client: %s",
        _format_count_mapping(dict(out_scope_by_client.most_common(20))),
    )
    LOGGER.info(
        "Audit manquantes RCI hors periode: %s",
        summary.get("missing_rci_out_of_period", 0),
    )

    if summary.get("missing_rci_out_of_period_alert"):
        LOGGER.warning(
            "Alerte audit: %s MANQUANTE_RCI ont une date ERP hors periode filtree.",
            summary.get("missing_rci_out_of_period", 0),
        )
    if summary.get("low_matching_rate_alert"):
        LOGGER.warning(
            "Alerte audit: taux de rapprochement inferieur a 70%% (%.2f%%).",
            float(summary.get("matching_rate", 0) or 0) * 100,
        )
    if summary.get("out_of_scope_rate_alert"):
        LOGGER.warning(
            "Alerte audit: plus de 20%% des factures sont HORS_SCOPE_RCI (%.2f%%).",
            float(summary.get("out_of_scope_rci_percent", 0) or 0) * 100,
        )
    if summary.get("no_rci_flux_in_period_alert"):
        LOGGER.warning("Attention : aucun flux RCI dans la période de rapprochement.")


def _update_summary_alerts(report: dict[str, Any], audits: dict[str, list[dict[str, Any]]]) -> None:
    summary = report.setdefault("summary", {})
    reconciliation = list(report.get("reconciliation", []))
    out_of_scope_count = int(summary.get("out_of_scope_rci", _count_status(reconciliation, "HORS_SCOPE_RCI")) or 0)
    matchable_count = int(
        summary.get("erp_matchable_invoices", _count_statuses(reconciliation, IN_SCOPE_ERP_STATUSES)) or 0
    )
    ok_count = int(summary.get("matched_invoices", _count_status(reconciliation, "OK")) or 0)
    matching_rate = float(summary.get("matching_rate", ok_count / matchable_count if matchable_count else 0) or 0)
    erp_analyzed_total = matchable_count + out_of_scope_count
    out_of_scope_percent = out_of_scope_count / erp_analyzed_total if erp_analyzed_total else 0.0
    missing_out_of_period = sum(1 for row in audits.get("missing_rci", []) if row.get("date_in_filter") is False)

    summary["matching_rate"] = round(matching_rate, 4)
    summary["missing_rci_out_of_period"] = missing_out_of_period
    summary["out_of_scope_rci_percent"] = round(out_of_scope_percent, 4)
    summary["low_matching_rate_alert"] = bool(matchable_count and matching_rate < 0.70)
    summary["missing_rci_out_of_period_alert"] = missing_out_of_period > 0
    summary["out_of_scope_rate_alert"] = out_of_scope_percent > 0.20
    summary["audit_alerts"] = _audit_alerts(summary)


def _audit_alerts(summary: dict[str, Any]) -> list[str]:
    alerts = []
    if summary.get("missing_rci_out_of_period_alert"):
        alerts.append("Des MANQUANTE_RCI ont une date ERP hors période filtrée.")
    if summary.get("low_matching_rate_alert"):
        alerts.append("Le taux de rapprochement est inférieur à 70%.")
    if summary.get("out_of_scope_rate_alert"):
        alerts.append("Plus de 20% des factures sont hors périmètre RCI.")
    if summary.get("no_rci_flux_in_period_alert"):
        alerts.append("Attention : aucun flux RCI dans la période de rapprochement.")
    return alerts


def _empty_date_group() -> dict[str, Any]:
    return {
        "erp_date": None,
        "nombre_total_factures_erp": 0,
        "nombre_ok": 0,
        "nombre_manquante_rci": 0,
        "nombre_hors_scope_rci": 0,
        "montant_manquant_rci": 0.0,
        "taux_rapprochement_date": 0.0,
        "rci_count": 0,
        "rci_hors_periode_count": 0,
    }


def _period_bounds(summary: dict[str, Any]) -> tuple[date | None, date | None]:
    start_date = _parse_iso_date(summary.get("reconciliation_start_date"))
    end_date = _parse_iso_date(summary.get("reconciliation_end_date"))
    return start_date, end_date


def _parse_iso_date(value: Any) -> date | None:
    normalized = normalize_date(value)
    if normalized is None:
        return None
    try:
        return date.fromisoformat(normalized)
    except ValueError:
        return None


def _date_in_filter(value: str | None, start_date: date | None, end_date: date | None) -> bool:
    if start_date is None or end_date is None:
        return True
    if value is None:
        return False
    try:
        parsed = date.fromisoformat(value)
    except ValueError:
        return False
    return start_date <= parsed <= end_date


def _normalized_reference_names(reference_names: Iterable[Any] | None) -> list[str]:
    names = {
        normalized
        for value in (reference_names or [])
        if (normalized := normalize_reference_name(value))
    }
    return sorted(names)


def _closest_reference_name(customer_name: str, reference_names: list[str]) -> tuple[str | None, float | None]:
    if not customer_name or not reference_names:
        return None, None

    best_name = None
    best_score = 0.0
    for reference_name in reference_names:
        score = SequenceMatcher(None, customer_name, reference_name).ratio()
        if score > best_score:
            best_name = reference_name
            best_score = score
    return best_name, round(best_score, 4) if best_name is not None else None


def _out_of_scope_comment(closest_name: str | None, similarity: float | None) -> str:
    if closest_name is None or similarity is None:
        return "Référentiel RCI indisponible pour comparaison"
    if similarity >= 0.80:
        return "Nom proche du référentiel RCI - vérifier une différence de libellé"
    return "Aucun reclassement automatique - validation métier nécessaire"


def _count_status(records: Iterable[dict[str, Any]], status: str) -> int:
    return sum(1 for record in records if record.get("status") == status)


def _count_statuses(records: Iterable[dict[str, Any]], statuses: set[str]) -> int:
    return sum(1 for record in records if record.get("status") in statuses)


def _number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return round(number, 2)


def _abs_number(value: Any) -> float:
    number = _number(value)
    return abs(number) if number is not None else 0.0


def _format_count_mapping(values: dict[str, int]) -> str:
    if not values:
        return "(aucune)"
    return ", ".join(f"{key}={value}" for key, value in values.items())
