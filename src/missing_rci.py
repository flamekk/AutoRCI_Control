from __future__ import annotations

import csv
import math
from pathlib import Path
from typing import Any, Iterable


CATEGORY_FACTURE = "Facture absente"
CATEGORY_AVOIR = "Avoir absent"
CATEGORY_UNKNOWN = "Non déterminé"
NO_MISSING_RCI_MESSAGE = "Aucune facture ou avoir absent côté RCI pour cette période."

ABSENT_RCI_COLUMNS = [
    "categorie",
    "severity",
    "invoice_number",
    "document_type",
    "erp_date",
    "due_date",
    "customer_code",
    "customer_name",
    "amount_erp",
    "montant_impacte",
    "status",
    "action_recommandee",
    "included_in_corrective_batch",
    "source_erp",
    "source_rci",
    "source_pdf",
    "commentaire_audit",
]

ABSENT_RCI_HEADERS = {
    "categorie": "Catégorie",
    "severity": "Sévérité",
    "invoice_number": "N° facture / avoir",
    "document_type": "Type document",
    "erp_date": "Date facture ERP",
    "due_date": "Date échéance",
    "customer_code": "Code concessionnaire",
    "customer_name": "Concessionnaire",
    "amount_erp": "Montant ERP",
    "montant_impacte": "Montant impacté",
    "status": "Statut",
    "action_recommandee": "Action recommandée",
    "included_in_corrective_batch": "Inclus batch correctif",
    "source_erp": "Source ERP",
    "source_rci": "Source RCI",
    "source_pdf": "Source PDF",
    "commentaire_audit": "Commentaire audit",
}

CATEGORY_ORDER = {
    CATEGORY_FACTURE: 0,
    CATEGORY_AVOIR: 1,
    CATEGORY_UNKNOWN: 2,
}

ABSENT_SEVERITY_ORDER = {
    "CRITIQUE": 0,
    "ELEVEE": 1,
    "MOYENNE": 2,
}


def determine_absent_category(record: dict[str, Any]) -> str:
    document_type = str(record.get("document_type") or "").strip().upper()
    invoice_number = str(record.get("invoice_number") or "").strip().upper().replace(" ", "")

    if document_type == "FACTURE":
        return CATEGORY_FACTURE
    if document_type == "AVOIR":
        return CATEGORY_AVOIR

    if invoice_number.startswith(("FVF", "VF")):
        return CATEGORY_FACTURE
    if invoice_number.startswith("AAF"):
        return CATEGORY_AVOIR
    return CATEGORY_UNKNOWN


def build_missing_rci_records(records: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    missing_rows: list[dict[str, Any]] = []
    for record in records:
        if record.get("status") != "MANQUANTE_RCI":
            continue
        enriched = {column: record.get(column, "") for column in ABSENT_RCI_COLUMNS}
        enriched["categorie"] = determine_absent_category(record)
        enriched["commentaire_audit"] = (
            record.get("commentaire_audit")
            or "Présent dans le batch ERP mais absent du batch RCI."
        )
        enriched["included_in_corrective_batch"] = _as_bool(record.get("included_in_corrective_batch", False))
        missing_rows.append(enriched)

    return sorted(
        missing_rows,
        key=lambda row: (
            CATEGORY_ORDER.get(str(row.get("categorie") or ""), 99),
            ABSENT_SEVERITY_ORDER.get(str(row.get("severity") or ""), 99),
            -abs(_number(row.get("montant_impacte")) or 0.0),
            str(row.get("invoice_number") or ""),
        ),
    )


def missing_rci_summary(records: Iterable[dict[str, Any]]) -> dict[str, Any]:
    rows = build_missing_rci_records(records)
    invoice_count = sum(1 for row in rows if row.get("categorie") == CATEGORY_FACTURE)
    credit_note_count = sum(1 for row in rows if row.get("categorie") == CATEGORY_AVOIR)
    unknown_count = sum(1 for row in rows if row.get("categorie") == CATEGORY_UNKNOWN)
    total_amount = round(sum(abs(_number(row.get("montant_impacte")) or 0.0) for row in rows), 2)
    severity_counts: dict[str, int] = {"CRITIQUE": 0, "ELEVEE": 0, "MOYENNE": 0}
    for row in rows:
        severity = str(row.get("severity") or "")
        if severity in severity_counts:
            severity_counts[severity] += 1

    return {
        "missing_rci_total_count": len(rows),
        "missing_rci_invoice_count": invoice_count,
        "missing_rci_credit_note_count": credit_note_count,
        "missing_rci_unknown_count": unknown_count,
        "missing_rci_total_amount": total_amount,
        "total_absent_rci_amount": total_amount,
        "missing_rci_critical_count": severity_counts["CRITIQUE"],
        "missing_rci_high_count": severity_counts["ELEVEE"],
        "missing_rci_medium_count": severity_counts["MOYENNE"],
    }


def enrich_report_with_missing_rci(report: dict[str, Any]) -> dict[str, Any]:
    rows = build_missing_rci_records(report.get("reconciliation", []))
    report["missing_rci_records"] = rows
    report.setdefault("summary", {}).update(missing_rci_summary(report.get("reconciliation", [])))
    return report


def write_missing_rci_export(report: dict[str, Any], output_dir: Path, run_id: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    enrich_report_with_missing_rci(report)
    rows = report.get("missing_rci_records", [])
    path = output_dir / f"factures_avoirs_absents_RCI_{run_id}.csv"

    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(
            stream,
            fieldnames=[ABSENT_RCI_HEADERS[column] for column in ABSENT_RCI_COLUMNS],
            delimiter=";",
        )
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    ABSENT_RCI_HEADERS[column]: _csv_value(row.get(column))
                    for column in ABSENT_RCI_COLUMNS
                }
            )
    return path


def _csv_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value or "").strip().lower()
    return text in {"1", "true", "yes", "oui", "y", "on"}


def _number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return number
