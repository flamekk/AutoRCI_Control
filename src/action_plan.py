from __future__ import annotations

import csv
import math
from pathlib import Path
from typing import Any, Iterable

try:
    from src.reference_loader import normalize_reference_name
except ModuleNotFoundError:  # pragma: no cover - used when running python src/main.py.
    from reference_loader import normalize_reference_name


SEVERITY_ORDER = {
    "CRITIQUE": 0,
    "ELEVEE": 1,
    "MOYENNE": 2,
    "A_VERIFIER": 3,
    "INFORMATION": 4,
    "OK": 5,
}

STATUS_ACTIONS = {
    "OK": "Aucune action requise.",
    "MANQUANTE_RCI": "Vérifier la transmission vers RCI et préparer un renvoi si nécessaire.",
    "HORS_SCOPE_RCI": "Vérifier si le concessionnaire doit être ajouté au référentiel RCI.",
    "RCI_HORS_PERIODE": "Contrôler la période du fichier RCI/PDF chargé.",
    "ANOMALIE_MONTANT": "Comparer le montant ERP avec le montant RCI.",
    "ANOMALIE_DATE": "Comparer la date ERP avec la date RCI/PDF.",
    "DOUBLON": "Vérifier les doublons dans les fichiers source.",
    "RCI_SEULEMENT": "Vérifier origine côté RCI ou historique ERP.",
}

REFERENCE_SUGGESTION_COLUMNS = [
    "customer_name",
    "nombre_factures",
    "montant_total_erp",
    "closest_reference_name",
    "closest_reference_similarity",
    "suggestion_action",
]


def recommended_action(status: Any) -> str:
    return STATUS_ACTIONS.get(str(status or ""), "Analyser l'écart et définir une action de régularisation.")


def severity_for_status(status: Any, montant_impacte: Any = None) -> str:
    normalized_status = str(status or "").strip().upper()
    impacted_amount = abs(_number(montant_impacte) or 0.0)

    if normalized_status == "OK":
        return "OK"
    if normalized_status == "MANQUANTE_RCI":
        if impacted_amount >= 100000:
            return "CRITIQUE"
        if impacted_amount >= 20000:
            return "ELEVEE"
        return "MOYENNE"
    if normalized_status == "HORS_SCOPE_RCI":
        return "A_VERIFIER"
    if normalized_status == "RCI_HORS_PERIODE":
        return "INFORMATION"
    if normalized_status in {"ANOMALIE_MONTANT", "DOUBLON"}:
        return "ELEVEE"
    if normalized_status in {"ANOMALIE_DATE", "RCI_SEULEMENT"}:
        return "MOYENNE"
    return "A_VERIFIER"


def sort_action_plan_records(records: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    actionable = []
    for record in records:
        enriched = dict(record)
        severity = enriched.get("severity") or severity_for_status(
            enriched.get("status"),
            enriched.get("montant_impacte"),
        )
        if severity == "OK":
            continue
        enriched["severity"] = severity
        enriched["action_recommandee"] = enriched.get("action_recommandee") or recommended_action(enriched.get("status"))
        actionable.append(enriched)

    return sorted(
        actionable,
        key=lambda row: (
            SEVERITY_ORDER.get(str(row.get("severity") or ""), 99),
            -abs(_number(row.get("montant_impacte")) or 0.0),
            str(row.get("invoice_number") or ""),
        ),
    )


def severity_counts(records: Iterable[dict[str, Any]]) -> dict[str, int]:
    counts = {severity: 0 for severity in SEVERITY_ORDER}
    for record in records:
        severity = str(
            record.get("severity")
            or severity_for_status(record.get("status"), record.get("montant_impacte"))
        )
        counts[severity] = counts.get(severity, 0) + 1
    return counts


def build_reference_suggestions(records: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[str, dict[str, Any]] = {}
    for record in records:
        if record.get("status") != "HORS_SCOPE_RCI":
            continue

        customer_name = str(record.get("customer_name") or "NON RENSEIGNE").strip() or "NON RENSEIGNE"
        normalized_name = normalize_reference_name(customer_name) or "NON RENSEIGNE"
        if normalized_name not in groups:
            groups[normalized_name] = {
                "customer_name": customer_name,
                "nombre_factures": 0,
                "montant_total_erp": 0.0,
                "sample_invoice_numbers": [],
                "closest_reference_name": record.get("closest_reference_name") or "",
                "closest_reference_similarity": record.get("closest_reference_similarity") or "",
                "suggestion_action": "Vérifier si le concessionnaire doit être ajouté au référentiel RCI.",
            }
        group = groups[normalized_name]
        group["nombre_factures"] += 1
        group["montant_total_erp"] += abs(_number(record.get("amount_erp")) or 0.0)
        invoice_number = str(record.get("invoice_number") or "").strip()
        if invoice_number and len(group["sample_invoice_numbers"]) < 10:
            group["sample_invoice_numbers"].append(invoice_number)
        if record.get("closest_reference_name") and not group.get("closest_reference_name"):
            group["closest_reference_name"] = record.get("closest_reference_name")
            group["closest_reference_similarity"] = record.get("closest_reference_similarity") or ""

    rows = []
    for group in groups.values():
        row = dict(group)
        row["montant_total_erp"] = round(float(row["montant_total_erp"]), 2)
        row.pop("sample_invoice_numbers", None)
        rows.append(row)

    return sorted(rows, key=lambda row: (-float(row["montant_total_erp"] or 0), str(row["customer_name"])))


def write_reference_suggestions(records: Iterable[dict[str, Any]], output_dir: Path, run_id: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"reference_suggestions_{run_id}.csv"
    rows = build_reference_suggestions(records)
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=REFERENCE_SUGGESTION_COLUMNS, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)
    return path


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
