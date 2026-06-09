from __future__ import annotations

import csv
import math
from pathlib import Path
from typing import Any, Iterable

try:
    from src.action_plan import severity_for_status
    from src.missing_rci import enrich_report_with_missing_rci
except ModuleNotFoundError:  # pragma: no cover - used when running python src/main.py.
    from action_plan import severity_for_status
    from missing_rci import enrich_report_with_missing_rci


CORRECTIVE_BATCH_SEVERITIES = {"CRITIQUE", "ELEVEE"}
NO_CORRECTIVE_BATCH_MESSAGE = (
    "Aucun batch correctif candidat généré : aucune facture ou avoir prioritaire absent côté RCI."
)
CORRECTIVE_BATCH_WARNING = "Fichier candidat à valider par l’équipe facturation avant transmission à RCI."

CONTROL_COLUMNS = [
    "invoice_number",
    "erp_date",
    "customer_name",
    "amount_erp",
    "montant_impacte",
    "severity",
    "status",
    "action_recommandee",
    "included_in_corrective_batch",
]

TXT_COLUMNS = [
    "invoice_number",
    "erp_date",
    "customer_name",
    "amount_erp",
    "montant_impacte",
    "severity",
]


def is_corrective_batch_candidate(record: dict[str, Any]) -> bool:
    status = str(record.get("status") or "")
    severity = str(record.get("severity") or severity_for_status(status, record.get("montant_impacte")))
    return status == "MANQUANTE_RCI" and severity in CORRECTIVE_BATCH_SEVERITIES


def mark_corrective_batch_records(records: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    marked = []
    for record in records:
        enriched = dict(record)
        enriched["included_in_corrective_batch"] = is_corrective_batch_candidate(enriched)
        marked.append(enriched)
    return marked


def select_corrective_batch_records(records: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    candidates = []
    for record in mark_corrective_batch_records(records):
        if record.get("included_in_corrective_batch") is True:
            candidates.append(record)
    return sorted(
        candidates,
        key=lambda row: (
            0 if row.get("severity") == "CRITIQUE" else 1,
            -abs(_number(row.get("montant_impacte")) or 0.0),
            str(row.get("invoice_number") or ""),
        ),
    )


def write_corrective_batch_outputs(report: dict[str, Any], output_dir: Path, run_id: str) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    marked_records = mark_corrective_batch_records(report.get("reconciliation", []))
    report["reconciliation"] = marked_records
    enrich_report_with_missing_rci(report)
    candidates = select_corrective_batch_records(report.get("missing_rci_records", []))

    if not candidates:
        _attach_batch_metadata(report, [], None, None)
        return []

    batch_path = output_dir / f"batch_correctif_candidat_{run_id}.txt"
    control_path = output_dir / f"batch_correctif_candidat_{run_id}_control.csv"

    _write_batch_txt(candidates, batch_path)
    _write_control_csv(candidates, control_path)
    _attach_batch_metadata(report, candidates, batch_path, control_path)
    return [batch_path, control_path]


def _write_batch_txt(records: list[dict[str, Any]], path: Path) -> None:
    with path.open("w", encoding="utf-8", newline="") as stream:
        stream.write(";".join(TXT_COLUMNS))
        stream.write("\n")
        for record in records:
            stream.write(";".join(_txt_value(record.get(column)) for column in TXT_COLUMNS))
            stream.write("\n")


def _write_control_csv(records: list[dict[str, Any]], path: Path) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=CONTROL_COLUMNS, delimiter=";")
        writer.writeheader()
        for record in records:
            writer.writerow(
                {
                    column: (
                        "true"
                        if column == "included_in_corrective_batch"
                        else _csv_value(record.get(column))
                    )
                    for column in CONTROL_COLUMNS
                }
            )


def _attach_batch_metadata(
    report: dict[str, Any],
    records: list[dict[str, Any]],
    batch_path: Path | None,
    control_path: Path | None,
) -> None:
    total_amount = round(sum(abs(_number(record.get("montant_impacte")) or 0.0) for record in records), 2)
    generated = bool(records)
    metadata = {
        "generated": generated,
        "batch_path": str(batch_path) if batch_path else "",
        "control_path": str(control_path) if control_path else "",
        "invoice_count": len(records),
        "total_amount": total_amount,
        "warning": CORRECTIVE_BATCH_WARNING if generated else NO_CORRECTIVE_BATCH_MESSAGE,
        "invoices": [record.get("invoice_number") for record in records if record.get("invoice_number")],
        "records": records,
    }
    report["corrective_batch"] = metadata
    summary = report.setdefault("summary", {})
    summary.update(
        {
            "corrective_batch_generated": generated,
            "corrective_batch_path": str(batch_path) if batch_path else "",
            "corrective_batch_control_path": str(control_path) if control_path else "",
            "corrective_batch_invoice_count": len(records),
            "corrective_batch_total_amount": total_amount,
        }
    )


def _csv_value(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _txt_value(value: Any) -> str:
    text = _csv_value(value)
    return text.replace("\n", " ").replace("\r", " ").replace(";", ",")


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
