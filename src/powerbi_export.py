from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Any


HISTORY_FILENAME = "reconciliation_history.csv"
HISTORY_COLUMNS = [
    "processing_date",
    "processing_run_id",
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
    "pdf_invoice_date",
    "due_date",
    "origin",
    "status",
    "priority",
    "action_recommandee",
]


def write_powerbi_exports(report: dict[str, Any], output_dir: Path, run_id: str) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)

    summary_path = output_dir / f"run_summary_{run_id}.csv"
    files_path = output_dir / f"source_files_{run_id}.csv"
    reconciliation_path = output_dir / f"reconciliation_{run_id}.csv"
    history_path = output_dir / HISTORY_FILENAME

    with summary_path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream, delimiter=";")
        writer.writerow(["metric", "value"])
        writer.writerow(["status", report["status"]])
        for key, value in report["summary"].items():
            writer.writerow([key, value])

    with files_path.open("w", encoding="utf-8-sig", newline="") as stream:
        fieldnames = _source_fieldnames(report["source_files"])
        writer = csv.DictWriter(stream, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        for record in report["source_files"]:
            writer.writerow({field: record.get(field, "") for field in fieldnames})

    with reconciliation_path.open("w", encoding="utf-8-sig", newline="") as stream:
        fieldnames = _reconciliation_fieldnames(report.get("reconciliation", []))
        writer = csv.DictWriter(stream, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        for record in report.get("reconciliation", []):
            writer.writerow({field: record.get(field, "") for field in fieldnames})

    update_reconciliation_history(
        report.get("reconciliation", []),
        history_path,
        run_id,
        processing_date=_processing_date(report, run_id),
    )

    return [summary_path, files_path, reconciliation_path, history_path]


def update_reconciliation_history(
    reconciliation_records: list[dict[str, Any]],
    history_path: Path,
    run_id: str,
    processing_date: str | None = None,
) -> Path:
    history_path.parent.mkdir(parents=True, exist_ok=True)
    processing_date = processing_date or _processing_date({}, run_id)

    existing_rows = _read_history_rows(history_path)
    new_rows = [
        _history_row(record, processing_date=processing_date, run_id=run_id)
        for record in reconciliation_records
    ]
    rows = _deduplicate_rows([*existing_rows, *new_rows])

    with history_path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=HISTORY_COLUMNS, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)

    return history_path


def _read_history_rows(history_path: Path) -> list[dict[str, Any]]:
    if not history_path.exists():
        return []

    with history_path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream, delimiter=";")
        return [
            {column: row.get(column, "") for column in HISTORY_COLUMNS}
            for row in reader
        ]


def _history_row(record: dict[str, Any], processing_date: str, run_id: str) -> dict[str, Any]:
    row = {
        "processing_date": processing_date,
        "processing_run_id": run_id,
    }
    for column in HISTORY_COLUMNS:
        if column in row:
            continue
        row[column] = _csv_value(record.get(column, ""))
    return row


def _deduplicate_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    unique_rows = []
    seen = set()
    for row in rows:
        normalized_row = {column: _csv_value(row.get(column, "")) for column in HISTORY_COLUMNS}
        key = tuple(normalized_row[column] for column in HISTORY_COLUMNS)
        if key in seen:
            continue
        seen.add(key)
        unique_rows.append(normalized_row)
    return unique_rows


def _processing_date(report: dict[str, Any], run_id: str) -> str:
    generated_at = report.get("generated_at")
    if generated_at:
        try:
            return datetime.fromisoformat(str(generated_at).replace("Z", "+00:00")).date().isoformat()
        except ValueError:
            pass

    for date_format in ("%Y%m%d_%H%M%S", "%Y%m%d_%H%M"):
        try:
            return datetime.strptime(run_id, date_format).date().isoformat()
        except ValueError:
            continue
    return datetime.now().date().isoformat()


def _csv_value(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _source_fieldnames(records: list[dict[str, Any]]) -> list[str]:
    preferred = [
        "source_type",
        "source_file",
        "file_name",
        "source_sheet",
        "page",
        "invoice_number",
        "document_type",
        "erp_date",
        "customer_code",
        "customer_name",
        "is_rci_covered",
        "amount_erp",
        "department_code",
        "brand_code",
        "sales_order_type",
        "rci_date",
        "dealer_code",
        "amount_rci",
        "cf_code",
        "pdf_invoice_date",
        "due_date",
        "model",
        "chassis_number",
        "amount_pdf",
        "origin",
        "raw_line",
        "file_path",
        "extension",
        "size_bytes",
        "modified_at",
        "status",
        "ingested_at",
    ]
    discovered = sorted({key for record in records for key in record})
    if not discovered:
        return preferred
    return [field for field in preferred if field in discovered] + [
        field for field in discovered if field not in preferred
    ]


def _reconciliation_fieldnames(records: list[dict[str, Any]]) -> list[str]:
    preferred = [
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
    discovered = sorted({key for record in records for key in record})
    if not discovered:
        return preferred
    return [field for field in preferred if field in discovered] + [
        field for field in discovered if field not in preferred
    ]
