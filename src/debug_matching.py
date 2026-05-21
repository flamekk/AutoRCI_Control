from __future__ import annotations

import csv
import json
import logging
from pathlib import Path
from typing import Any

try:
    import pandas as pd
except ModuleNotFoundError:  # pragma: no cover - exercised only before dependency install.
    pd = None  # type: ignore[assignment]

try:
    from src.normalize import normalize_invoice_number
    from src.reconcile import _aggregate_pdf, _aggregate_rci, _consolidate_rci_pdf
except ModuleNotFoundError:  # pragma: no cover - used when running python src/main.py.
    from normalize import normalize_invoice_number
    from reconcile import _aggregate_pdf, _aggregate_rci, _consolidate_rci_pdf


LOGGER = logging.getLogger(__name__)

DEFAULT_AUDIT_INVOICES = ["VF384312", "VF384313", "VF384314", "VF384317", "VF384318"]

MATCHING_DEBUG_COLUMNS = [
    "total_erp_rows",
    "total_erp_unique_invoices",
    "total_erp_unique_invoice_numbers",
    "total_rci_txt_rows",
    "total_rci_txt_unique_invoices",
    "total_rci_rows",
    "total_rci_unique_invoice_numbers",
    "total_pdf_rows",
    "total_pdf_unique_invoices",
    "intersection_erp_rci_txt",
    "intersection_erp_rci",
    "intersection_erp_pdf",
    "intersection_erp_rci_consolidated",
    "only_erp_count",
    "only_rci_count",
    "sample_common_erp_rci_txt",
    "sample_common_invoices",
    "sample_only_erp",
    "sample_only_rci_txt",
    "sample_only_rci",
    "sample_only_pdf",
]


def write_matching_debug(
    erp_records: Any,
    rci_records: Any,
    pdf_records: Any,
    output_dir: Path,
    run_id: str,
) -> Path:
    """Write a one-line CSV audit of invoice-number matching between sources."""

    erp = _to_dataframe(erp_records)
    rci = _to_dataframe(rci_records)
    pdf = _to_dataframe(pdf_records)

    erp_invoices = _invoice_set(erp)
    rci_invoices = _invoice_set(rci)
    pdf_invoices = _invoice_set(pdf)
    consolidated_invoices = rci_invoices | pdf_invoices

    common_erp_rci = erp_invoices & rci_invoices
    common_erp_pdf = erp_invoices & pdf_invoices
    common_erp_consolidated = erp_invoices & consolidated_invoices
    only_erp = erp_invoices - consolidated_invoices
    only_rci = rci_invoices - erp_invoices
    only_pdf = pdf_invoices - erp_invoices

    row = {
        "total_erp_rows": len(erp),
        "total_erp_unique_invoices": len(erp_invoices),
        "total_erp_unique_invoice_numbers": len(erp_invoices),
        "total_rci_txt_rows": len(rci),
        "total_rci_txt_unique_invoices": len(rci_invoices),
        "total_rci_rows": len(rci),
        "total_rci_unique_invoice_numbers": len(rci_invoices),
        "total_pdf_rows": len(pdf),
        "total_pdf_unique_invoices": len(pdf_invoices),
        "intersection_erp_rci_txt": len(common_erp_rci),
        "intersection_erp_rci": len(common_erp_rci),
        "intersection_erp_pdf": len(common_erp_pdf),
        "intersection_erp_rci_consolidated": len(common_erp_consolidated),
        "only_erp_count": len(only_erp),
        "only_rci_count": len(only_rci),
        "sample_common_erp_rci_txt": _sample(common_erp_rci),
        "sample_common_invoices": _sample(common_erp_rci),
        "sample_only_erp": _sample(only_erp),
        "sample_only_rci_txt": _sample(only_rci),
        "sample_only_rci": _sample(only_rci),
        "sample_only_pdf": _sample(only_pdf),
    }

    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"matching_debug_{run_id}.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=MATCHING_DEBUG_COLUMNS, delimiter=";")
        writer.writeheader()
        writer.writerow(row)

    LOGGER.info(
        "Audit matching: ERP uniques=%s, RCI uniques=%s, PDF uniques=%s, intersection ERP/RCI=%s, ERP seul=%s, RCI seul=%s.",
        row["total_erp_unique_invoice_numbers"],
        row["total_rci_unique_invoice_numbers"],
        row["total_pdf_unique_invoices"],
        row["intersection_erp_rci"],
        row["only_erp_count"],
        row["only_rci_count"],
    )
    LOGGER.info("Audit matching genere: %s", path)
    return path


def log_invoice_presence(erp_records: Any, rci_records: Any, pdf_records: Any, invoices: list[str] | None = None) -> None:
    erp_invoices = _invoice_set(_to_dataframe(erp_records))
    rci_invoices = _invoice_set(_to_dataframe(rci_records))
    pdf_invoices = _invoice_set(_to_dataframe(pdf_records))

    for raw_invoice in invoices or DEFAULT_AUDIT_INVOICES:
        invoice = normalize_invoice_number(raw_invoice) or str(raw_invoice).strip().upper()
        LOGGER.info(
            "Presence facture %s: ERP=%s, RCI TXT=%s, PDF=%s",
            invoice,
            "oui" if invoice in erp_invoices else "non",
            "oui" if invoice in rci_invoices else "non",
            "oui" if invoice in pdf_invoices else "non",
        )


def log_debug_invoice(
    invoice_number: str,
    erp_records: Any,
    rci_records: Any,
    pdf_records: Any,
    reconciliation_records: Any,
) -> None:
    invoice = normalize_invoice_number(invoice_number)
    if invoice is None:
        LOGGER.warning("Debug facture ignore: numero invalide fourni (%s).", invoice_number)
        return

    erp = _to_dataframe(erp_records)
    rci = _to_dataframe(rci_records)
    pdf = _to_dataframe(pdf_records)
    reconciliation = _to_dataframe(reconciliation_records)
    consolidated = _safe_consolidated_rci(rci, pdf)

    LOGGER.info("===== DEBUG FACTURE %s =====", invoice)
    _log_records("ERP", invoice, _filter_invoice(erp, invoice))
    _log_records("RCI TXT", invoice, _filter_invoice(rci, invoice))
    _log_records("PDF", invoice, _filter_invoice(pdf, invoice))
    _log_records("RCI consolide", invoice, _filter_invoice(consolidated, invoice))

    final_rows = _filter_invoice(reconciliation, invoice)
    _log_records("Statut final", invoice, final_rows)
    if not final_rows:
        LOGGER.info("Raison statut %s: aucune ligne finale de rapprochement trouvee.", invoice)
    for row in final_rows:
        LOGGER.info("Raison statut %s: %s", invoice, _explain_status(row))
    LOGGER.info("===== FIN DEBUG FACTURE %s =====", invoice)


def _to_dataframe(records: Any) -> "pd.DataFrame":
    pandas = _require_pandas()
    if records is None:
        return pandas.DataFrame()
    if hasattr(records, "copy") and hasattr(records, "columns"):
        return records.copy()
    if isinstance(records, list):
        return pandas.DataFrame(records)
    return pandas.DataFrame(list(records or []))


def _invoice_set(frame: "pd.DataFrame") -> set[str]:
    if frame.empty or "invoice_number" not in frame.columns:
        return set()

    invoices = set()
    for value in frame["invoice_number"]:
        invoice = normalize_invoice_number(value)
        if invoice is not None:
            invoices.add(invoice)
    return invoices


def _sample(values: set[str], limit: int = 20) -> str:
    return ", ".join(sorted(values)[:limit])


def _filter_invoice(frame: "pd.DataFrame", invoice: str) -> list[dict[str, Any]]:
    if frame.empty or "invoice_number" not in frame.columns:
        return []

    mask = frame["invoice_number"].map(normalize_invoice_number) == invoice
    return [_clean_record(record) for record in frame[mask].to_dict("records")]


def _safe_consolidated_rci(rci: "pd.DataFrame", pdf: "pd.DataFrame") -> "pd.DataFrame":
    try:
        return _consolidate_rci_pdf(_aggregate_rci(rci), _aggregate_pdf(pdf))
    except Exception:
        LOGGER.exception("Impossible de construire la vue RCI consolidee pour le debug facture.")
        return _require_pandas().DataFrame()


def _log_records(label: str, invoice: str, records: list[dict[str, Any]]) -> None:
    if not records:
        LOGGER.info("%s %s: aucune ligne trouvee.", label, invoice)
        return

    LOGGER.info("%s %s: %s ligne(s) trouvee(s).", label, invoice, len(records))
    for index, record in enumerate(records, start=1):
        LOGGER.info("%s %s ligne %s: %s", label, invoice, index, json.dumps(record, ensure_ascii=False, default=str))


def _explain_status(row: dict[str, Any]) -> str:
    status = row.get("status")
    amount_erp = row.get("amount_erp")
    amount_rci = row.get("amount_rci")
    amount_pdf = row.get("amount_pdf")
    amount_gap = row.get("amount_gap")
    impacted = row.get("montant_impacte")
    action = row.get("action_recommandee")

    if status == "OK":
        return (
            f"presente ERP et RCI/PDF, montant ERP={amount_erp}, montant RCI={amount_rci}, "
            f"montant PDF={amount_pdf}, ecart={amount_gap}; dans la tolerance."
        )
    if status == "MANQUANTE_RCI":
        return f"presente ERP mais absente RCI/PDF; montant impacte={impacted}. Action: {action}"
    if status == "RCI_SEULEMENT":
        return f"presente RCI/PDF mais absente ERP; montant impacte={impacted}. Action: {action}"
    if status == "ANOMALIE_MONTANT":
        return f"ecart montant hors tolerance; ecart={amount_gap}, montant impacte={impacted}. Action: {action}"
    if status == "ANOMALIE_DATE":
        return f"date incoherente ou manquante; montant impacte={impacted}. Action: {action}"
    if status == "DOUBLON":
        return f"doublon conflictuel detecte; montant impacte={impacted}. Action: {action}"
    if status == "HORS_SCOPE_RCI":
        return f"facture ERP hors perimetre RCI; pas d'ecart impacte. Action: {action}"
    return f"statut {status}; Action: {action}"


def _clean_record(record: dict[str, Any]) -> dict[str, Any]:
    return {key: _clean_value(value) for key, value in record.items()}


def _clean_value(value: Any) -> Any:
    try:
        if pd is not None and pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return value


def _require_pandas() -> Any:
    if pd is None:
        raise RuntimeError(
            "Le module pandas est requis pour generer le debug matching. "
            "Installez les dependances avec: python -m pip install -r requirements.txt"
        )
    return pd
