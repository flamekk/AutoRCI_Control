from __future__ import annotations

import csv

from src.missing_rci import (
    CATEGORY_AVOIR,
    CATEGORY_FACTURE,
    build_missing_rci_records,
    write_missing_rci_export,
)


def test_missing_rci_records_split_invoices_and_credit_notes() -> None:
    rows = [
        _row("VF1000", "FACTURE", "CRITIQUE", 100000.0),
        _row("AAF31700", "AVOIR", "ELEVEE", 25000.0),
        {"invoice_number": "VFOK", "status": "OK", "montant_impacte": 0.0},
    ]

    missing = build_missing_rci_records(rows)

    assert [row["categorie"] for row in missing] == [CATEGORY_FACTURE, CATEGORY_AVOIR]
    assert [row["invoice_number"] for row in missing] == ["VF1000", "AAF31700"]


def test_missing_rci_export_contains_only_missing_rows(tmp_path) -> None:
    report = {
        "reconciliation": [
            _row("VF1000", "FACTURE", "CRITIQUE", 100000.0),
            {"invoice_number": "VFOK", "status": "OK", "montant_impacte": 0.0},
        ],
        "summary": {},
    }

    path = write_missing_rci_export(report, tmp_path, "20260609_120000")

    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream, delimiter=";"))
    assert path.name == "factures_avoirs_absents_RCI_20260609_120000.csv"
    assert len(rows) == 1
    assert rows[0]["Catégorie"] == CATEGORY_FACTURE
    assert rows[0]["N° facture / avoir"] == "VF1000"


def _row(invoice_number: str, document_type: str, severity: str, amount: float) -> dict:
    return {
        "invoice_number": invoice_number,
        "document_type": document_type,
        "status": "MANQUANTE_RCI",
        "severity": severity,
        "amount_erp": amount,
        "montant_impacte": amount,
        "action_recommandee": "Action",
    }
