from __future__ import annotations

import csv

from src.powerbi_export import HISTORY_COLUMNS, update_reconciliation_history, write_powerbi_exports


def test_update_reconciliation_history_appends_and_deduplicates_same_run(tmp_path) -> None:
    history_path = tmp_path / "reconciliation_history.csv"
    records = [
        {
            "invoice_number": "VF385380",
            "document_type": "FACTURE",
            "customer_code": "000009",
            "customer_name": "RABAT",
            "amount_erp": 100.0,
            "amount_rci": 100.0,
            "amount_pdf": 100.0,
            "amount_gap": 0.0,
            "erp_date": "2026-05-20",
            "pdf_invoice_date": "2026-05-20",
            "due_date": "2026-08-01",
            "origin": "ENTREE BATCH",
            "status": "OK",
            "priority": "BASSE",
            "severity": "OK",
            "included_in_corrective_batch": False,
            "action_recommandee": "Aucune action",
        }
    ]

    update_reconciliation_history(records, history_path, "20260520_100000", "2026-05-20")
    update_reconciliation_history(records, history_path, "20260520_100000", "2026-05-20")
    update_reconciliation_history(records, history_path, "20260521_100000", "2026-05-21")

    rows = _read_csv(history_path)

    assert list(rows[0].keys()) == HISTORY_COLUMNS
    assert len(rows) == 2
    assert rows[0]["processing_run_id"] == "20260520_100000"
    assert rows[1]["processing_run_id"] == "20260521_100000"


def test_write_powerbi_exports_creates_history_file(tmp_path) -> None:
    report = {
        "generated_at": "2026-05-20T14:28:11",
        "status": "ok",
        "summary": {"matched_invoices": 1},
        "source_files": [],
        "reconciliation": [
            {
                "invoice_number": "VF385380",
                "document_type": "FACTURE",
                "status": "OK",
                "priority": "BASSE",
                "severity": "OK",
                "included_in_corrective_batch": False,
                "amount_erp": 100,
                "amount_gap": 0,
                "action_recommandee": "Aucune action",
            }
        ],
    }

    artifacts = write_powerbi_exports(report, tmp_path, "20260520_142811")

    history_path = tmp_path / "reconciliation_history.csv"
    assert history_path in artifacts
    rows = _read_csv(history_path)
    assert len(rows) == 1
    assert rows[0]["processing_date"] == "2026-05-20"
    assert rows[0]["processing_run_id"] == "20260520_142811"
    assert rows[0]["invoice_number"] == "VF385380"
    assert "severity" in rows[0]
    assert "included_in_corrective_batch" in rows[0]
    assert rows[0]["severity"] == "OK"


def _read_csv(path):
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        return list(csv.DictReader(stream, delimiter=";"))
