from __future__ import annotations

import csv

from src.action_plan import build_reference_suggestions, severity_for_status
from src.corrective_batch import select_corrective_batch_records, write_corrective_batch_outputs


def test_severity_missing_rci_critical_threshold() -> None:
    assert severity_for_status("MANQUANTE_RCI", 100000) == "CRITIQUE"
    assert severity_for_status("MANQUANTE_RCI", 250000) == "CRITIQUE"


def test_severity_missing_rci_high_threshold() -> None:
    assert severity_for_status("MANQUANTE_RCI", 20000) == "ELEVEE"
    assert severity_for_status("MANQUANTE_RCI", 99999.99) == "ELEVEE"


def test_severity_missing_rci_medium_threshold() -> None:
    assert severity_for_status("MANQUANTE_RCI", 19999.99) == "MOYENNE"


def test_severity_special_statuses() -> None:
    assert severity_for_status("HORS_SCOPE_RCI", 50000) == "A_VERIFIER"
    assert severity_for_status("RCI_HORS_PERIODE", 50000) == "INFORMATION"
    assert severity_for_status("OK", 0) == "OK"


def test_reference_suggestions_group_hors_scope_customers() -> None:
    rows = [
        {
            "invoice_number": "VF1",
            "status": "HORS_SCOPE_RCI",
            "customer_name": "Client Test",
            "amount_erp": 100.0,
        },
        {
            "invoice_number": "VF2",
            "status": "HORS_SCOPE_RCI",
            "customer_name": "CLIENT   TEST",
            "amount_erp": 150.0,
        },
        {
            "invoice_number": "VF3",
            "status": "OK",
            "customer_name": "Client Test",
            "amount_erp": 200.0,
        },
    ]

    suggestions = build_reference_suggestions(rows)

    assert len(suggestions) == 1
    assert suggestions[0]["nombre_factures"] == 2
    assert suggestions[0]["montant_total_erp"] == 250.0


def test_corrective_batch_contains_only_critical_and_high_missing_rci(tmp_path) -> None:
    rows = [
        _row("VF1", "MANQUANTE_RCI", "CRITIQUE", 150000.0),
        _row("VF2", "MANQUANTE_RCI", "ELEVEE", 25000.0),
        _row("VF3", "MANQUANTE_RCI", "MOYENNE", 1000.0),
        _row("VF4", "HORS_SCOPE_RCI", "A_VERIFIER", 500000.0),
        _row("VF5", "RCI_HORS_PERIODE", "INFORMATION", 500000.0),
    ]
    report = {"reconciliation": rows, "summary": {}}

    artifacts = write_corrective_batch_outputs(report, tmp_path, "20260608_120000")
    candidates = select_corrective_batch_records(report["reconciliation"])

    assert [row["invoice_number"] for row in candidates] == ["VF1", "VF2"]
    assert all(row["status"] == "MANQUANTE_RCI" for row in candidates)
    assert all(row["severity"] in {"CRITIQUE", "ELEVEE"} for row in candidates)
    assert report["summary"]["corrective_batch_invoice_count"] == 2
    assert report["summary"]["corrective_batch_total_amount"] == 175000.0

    batch_text = (tmp_path / "batch_correctif_candidat_20260608_120000.txt").read_text(encoding="utf-8")
    assert "VF1" in batch_text
    assert "VF2" in batch_text
    assert "VF3" not in batch_text
    assert "VF4" not in batch_text
    assert "VF5" not in batch_text

    control_path = tmp_path / "batch_correctif_candidat_20260608_120000_control.csv"
    with control_path.open("r", encoding="utf-8-sig", newline="") as stream:
        control_rows = list(csv.DictReader(stream, delimiter=";"))
    assert len(control_rows) == 2
    assert {row["included_in_corrective_batch"] for row in control_rows} == {"true"}
    assert artifacts == [tmp_path / "batch_correctif_candidat_20260608_120000.txt", control_path]


def test_corrective_batch_not_generated_when_no_priority_missing_rci(tmp_path) -> None:
    rows = [
        _row("VF3", "MANQUANTE_RCI", "MOYENNE", 1000.0),
        _row("VF4", "HORS_SCOPE_RCI", "A_VERIFIER", 500000.0),
        _row("VF5", "RCI_HORS_PERIODE", "INFORMATION", 500000.0),
    ]
    report = {"reconciliation": rows, "summary": {}}

    artifacts = write_corrective_batch_outputs(report, tmp_path, "20260608_120000")

    assert artifacts == []
    assert report["summary"]["corrective_batch_generated"] is False
    assert not (tmp_path / "batch_correctif_candidat_20260608_120000.txt").exists()
    assert not (tmp_path / "batch_correctif_candidat_20260608_120000_control.csv").exists()


def _row(invoice_number: str, status: str, severity: str, amount: float) -> dict:
    return {
        "invoice_number": invoice_number,
        "erp_date": "2026-05-01",
        "customer_name": "Client",
        "amount_erp": amount,
        "montant_impacte": amount,
        "severity": severity,
        "status": status,
        "action_recommandee": "Action",
    }
