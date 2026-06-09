from __future__ import annotations

import pytest

pd = pytest.importorskip("pandas")

from src.date_filter import apply_reconciliation_date_filter
from src.reconcile import reconcile


DATE_FILTER_CONFIG = {
    "enabled": True,
    "mode": "auto",
    "days_before": 0,
    "days_after": 0,
    "fallback_days": 7,
}


def _erp(invoice: str, amount: float, erp_date: str) -> dict[str, object]:
    return {
        "source_file": "erp.xlsx",
        "source_sheet": "Factures",
        "invoice_number": invoice,
        "document_type": "FACTURE",
        "amount_erp": amount,
        "erp_date": erp_date,
        "customer_code": "50413060",
        "customer_name": "Garage Test",
        "is_rci_covered": True,
    }


def _rci(invoice: str, amount: float, rci_date: str) -> dict[str, object]:
    return {
        "source_file": "rci.txt",
        "invoice_number": invoice,
        "document_type": "FACTURE",
        "amount_rci": amount,
        "rci_date": rci_date,
    }


def _run_period(
    erp_rows: list[dict[str, object]],
    rci_rows: list[dict[str, object]],
    date_from: str,
    date_to: str,
) -> dict[str, object]:
    filtered = apply_reconciliation_date_filter(
        pd.DataFrame(erp_rows),
        pd.DataFrame(rci_rows),
        pd.DataFrame(),
        DATE_FILTER_CONFIG,
        date_from=date_from,
        date_to=date_to,
    )
    report = reconcile(
        filtered.erp_records,
        filtered.rci_records,
        filtered.pdf_records,
        amount_tolerance=1.0,
        rci_out_of_period_records=filtered.rci_out_of_period_records,
        pdf_out_of_period_records=filtered.pdf_out_of_period_records,
    )
    report["summary"].update(filtered.summary)
    return report


def test_rci_out_of_period_becomes_dedicated_status_with_zero_impact() -> None:
    report = _run_period(
        erp_rows=[],
        rci_rows=[_rci("VF384312", 100.0, "2026-04-29")],
        date_from="2026-04-30",
        date_to="2026-04-30",
    )

    row = report["reconciliation"][0]
    assert row["status"] == "RCI_HORS_PERIODE"
    assert row["montant_impacte"] == pytest.approx(0.0)
    assert report["summary"]["rci_out_of_period"] == 1
    assert report["summary"]["total_rci_pdf_out_of_period"] == 1
    assert report["summary"]["erp_analyzed_invoices"] == 0
    assert report["summary"]["unmatched_rci"] == 0
    assert report["summary"]["gaps_detected"] == 0
    assert report["summary"]["no_rci_flux_in_period_alert"] is True


def test_period_30_04_does_not_create_rci_seulement_for_previous_rci_batch() -> None:
    rci_rows = [_rci(f"VF38{i:04d}", 100.0, "2026-04-29") for i in range(97)]
    erp_rows = [_erp(f"VF99{i:04d}", 200.0, "2026-04-30") for i in range(63)]

    report = _run_period(erp_rows, rci_rows, "2026-04-30", "2026-04-30")

    assert report["summary"]["matched_invoices"] == 0
    assert report["summary"]["unmatched_erp"] == 63
    assert report["summary"]["unmatched_rci"] == 0
    assert report["summary"]["rci_out_of_period"] == 97
    assert report["summary"]["total_rci_pdf_out_of_period"] == 97
    assert report["summary"]["erp_analyzed_invoices"] == 63
    assert report["summary"]["gaps_detected"] == 63


def test_period_29_04_keeps_matching_rci_batch_ok() -> None:
    erp_rows = [_erp(f"VF38{i:04d}", 100.0, "2026-04-29") for i in range(97)]
    rci_rows = [_rci(f"VF38{i:04d}", 100.0, "2026-04-29") for i in range(97)]

    report = _run_period(erp_rows, rci_rows, "2026-04-29", "2026-04-29")

    assert report["summary"]["matched_invoices"] == 97
    assert report["summary"]["erp_analyzed_invoices"] == 97
    assert report["summary"]["unmatched_erp"] == 0
    assert report["summary"]["unmatched_rci"] == 0
    assert report["summary"]["rci_out_of_period"] == 0
