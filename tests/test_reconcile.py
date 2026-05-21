from __future__ import annotations

import pytest

pd = pytest.importorskip("pandas")

from src.reconcile import RECONCILIATION_COLUMNS, reconcile, reconcile_dataframes


def test_reconcile_dataframes_produces_ok_with_pdf_priority() -> None:
    erp = pd.DataFrame(
        [
            {
                "source_file": "erp.xlsx",
                "source_sheet": "Factures",
                "invoice_number": "FVF385380",
                "document_type": "FACTURE",
                "customer_code": "50413060",
                "customer_name": "CLIENT A",
                "amount_erp": 100.50,
                "erp_date": "2026-05-15",
            }
        ]
    )
    rci = pd.DataFrame(
        [
            {
                "source_file": "rci.txt",
                "invoice_number": "VF385380",
                "document_type": "FACTURE",
                "rci_date": "2026-05-14",
                "amount_rci": 100.40,
            }
        ]
    )
    pdf = pd.DataFrame(
        [
            {
                "source_file": "daily.pdf",
                "page": 2,
                "invoice_number": "VF385380",
                "document_type": "FACTURE",
                "pdf_invoice_date": "2026-05-15",
                "due_date": "2026-08-01",
                "amount_pdf": 100.00,
                "origin": "ENTREE BATCH",
            }
        ]
    )

    result = reconcile_dataframes(erp, rci, pdf, amount_tolerance=1.0)

    assert list(result.columns) == RECONCILIATION_COLUMNS
    row = result.iloc[0]
    assert row["invoice_number"] == "VF385380"
    assert row["status"] == "OK"
    assert row["amount_gap"] == 0.5
    assert row["pdf_invoice_date"] == "2026-05-15"
    assert row["origin"] == "ENTREE BATCH"


def test_reconcile_dataframes_detects_expected_statuses() -> None:
    erp = pd.DataFrame(
        [
            {"source_file": "erp.xlsx", "invoice_number": "VF1000", "amount_erp": 100, "erp_date": "2026-05-15"},
            {"source_file": "erp.xlsx", "invoice_number": "VF2000", "amount_erp": 100, "erp_date": "2026-05-15"},
            {"source_file": "erp.xlsx", "invoice_number": "VF3000", "amount_erp": 100, "erp_date": "2026-05-15"},
            {"source_file": "erp.xlsx", "invoice_number": "VF4000", "amount_erp": 100, "erp_date": "2026-05-15"},
            {"source_file": "erp.xlsx", "invoice_number": "VF4000", "amount_erp": 110, "erp_date": "2026-05-15"},
        ]
    )
    rci = pd.DataFrame(
        [
            {"source_file": "rci.txt", "invoice_number": "VF2000", "amount_rci": 150, "rci_date": "2026-05-15"},
            {"source_file": "rci.txt", "invoice_number": "VF3000", "amount_rci": 100, "rci_date": "2026-05-14"},
            {"source_file": "rci.txt", "invoice_number": "VF4000", "amount_rci": 100, "rci_date": "2026-05-15"},
            {"source_file": "rci.txt", "invoice_number": "VF5000", "amount_rci": 100, "rci_date": "2026-05-15"},
        ]
    )
    pdf = pd.DataFrame(
        [
            {"source_file": "pdf.pdf", "invoice_number": "VF2000", "amount_pdf": 150, "pdf_invoice_date": "2026-05-15", "due_date": "2026-08-01"},
            {"source_file": "pdf.pdf", "invoice_number": "VF3000", "amount_pdf": 100, "pdf_invoice_date": "2026-05-14", "due_date": "2026-08-01"},
            {"source_file": "pdf.pdf", "invoice_number": "VF4000", "amount_pdf": 100, "pdf_invoice_date": "2026-05-15", "due_date": "2026-08-01"},
            {"source_file": "pdf.pdf", "invoice_number": "VF5000", "amount_pdf": 100, "pdf_invoice_date": "2026-05-15", "due_date": "2026-08-01"},
        ]
    )

    result = reconcile_dataframes(erp, rci, pdf, amount_tolerance=1.0)
    statuses = dict(zip(result["invoice_number"], result["status"]))

    assert statuses["VF1000"] == "MANQUANTE_RCI"
    assert statuses["VF2000"] == "ANOMALIE_MONTANT"
    assert statuses["VF3000"] == "ANOMALIE_DATE"
    assert statuses["VF4000"] == "DOUBLON"
    assert statuses["VF5000"] == "RCI_SEULEMENT"


def test_reconcile_dataframes_normalizes_credit_amount_signs() -> None:
    erp = pd.DataFrame(
        [
            {"source_file": "erp.xlsx", "invoice_number": "AAF31700", "amount_erp": 865.79, "erp_date": "2026-05-15"}
        ]
    )
    rci = pd.DataFrame(
        [
            {"source_file": "rci.txt", "invoice_number": "AAF31700", "amount_rci": 865.79, "rci_date": "2026-05-15"}
        ]
    )
    pdf = pd.DataFrame(
        [
            {
                "source_file": "pdf.pdf",
                "invoice_number": "AAF31700",
                "amount_pdf": "865,79-",
                "pdf_invoice_date": "2026-05-15",
                "due_date": "2026-08-01",
            }
        ]
    )

    result = reconcile_dataframes(erp, rci, pdf)
    row = result.iloc[0]

    assert row["document_type"] == "AVOIR"
    assert row["amount_erp"] == -865.79
    assert row["amount_rci"] == -865.79
    assert row["amount_pdf"] == -865.79
    assert row["status"] == "OK"


def test_reconcile_report_contains_summary_and_reconciliation_rows() -> None:
    erp = pd.DataFrame(
        [{"source_file": "erp.xlsx", "invoice_number": "VF1000", "amount_erp": 100, "erp_date": "2026-05-15"}]
    )
    rci = pd.DataFrame()
    pdf = pd.DataFrame()

    report = reconcile(erp, rci, pdf, amount_tolerance=1.0)

    assert report["summary"]["unmatched_erp"] == 1
    assert report["summary"]["anomalies"] == 1
    assert report["reconciliation"][0]["status"] == "MANQUANTE_RCI"
    assert "transmission" in report["reconciliation"][0]["action_recommandee"]
