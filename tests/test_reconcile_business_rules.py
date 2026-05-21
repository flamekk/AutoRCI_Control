from __future__ import annotations

import pytest

pd = pytest.importorskip("pandas")

from src.reconcile import reconcile, reconcile_dataframes


def _erp(invoice="VF384312", amount=100.0, document_type="FACTURE", date="2026-04-29"):
    return {
        "source_file": "erp.xlsx",
        "source_sheet": "Factures",
        "invoice_number": invoice,
        "document_type": document_type,
        "amount_erp": amount,
        "erp_date": date,
        "customer_code": "50413060",
        "customer_name": "Garage Test",
    }


def _rci(invoice="VF384312", amount=100.0, document_type="FACTURE", date="2026-04-29"):
    return {
        "source_file": "rci.txt",
        "invoice_number": invoice,
        "document_type": document_type,
        "amount_rci": amount,
        "rci_date": date,
    }


def test_ok_when_invoice_exists_in_erp_and_rci_same_amount() -> None:
    result = reconcile_dataframes(
        pd.DataFrame([_erp(amount=100.0)]),
        pd.DataFrame([_rci(amount=100.0)]),
        pd.DataFrame(),
        amount_tolerance=1.0,
    )

    row = result.iloc[0]
    assert row["status"] == "OK"
    assert row["montant_impacte"] == pytest.approx(0.0)


def test_duplicate_same_amount_is_merged_not_doublon() -> None:
    result = reconcile_dataframes(
        pd.DataFrame([_erp(amount=100.0), _erp(amount=100.0)]),
        pd.DataFrame([_rci(amount=100.0)]),
        pd.DataFrame(),
        amount_tolerance=1.0,
    )

    assert len(result) == 1
    assert result.iloc[0]["status"] == "OK"


def test_duplicate_different_amount_is_doublon() -> None:
    result = reconcile_dataframes(
        pd.DataFrame([_erp(amount=100.0), _erp(amount=120.0)]),
        pd.DataFrame([_rci(amount=100.0)]),
        pd.DataFrame(),
        amount_tolerance=1.0,
    )

    row = result.iloc[0]
    assert row["status"] == "DOUBLON"
    assert row["montant_impacte"] == pytest.approx(100.0)


def test_missing_rci_has_montant_impacte() -> None:
    result = reconcile_dataframes(
        pd.DataFrame([_erp(amount=250.0)]),
        pd.DataFrame(),
        pd.DataFrame(),
        amount_tolerance=1.0,
    )

    row = result.iloc[0]
    assert row["status"] == "MANQUANTE_RCI"
    assert row["montant_impacte"] == pytest.approx(250.0)


def test_summary_gap_amount_not_zero_when_missing_exists() -> None:
    report = reconcile(
        pd.DataFrame([_erp(amount=250.0)]),
        pd.DataFrame(),
        pd.DataFrame(),
        amount_tolerance=1.0,
    )

    assert report["summary"]["unmatched_erp"] == 1
    assert report["summary"]["total_impacted_amount"] == pytest.approx(250.0)
    assert report["summary"]["total_amount_gap"] == pytest.approx(250.0)
