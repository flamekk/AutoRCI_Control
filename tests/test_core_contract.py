import pytest

pd = pytest.importorskip("pandas")

from src.normalize import (
    detect_document_type,
    normalize_invoice_number,
    parse_french_amount,
)
from src.reconcile import reconcile_dataframes


def _erp_row(invoice_number="VF1000", amount=100.0, erp_date="2026-05-15", customer_name="Garage Test"):
    return {
        "source_file": "erp.xlsx",
        "source_sheet": "Feuil1",
        "invoice_number": invoice_number,
        "document_type": detect_document_type(invoice_number),
        "erp_date": erp_date,
        "customer_code": "C001",
        "customer_name": customer_name,
        "amount_erp": amount,
        "department_code": "75",
        "brand_code": "REN",
        "sales_order_type": "PR",
    }


def _rci_row(invoice_number="VF1000", amount=100.0, rci_date="2026-05-15"):
    return {
        "source_file": "rci.txt",
        "invoice_number": invoice_number,
        "document_type": detect_document_type(invoice_number),
        "rci_date": rci_date,
        "dealer_code": "C001",
        "amount_rci": amount,
        "raw_line": f"D {invoice_number} {amount}",
    }


def test_normalize_invoice_number():
    assert normalize_invoice_number(" FVF385380 ") == "VF385380"
    assert normalize_invoice_number("vf385380") == "VF385380"
    assert normalize_invoice_number("AAF31700") == "AAF31700"
    assert normalize_invoice_number("") is None
    assert normalize_invoice_number("ABC") is None


def test_parse_french_amount():
    assert parse_french_amount("1.189.358,56") == pytest.approx(1189358.56)
    assert parse_french_amount("865,79-") == pytest.approx(-865.79)
    assert parse_french_amount("-865,79") == pytest.approx(-865.79)
    assert parse_french_amount(" 1 234,50 ") == pytest.approx(1234.50)
    assert parse_french_amount("montant inconnu") is None


def test_detect_document_type():
    assert detect_document_type("VF385380") == "FACTURE"
    assert detect_document_type("FVF385380") == "FACTURE"
    assert detect_document_type("AAF31700") == "AVOIR"
    assert detect_document_type("XYZ123") == "UNKNOWN"


def test_reconciliation_ok():
    erp_df = pd.DataFrame([_erp_row()])
    rci_df = pd.DataFrame([_rci_row()])
    result = reconcile_dataframes(erp_df, rci_df, pd.DataFrame(), amount_tolerance=1.0)

    assert len(result) == 1
    assert result.loc[0, "status"] == "OK"
    assert result.loc[0, "amount_gap"] == pytest.approx(0.0)


def test_reconciliation_missing_rci():
    erp_df = pd.DataFrame([_erp_row()])
    result = reconcile_dataframes(erp_df, pd.DataFrame(), pd.DataFrame(), amount_tolerance=1.0)

    assert len(result) == 1
    assert result.loc[0, "status"] == "MANQUANTE_RCI"


def test_reconciliation_amount_anomaly():
    erp_df = pd.DataFrame([_erp_row(amount=100.0)])
    rci_df = pd.DataFrame([_rci_row(amount=103.0)])
    result = reconcile_dataframes(erp_df, rci_df, pd.DataFrame(), amount_tolerance=1.0)

    assert len(result) == 1
    assert result.loc[0, "status"] == "ANOMALIE_MONTANT"
    assert result.loc[0, "amount_gap"] == pytest.approx(-3.0)


def test_duplicate_detection():
    erp_df = pd.DataFrame([_erp_row(amount=100.0), _erp_row(amount=120.0, customer_name="Garage Test Bis")])
    rci_df = pd.DataFrame([_rci_row()])
    result = reconcile_dataframes(erp_df, rci_df, pd.DataFrame(), amount_tolerance=1.0)

    assert len(result) == 1
    assert result.loc[0, "status"] == "DOUBLON"
    assert "doublons" in result.loc[0, "action_recommandee"].lower()
