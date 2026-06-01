from __future__ import annotations

import pytest

pd = pytest.importorskip("pandas")

from src.date_filter import apply_reconciliation_date_filter
from src.main import parse_args
from src.reconcile import reconcile_dataframes


DATE_FILTER_CONFIG = {
    "enabled": True,
    "mode": "auto",
    "days_before": 1,
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
    }


def _rci(invoice: str, amount: float, rci_date: str) -> dict[str, object]:
    return {
        "source_file": "rci.txt",
        "invoice_number": invoice,
        "document_type": "FACTURE",
        "amount_rci": amount,
        "rci_date": rci_date,
    }


def test_date_filter_reduces_erp_rows() -> None:
    erp = pd.DataFrame(
        [
            _erp("VF1000", 100.0, "2026-05-08"),
            _erp("VF1001", 100.0, "2026-05-09"),
            _erp("VF1002", 100.0, "2026-05-10"),
        ]
    )
    rci = pd.DataFrame([_rci("VF1002", 100.0, "2026-05-10")])

    result = apply_reconciliation_date_filter(erp, rci, pd.DataFrame(), DATE_FILTER_CONFIG)

    assert len(result.erp_records) == 2
    assert result.summary["reconciliation_period"] == "2026-05-09 -> 2026-05-10"
    assert result.summary["erp_rows_excluded_by_date"] == 1


def test_invoice_outside_period_is_not_classified_missing_rci() -> None:
    erp = pd.DataFrame(
        [
            _erp("VF1000", 100.0, "2026-05-01"),
            _erp("VF1001", 200.0, "2026-05-10"),
        ]
    )
    rci = pd.DataFrame([_rci("VF1001", 200.0, "2026-05-10")])
    filtered = apply_reconciliation_date_filter(
        erp,
        rci,
        pd.DataFrame(),
        DATE_FILTER_CONFIG,
        date_from="2026-05-09",
        date_to="2026-05-10",
    )

    reconciliation = reconcile_dataframes(filtered.erp_records, rci, pd.DataFrame())

    assert "VF1000" not in set(reconciliation["invoice_number"])
    assert "MANQUANTE_RCI" not in set(reconciliation["status"])


def test_invoice_inside_period_is_reconciled() -> None:
    erp = pd.DataFrame([_erp("VF1001", 200.0, "2026-05-10")])
    rci = pd.DataFrame([_rci("VF1001", 200.0, "2026-05-10")])
    filtered = apply_reconciliation_date_filter(erp, rci, pd.DataFrame(), DATE_FILTER_CONFIG)

    reconciliation = reconcile_dataframes(filtered.erp_records, rci, pd.DataFrame())

    assert reconciliation.iloc[0]["invoice_number"] == "VF1001"
    assert reconciliation.iloc[0]["status"] == "OK"


def test_no_date_filter_disables_filter() -> None:
    args = parse_args(["--use-samples", "--dry-run", "--ignore-pdf", "--no-date-filter"])
    erp = pd.DataFrame(
        [
            _erp("VF1000", 100.0, "2026-05-01"),
            _erp("VF1001", 200.0, "2026-05-10"),
        ]
    )
    rci = pd.DataFrame([_rci("VF1001", 200.0, "2026-05-10")])

    result = apply_reconciliation_date_filter(
        erp,
        rci,
        pd.DataFrame(),
        DATE_FILTER_CONFIG,
        disable=args.no_date_filter,
    )

    assert args.no_date_filter is True
    assert len(result.erp_records) == 2
    assert len(result.rci_records) == 1
    assert result.summary["date_filter_enabled"] is False


def test_date_filter_applies_to_rci_and_keeps_out_of_period_separately() -> None:
    erp = pd.DataFrame([_erp("VF1001", 200.0, "2026-04-30")])
    rci = pd.DataFrame([_rci("VF1000", 100.0, "2026-04-29")])

    result = apply_reconciliation_date_filter(
        erp,
        rci,
        pd.DataFrame(),
        DATE_FILTER_CONFIG,
        date_from="2026-04-30",
        date_to="2026-04-30",
    )

    assert len(result.rci_records) == 0
    assert len(result.rci_out_of_period_records) == 1
    assert result.summary["rci_rows_before_date_filter"] == 1
    assert result.summary["rci_rows_after_date_filter"] == 0
    assert result.summary["rci_rows_excluded_by_date"] == 1
    assert result.summary["no_rci_flux_in_period_alert"] is True
