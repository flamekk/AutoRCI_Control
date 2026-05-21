from __future__ import annotations

import logging

import pytest

pd = pytest.importorskip("pandas")

from src.reference_loader import enrich_erp_with_rci_coverage, load_rci_coverage_reference
from src.reconcile import reconcile_dataframes


def _erp(invoice="VF1000", amount=100.0, is_rci_covered=True):
    return {
        "source_file": "erp.xlsx",
        "source_sheet": "Factures",
        "invoice_number": invoice,
        "document_type": "FACTURE",
        "erp_date": "2026-05-20",
        "customer_code": "C001",
        "customer_name": "Garage Test",
        "amount_erp": amount,
        "is_rci_covered": is_rci_covered,
    }


def test_non_covered_erp_invoice_becomes_hors_scope_rci_not_missing() -> None:
    result = reconcile_dataframes(
        pd.DataFrame([_erp(is_rci_covered=False)]),
        pd.DataFrame(),
        pd.DataFrame(),
    )

    row = result.iloc[0]
    assert row["status"] == "HORS_SCOPE_RCI"
    assert row["status"] != "MANQUANTE_RCI"


def test_covered_erp_invoice_missing_in_rci_stays_missing_rci() -> None:
    result = reconcile_dataframes(
        pd.DataFrame([_erp(is_rci_covered=True)]),
        pd.DataFrame(),
        pd.DataFrame(),
    )

    assert result.iloc[0]["status"] == "MANQUANTE_RCI"


def test_hors_scope_rci_has_zero_montant_impacte() -> None:
    result = reconcile_dataframes(
        pd.DataFrame([_erp(amount=999.0, is_rci_covered=False)]),
        pd.DataFrame(),
        pd.DataFrame(),
    )

    row = result.iloc[0]
    assert row["status"] == "HORS_SCOPE_RCI"
    assert row["montant_impacte"] == pytest.approx(0.0)


def test_missing_reference_continues_with_warning(tmp_path, caplog) -> None:
    erp = pd.DataFrame([_erp(is_rci_covered=False)])

    with caplog.at_level(logging.WARNING):
        coverage = load_rci_coverage_reference(tmp_path / "reference_absente")
        enriched = enrich_erp_with_rci_coverage(erp, coverage, enabled=True)

    assert coverage.loaded is False
    assert "Referentiel RCI" in caplog.text
    assert enriched["is_rci_covered"].tolist() == [True]


def test_reference_file_with_affaires_column(tmp_path) -> None:
    reference_dir = tmp_path / "reference"
    reference_dir.mkdir()
    pd.DataFrame({"Affaires": ["Garage Alpha", "Garage Beta"]}).to_excel(
        reference_dir / "rci.xlsx",
        index=False,
    )

    coverage = load_rci_coverage_reference(reference_dir)

    assert coverage.loaded is True
    assert coverage.rows == 2
    assert "GARAGE ALPHA" in coverage.name_values


def test_reference_file_with_nom_column(tmp_path) -> None:
    reference_dir = tmp_path / "reference"
    reference_dir.mkdir()
    pd.DataFrame({"Nom": ["Société Étoile", "Client Nord"]}).to_excel(
        reference_dir / "rci.xlsx",
        index=False,
    )

    coverage = load_rci_coverage_reference(reference_dir)

    assert coverage.loaded is True
    assert "SOCIETE ETOILE" in coverage.name_values


def test_reference_file_with_single_text_column(tmp_path) -> None:
    reference_dir = tmp_path / "reference"
    reference_dir.mkdir()
    pd.DataFrame([["Garage Alpha"], ["Garage Beta"]]).to_excel(
        reference_dir / "rci.xlsx",
        index=False,
        header=False,
    )

    coverage = load_rci_coverage_reference(reference_dir)

    assert coverage.loaded is True
    assert coverage.rows == 2
    assert {"GARAGE ALPHA", "GARAGE BETA"}.issubset(coverage.name_values)


def test_reference_file_with_shifted_header(tmp_path) -> None:
    reference_dir = tmp_path / "reference"
    reference_dir.mkdir()
    pd.DataFrame(
        [
            ["Liste affaires couvertes RCI", None],
            ["Code client", "Raison sociale"],
            ["000123", "Garage Alpha"],
            ["000456", "Garage Beta"],
        ]
    ).to_excel(reference_dir / "rci.xlsx", index=False, header=False)

    coverage = load_rci_coverage_reference(reference_dir)

    assert coverage.loaded is True
    assert coverage.rows == 2
    assert "123" in coverage.code_values
    assert "GARAGE BETA" in coverage.name_values
