from __future__ import annotations

import logging

import pytest

pd = pytest.importorskip("pandas")
pytest.importorskip("pdfplumber")
canvas_module = pytest.importorskip("reportlab.pdfgen.canvas")
pagesizes = pytest.importorskip("reportlab.lib.pagesizes")

from src.extract_pdf import STANDARD_COLUMNS, extract_pdf_files, extract_pdf_folder


def test_extract_pdf_files_reads_invoice_and_credit_lines(tmp_path) -> None:
    pdf_path = tmp_path / "daily.pdf"
    _write_pdf(
        pdf_path,
        [
            "CONCES NO-FAC C.F. DATE FACTURE DATE ECHEANCE MODELE NUMERO DE CHASSIS MONTANT FINANCE CODE TVA ORIGINE",
            "000009 VF385380 PR01 15/05/26 01/08/26 PIE 7.810,93 ENTREE BATCH",
            "000009 AAF31700 PR01 15/05/26 01/08/26 PIE 865,79- ENTREE BATCH",
            "TOTAL CONCESSIONNAIRE : 000009 RABAT FACTURE : 1 AVOIR : 1 6.945,14",
        ],
    )

    result = extract_pdf_files([pdf_path])

    assert list(result.columns) == STANDARD_COLUMNS
    assert len(result) == 2
    facture = result[result["invoice_number"] == "VF385380"].iloc[0]
    avoir = result[result["invoice_number"] == "AAF31700"].iloc[0]
    assert facture["source_file"] == "daily.pdf"
    assert facture["page"] == 1
    assert facture["dealer_code"] == "000009"
    assert facture["document_type"] == "FACTURE"
    assert facture["cf_code"] == "PR01"
    assert facture["pdf_invoice_date"] == "2026-05-15"
    assert facture["due_date"] == "2026-08-01"
    assert facture["model"] == "PIE"
    assert facture["amount_pdf"] == 7810.93
    assert facture["origin"] == "ENTREE BATCH"
    assert avoir["document_type"] == "AVOIR"
    assert avoir["amount_pdf"] == -865.79


def test_extract_pdf_files_reads_model_chassis_and_large_amount(tmp_path) -> None:
    pdf_path = tmp_path / "with_chassis.pdf"
    _write_pdf(
        pdf_path,
        [
            "000024 FVF385399 PR01 15/05/2026 01/08/2026 CLIO VF1ABCDEF12345678 1.189.358,56 ENTREE TP MA",
        ],
    )

    result = extract_pdf_files([pdf_path])

    assert len(result) == 1
    row = result.iloc[0]
    assert row["invoice_number"] == "VF385399"
    assert row["model"] == "CLIO"
    assert row["chassis_number"] == "VF1ABCDEF12345678"
    assert row["amount_pdf"] == 1189358.56
    assert row["origin"] == "ENTREE TP MA"


def test_extract_pdf_files_warns_for_pdf_without_usable_lines(tmp_path, caplog) -> None:
    pdf_path = tmp_path / "empty.pdf"
    _write_pdf(pdf_path, ["Aucun detail exploitable dans ce PDF"])

    with caplog.at_level(logging.WARNING):
        result = extract_pdf_files([pdf_path])

    assert result.empty
    assert "aucune ligne exploitable" in caplog.text.lower()


def test_extract_pdf_folder_ignores_non_pdf_files(tmp_path) -> None:
    (tmp_path / "ignore.txt").write_text("ignore", encoding="utf-8")
    pdf_path = tmp_path / "daily.pdf"
    _write_pdf(pdf_path, ["000009 VF385380 PR01 15/05/26 01/08/26 PIE 7.810,93 ENTREE BATCH"])

    result = extract_pdf_folder(tmp_path)

    assert len(result) == 1
    assert result.iloc[0]["invoice_number"] == "VF385380"


def _write_pdf(path, lines: list[str]) -> None:
    canvas = canvas_module.Canvas(str(path), pagesize=pagesizes.letter)
    canvas.setFont("Courier", 8)
    y = 760
    for line in lines:
        canvas.drawString(30, y, line)
        y -= 14
    canvas.save()
