from __future__ import annotations

import logging

import pytest

pd = pytest.importorskip("pandas")

from src.extract_rci import STANDARD_COLUMNS, extract_rci_files, extract_rci_folder


def test_extract_rci_files_reads_txt_detail_lines(tmp_path, caplog) -> None:
    rci_path = tmp_path / "rci_export.txt"
    rci_path.write_text(
        "\n".join(
            [
                "HBMAPR-MA0343820260505",
                "D     0 RENAULT PIE     REN            0P1            SAP050413060      2026042920260429     00000010307380000001030738                00000008589480000000171790                                       VF384312",
                "D     0 RENAULT PIE     REN            1P1            SAP050413023      2026042920260429     00000001346450000000134645                00000001122040000000022441                                       AAF31571",
                "D ligne detail sans facture exploitable",
            ]
        ),
        encoding="utf-8",
    )

    with caplog.at_level(logging.WARNING):
        result = extract_rci_files([rci_path])

    assert list(result.columns) == STANDARD_COLUMNS
    assert len(result) == 2
    first = result[result["invoice_number"] == "VF384312"].iloc[0]
    second = result[result["invoice_number"] == "AAF31571"].iloc[0]
    assert first["document_type"] == "FACTURE"
    assert first["rci_date"] == "2026-04-29"
    assert first["dealer_code"] == "50413060"
    assert first["amount_rci"] == 10307.38
    assert second["document_type"] == "AVOIR"
    assert second["dealer_code"] == "50413023"
    assert second["amount_rci"] == 1346.45
    assert "sans facture valide" in caplog.text


def test_extract_rci_files_reads_csv_and_detects_columns(tmp_path) -> None:
    csv_path = tmp_path / "rci.csv"
    csv_path.write_text(
        "\n".join(
            [
                "Invoice number;Date RCI;Dealer code;Amount",
                "FVF385380;20260515;SAP050413060;1.189.358,56",
                "ABC123;20260515;SAP050413999;10,00",
            ]
        ),
        encoding="utf-8",
    )

    result = extract_rci_files([csv_path])

    assert len(result) == 1
    row = result.iloc[0]
    assert row["source_file"] == "rci.csv"
    assert row["invoice_number"] == "VF385380"
    assert row["document_type"] == "FACTURE"
    assert row["rci_date"] == "2026-05-15"
    assert row["dealer_code"] == "50413060"
    assert row["amount_rci"] == 1189358.56
    assert "FVF385380" in row["raw_line"]


def test_extract_rci_files_reads_all_excel_sheets(tmp_path) -> None:
    xlsx_path = tmp_path / "rci.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        pd.DataFrame(
            [
                {
                    "Facture": "VF385380",
                    "Date": "15/05/2026",
                    "Dealer": "50413060",
                    "Montant RCI": "865,79-",
                }
            ]
        ).to_excel(writer, sheet_name="Factures", index=False)
        pd.DataFrame(
            [
                {
                    "Document": "AAF31700",
                    "Accounting date": "20260515",
                    "Code concession": "SAP050413023",
                    "Total": "-865,79",
                }
            ]
        ).to_excel(writer, sheet_name="Avoirs", index=False)

    result = extract_rci_files([xlsx_path])

    assert len(result) == 2
    assert set(result["invoice_number"]) == {"VF385380", "AAF31700"}
    assert set(result["document_type"]) == {"FACTURE", "AVOIR"}
    assert set(result["rci_date"]) == {"2026-05-15"}
    assert set(result["amount_rci"]) == {-865.79}


def test_extract_rci_folder_ignores_non_rci_files(tmp_path) -> None:
    (tmp_path / "ignore.pdf").write_text("ignore", encoding="utf-8")
    (tmp_path / "rci.csv").write_text(
        "Facture;Date;Montant\nVF385380;15/05/2026;123.45\n",
        encoding="utf-8",
    )

    result = extract_rci_folder(tmp_path)

    assert len(result) == 1
    assert result.iloc[0]["invoice_number"] == "VF385380"
