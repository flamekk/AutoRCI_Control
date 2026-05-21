from __future__ import annotations

import pytest

pd = pytest.importorskip("pandas")

from src.extract_erp import STANDARD_COLUMNS, extract_erp_files, extract_erp_folder


def test_extract_erp_files_reads_csv_and_standardizes_columns(tmp_path) -> None:
    csv_path = tmp_path / "export_erp.csv"
    csv_path.write_text(
        "\n".join(
            [
                "No facture;Date facture;Total TTC;N° donneur d'ordre;Client;Code département;Code marque;Type Commande Vente",
                "FVF385380;15/05/26;1.189.358,56;50413091;FAS AUTOMOTIVE;PR;DIVERS;URGENT",
                "ABC123;15/05/26;10,00;999;CLIENT INVALIDE;PR;DIVERS;URGENT",
            ]
        ),
        encoding="utf-8",
    )

    result = extract_erp_files([csv_path])

    assert list(result.columns) == STANDARD_COLUMNS
    assert len(result) == 1
    row = result.iloc[0]
    assert row["source_file"] == "export_erp.csv"
    assert row["source_sheet"] == "CSV"
    assert row["invoice_number"] == "VF385380"
    assert row["document_type"] == "FACTURE"
    assert row["erp_date"] == "2026-05-15"
    assert row["customer_code"] == "50413091"
    assert row["customer_name"] == "FAS AUTOMOTIVE"
    assert row["amount_erp"] == 1189358.56
    assert row["department_code"] == "PR"
    assert row["brand_code"] == "DIVERS"
    assert row["sales_order_type"] == "URGENT"


def test_extract_erp_files_reads_all_excel_sheets(tmp_path) -> None:
    xlsx_path = tmp_path / "navision.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        pd.DataFrame(
            [
                {
                    "N° Facture": "VF385380",
                    "Date facture": "15/05/2026",
                    "Montant": "865,79-",
                    "Concession": "UNIVERS AUTO",
                }
            ]
        ).to_excel(writer, sheet_name="Encours", index=False)
        pd.DataFrame(
            [
                {
                    "N°": "AAF31700",
                    "Date comptabilisation": "15/05/26",
                    "Montant TTC": "-865,79",
                    "Nom du donneur d'ordre": "FAS AUTOMOTIVE",
                }
            ]
        ).to_excel(writer, sheet_name="Avoirs", index=False)

    result = extract_erp_files([xlsx_path])

    assert len(result) == 2
    assert set(result["source_sheet"]) == {"Encours", "Avoirs"}
    assert set(result["invoice_number"]) == {"VF385380", "AAF31700"}
    assert set(result["document_type"]) == {"FACTURE", "AVOIR"}
    assert set(result["amount_erp"]) == {-865.79}


def test_extract_erp_folder_ignores_non_erp_files(tmp_path) -> None:
    (tmp_path / "ignore.txt").write_text("not erp", encoding="utf-8")
    (tmp_path / "erp.csv").write_text(
        "N° facture,Date facture,Montant\nVF385380,15/05/2026,123.45\n",
        encoding="utf-8",
    )

    result = extract_erp_folder(tmp_path)

    assert len(result) == 1
    assert result.iloc[0]["invoice_number"] == "VF385380"


def test_extract_erp_files_handles_unnamed_excel_sheet_layouts(tmp_path) -> None:
    xlsx_path = tmp_path / "headerless.xlsx"
    columns = [f"Unnamed: {index}" for index in range(15)]
    frame = pd.DataFrame(
        [
            [
                "2026-04-29",
                "AX20917",
                "FVF384404",
                "50413061",
                "UNIVERS SYSTEME AUTO KARIMA MA",
                "50413061",
                "UNIVERS SYSTEME AUTO KARIMA MA",
                "3",
                "4 533,35",
                "PR",
                "DIVERS",
                "MCPR",
                "LIVRAISON DIR",
                None,
                None,
            ],
            [
                "AX20917",
                None,
                "AAF31578",
                "2026-04-29",
                None,
                "50413020",
                "N C R A",
                "FG",
                None,
                "305.53",
                "366.64",
                None,
                "MCPR",
                1,
                "FVF381673",
            ],
        ],
        columns=columns,
    )
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Feuil1", index=False)

    result = extract_erp_files([xlsx_path])

    facture = result[result["invoice_number"] == "VF384404"].iloc[0]
    avoir = result[result["invoice_number"] == "AAF31578"].iloc[0]
    assert facture["sales_order_type"] == "LIVRAISON DIR"
    assert facture["amount_erp"] == 4533.35
    assert avoir["erp_date"] == "2026-04-29"
    assert avoir["customer_code"] == "50413020"
    assert avoir["amount_erp"] == 366.64
    assert avoir["sales_order_type"] == "FG"
