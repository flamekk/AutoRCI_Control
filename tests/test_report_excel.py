from __future__ import annotations

import pytest

openpyxl = pytest.importorskip("openpyxl")

from src.report_excel import write_excel_report


def test_write_excel_report_creates_professional_workbook(tmp_path) -> None:
    report = {
        "generated_at": "2026-05-20T14:28:11",
        "status": "completed_with_anomalies",
        "summary": {
            "erp_rows": 2,
            "rci_rows": 1,
            "pdf_rows": 1,
            "matched_invoices": 1,
            "unmatched_erp": 1,
            "reconciled_invoices": 2,
            "erp_matchable_invoices": 2,
            "out_of_scope_rci": 0,
            "gaps_detected": 1,
            "total_impacted_amount": 250.0,
        },
        "reconciliation": [
            {
                "invoice_number": "VF385380",
                "document_type": "FACTURE",
                "customer_code": "000009",
                "customer_name": "RABAT",
                "amount_erp": 100.0,
                "amount_rci": 100.0,
                "amount_pdf": 100.0,
                "amount_gap": 0.0,
                "montant_impacte": 0.0,
                "erp_date": "2026-05-20",
                "rci_date": "2026-05-20",
                "pdf_invoice_date": "2026-05-20",
                "due_date": "2026-08-01",
                "origin": "ENTREE BATCH",
                "status": "OK",
                "priority": "BASSE",
                "action_recommandee": "Aucune action",
                "source_erp": "erp.xlsx:Factures",
                "source_rci": "rci.txt",
                "source_pdf": "daily.pdf:p1",
            },
            {
                "invoice_number": "VF999999",
                "status": "MANQUANTE_RCI",
                "priority": "HAUTE",
                "amount_erp": 250.0,
                "amount_gap": None,
                "montant_impacte": 250.0,
                "customer_code": "000015",
                "customer_name": "TETOUAN",
                "action_recommandee": "Verifier l'integration",
            },
        ],
        "source_files": [],
        "anomalies": [],
        "note": "test",
    }

    path = write_excel_report(report, tmp_path, "20260520_142811")

    assert path.name == "Rapport_Reconciliation_RCI_2026-05-20_1428.xlsx"
    workbook = openpyxl.load_workbook(path)
    assert workbook.sheetnames == [
        "Synthèse",
        "Détail rapprochement",
        "Factures manquantes RCI",
        "Anomalies",
        "Doublons",
        "RCI seulement",
        "Hors scope RCI",
        "Synthèse par concessionnaire",
    ]

    summary = workbook["Synthèse"]
    summary_values = {
        summary.cell(row=row_index, column=1).value: summary.cell(row=row_index, column=2).value
        for row_index in range(1, summary.max_row + 1)
    }
    assert summary_values["Factures analysées"] == 2
    assert summary_values["Factures dans le périmètre RCI"] == 2
    assert summary_values["Factures hors périmètre RCI"] == 0
    assert summary_values["Écarts détectés"] == 1
    assert summary_values["Montant impacté total"] == 250.0

    detail = workbook["Détail rapprochement"]
    assert detail.freeze_panes == "A2"
    assert detail.auto_filter.ref is not None
    status_column = None
    for cell in detail[1]:
        if cell.value == "Statut":
            status_column = cell.column
            break
    assert status_column is not None
    assert detail.cell(row=2, column=status_column).value == "OK"
    assert detail.cell(row=3, column=status_column).value == "MANQUANTE_RCI"
    assert detail.cell(row=2, column=status_column).fill.fgColor.rgb.endswith("C6EFCE")
    assert detail.cell(row=3, column=status_column).fill.fgColor.rgb.endswith("FCE4D6")

    missing = workbook["Factures manquantes RCI"]
    assert missing.max_row == 2
    assert missing["A2"].value == "VF999999"
