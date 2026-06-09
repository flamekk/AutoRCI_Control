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
            "erp_analyzed_invoices": 2,
            "reconciled_invoices": 2,
            "erp_matchable_invoices": 2,
            "out_of_scope_rci": 0,
            "rci_out_of_period": 1,
            "total_rci_pdf_out_of_period": 1,
            "gaps_detected": 1,
            "total_impacted_amount": 250.0,
            "matching_rate": 0.5,
            "missing_rci_out_of_period": 0,
            "out_of_scope_rci_percent": 0.0,
            "low_matching_rate_alert": True,
            "information_lines": 1,
            "corrective_batch_generated": True,
            "corrective_batch_invoice_count": 1,
            "corrective_batch_total_amount": 25000.0,
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
                "severity": "OK",
                "action_recommandee": "Aucune action",
                "source_erp": "erp.xlsx:Factures",
                "source_rci": "rci.txt",
                "source_pdf": "daily.pdf:p1",
            },
            {
                "invoice_number": "VF999999",
                "status": "MANQUANTE_RCI",
                "priority": "HAUTE",
                "severity": "MOYENNE",
                "amount_erp": 250.0,
                "amount_gap": None,
                "montant_impacte": 250.0,
                "customer_code": "000015",
                "customer_name": "TETOUAN",
                "action_recommandee": "Verifier l'integration",
            },
            {
                "invoice_number": "VF888888",
                "document_type": "FACTURE",
                "amount_rci": 100.0,
                "rci_date": "2026-04-29",
                "status": "RCI_HORS_PERIODE",
                "priority": "BASSE",
                "severity": "INFORMATION",
                "montant_impacte": 0.0,
                "source_rci": "rci.txt",
                "action_recommandee": "Aucune action",
                "commentaire_audit": "Flux RCI/PDF hors période de rapprochement",
            },
        ],
        "audits": {
            "dates": [
                {
                    "erp_date": "2026-05-20",
                    "nombre_total_factures_erp": 2,
                    "nombre_ok": 1,
                    "nombre_manquante_rci": 1,
                    "nombre_hors_scope_rci": 0,
                    "montant_manquant_rci": 250.0,
                    "taux_rapprochement_date": 0.5,
                    "rci_count": 1,
                    "rci_hors_periode_count": 0,
                },
                {
                    "erp_date": "2026-04-29",
                    "nombre_total_factures_erp": 0,
                    "nombre_ok": 0,
                    "nombre_manquante_rci": 0,
                    "nombre_hors_scope_rci": 0,
                    "montant_manquant_rci": 0.0,
                    "taux_rapprochement_date": 0.0,
                    "rci_count": 0,
                    "rci_hors_periode_count": 1,
                }
            ],
            "missing_rci": [
                {
                    "invoice_number": "VF999999",
                    "erp_date": "2026-05-20",
                    "customer_name": "TETOUAN",
                    "customer_code": "000015",
                    "amount_erp": 250.0,
                    "is_rci_covered": True,
                    "status": "MANQUANTE_RCI",
                    "montant_impacte": 250.0,
                    "source_erp": "erp.xlsx:Factures",
                    "date_in_filter": True,
                    "commentaire_audit": "Dans la période filtrée",
                }
            ],
            "out_of_scope_rci": [],
        },
        "corrective_batch": {
            "generated": True,
            "batch_path": str(tmp_path / "batch_correctif_candidat_20260520_142811.txt"),
            "control_path": str(tmp_path / "batch_correctif_candidat_20260520_142811_control.csv"),
            "invoice_count": 1,
            "total_amount": 25000.0,
            "warning": "Fichier candidat à valider par l’équipe facturation avant transmission à RCI.",
            "records": [
                {
                    "invoice_number": "VF777777",
                    "erp_date": "2026-05-20",
                    "customer_name": "CASABLANCA",
                    "amount_erp": 25000.0,
                    "montant_impacte": 25000.0,
                    "severity": "ELEVEE",
                    "status": "MANQUANTE_RCI",
                    "action_recommandee": "Vérifier la transmission vers RCI et préparer un renvoi si nécessaire.",
                }
            ],
        },
        "source_files": [],
        "anomalies": [],
        "note": "test",
    }

    path = write_excel_report(report, tmp_path, "20260520_142811")

    assert path.name == "Rapport_Reconciliation_RCI_2026-05-20_1428.xlsx"
    workbook = openpyxl.load_workbook(path)
    assert workbook.sheetnames == [
        "Synthèse",
        "Factures et avoirs absents RCI",
        "Détail rapprochement",
        "Plan action",
        "Batch correctif",
        "Hors périmètre RCI",
        "RCI PDF hors période",
        "Qualité référentiel RCI",
        "Synthèse par concessionnaire",
        "Audit",
    ]

    summary = workbook["Synthèse"]
    summary_values = {
        summary.cell(row=row_index, column=1).value: summary.cell(row=row_index, column=2).value
        for row_index in range(1, summary.max_row + 1)
    }
    assert summary_values["Nombre total de factures/avoirs ERP analysés"] == 2
    assert summary_values["Nombre de factures présentes dans RCI"] == 1
    assert summary_values["Nombre de factures absentes RCI"] == 1
    assert summary_values["Nombre d'avoirs absents RCI"] == 0
    assert summary_values["Montant total absent RCI"] == 250.0
    assert summary_values["Nombre d'écarts critiques"] == 0
    assert summary_values["Nombre d'écarts élevés"] == 0
    assert summary_values["Nombre d'écarts moyens"] == 1
    assert summary_values["Factures dans le périmètre RCI"] == 2
    assert summary_values["Factures hors périmètre RCI"] == 0
    assert summary_values["Écarts détectés"] == 1
    assert summary_values["Batch correctif candidat généré"] == "Oui"
    assert summary_values["Nombre factures incluses batch correctif"] == 1
    assert summary_values["Montant total inclus batch correctif"] == 25000.0
    assert summary_values["RCI hors période"] == 1
    assert summary_values["Total lignes RCI/PDF hors période"] == 1

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

    missing = workbook["Factures et avoirs absents RCI"]
    assert missing.max_row == 2
    assert missing["A2"].value == "Facture absente"
    assert missing["C2"].value == "VF999999"

    assert workbook["Audit"].max_row >= 6
    assert workbook["RCI PDF hors période"].max_row == 2
    assert workbook["Plan action"].max_row == 2
    assert workbook["Batch correctif"].max_row >= 9


def test_missing_rci_sheet_displays_empty_message(tmp_path) -> None:
    report = {
        "generated_at": "2026-05-20T14:28:11",
        "status": "completed",
        "summary": {
            "erp_analyzed_invoices": 1,
            "erp_matchable_invoices": 1,
            "matched_invoices": 1,
            "matching_rate": 1.0,
        },
        "reconciliation": [
            {
                "invoice_number": "VF385380",
                "document_type": "FACTURE",
                "amount_erp": 100.0,
                "montant_impacte": 0.0,
                "status": "OK",
                "severity": "OK",
            }
        ],
        "source_files": [],
        "anomalies": [],
    }

    path = write_excel_report(report, tmp_path, "20260520_142811")

    workbook = openpyxl.load_workbook(path)
    sheet = workbook["Factures et avoirs absents RCI"]
    assert sheet["A1"].value == "Message"
    assert sheet["A2"].value == "Aucune facture ou avoir absent côté RCI pour cette période."
