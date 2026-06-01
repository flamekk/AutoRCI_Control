from __future__ import annotations

import pytest

from src.audit import enrich_report_with_audits


def _record(
    invoice: str,
    status: str,
    erp_date: str,
    amount: float = 100.0,
    customer_name: str = "Garage Test",
) -> dict[str, object]:
    return {
        "invoice_number": invoice,
        "erp_date": erp_date,
        "customer_name": customer_name,
        "customer_code": "50413060",
        "amount_erp": amount,
        "montant_impacte": amount if status == "MANQUANTE_RCI" else 0.0,
        "is_rci_covered": status != "HORS_SCOPE_RCI",
        "status": status,
        "source_erp": "erp.xlsx:Factures",
    }


def _report(records: list[dict[str, object]]) -> dict[str, object]:
    return {
        "summary": {
            "reconciled_invoices": len(records),
            "matched_invoices": sum(1 for row in records if row["status"] == "OK"),
            "unmatched_erp": sum(1 for row in records if row["status"] == "MANQUANTE_RCI"),
            "out_of_scope_rci": sum(1 for row in records if row["status"] == "HORS_SCOPE_RCI"),
            "erp_matchable_invoices": sum(
                1
                for row in records
                if row["status"] in {"OK", "MANQUANTE_RCI", "ANOMALIE_MONTANT", "ANOMALIE_DATE", "DOUBLON"}
            ),
            "matching_rate": 0.5,
            "reconciliation_start_date": "2026-05-01",
            "reconciliation_end_date": "2026-05-05",
        },
        "reconciliation": records,
    }


def test_audit_dates_calcule_ok_manquante_hors_scope_par_date() -> None:
    report = _report(
        [
            {**_record("VF1", "OK", "2026-05-02", amount=100.0), "rci_date": "2026-05-02", "source_rci": "rci.txt"},
            _record("VF2", "MANQUANTE_RCI", "2026-05-02", amount=250.0),
            _record("VF3", "HORS_SCOPE_RCI", "2026-05-03", amount=80.0),
            {
                "invoice_number": "VF4",
                "status": "RCI_HORS_PERIODE",
                "rci_date": "2026-04-29",
                "amount_rci": 90.0,
                "source_rci": "rci.txt",
                "montant_impacte": 0.0,
            },
        ]
    )

    enrich_report_with_audits(report, reference_names=["GARAGE TEST"], log=False)

    rows = {row["erp_date"]: row for row in report["audits"]["dates"]}
    assert rows["2026-05-02"]["nombre_total_factures_erp"] == 2
    assert rows["2026-05-02"]["nombre_ok"] == 1
    assert rows["2026-05-02"]["nombre_manquante_rci"] == 1
    assert rows["2026-05-02"]["montant_manquant_rci"] == pytest.approx(250.0)
    assert rows["2026-05-02"]["taux_rapprochement_date"] == pytest.approx(0.5)
    assert rows["2026-05-02"]["rci_count"] == 1
    assert rows["2026-05-03"]["nombre_hors_scope_rci"] == 1
    assert rows["2026-04-29"]["rci_hors_periode_count"] == 1


def test_date_in_filter_false_si_facture_hors_periode() -> None:
    report = _report([_record("VF4", "MANQUANTE_RCI", "2026-04-30", amount=300.0)])

    enrich_report_with_audits(report, log=False)

    row = report["audits"]["missing_rci"][0]
    assert row["date_in_filter"] is False
    assert row["commentaire_audit"] == "Hors période filtrée - vérifier filtre date"
    assert report["summary"]["missing_rci_out_of_period"] == 1
    assert report["summary"]["missing_rci_out_of_period_alert"] is True


def test_closest_reference_name_est_renseigne_pour_hors_scope_rci() -> None:
    report = _report(
        [
            _record(
                "VF5",
                "HORS_SCOPE_RCI",
                "2026-05-03",
                amount=80.0,
                customer_name="Garage Centrale",
            )
        ]
    )

    enrich_report_with_audits(report, reference_names=["Garage Central RCI", "Autre Client"], log=False)

    row = report["audits"]["out_of_scope_rci"][0]
    assert row["normalized_customer_name"] == "GARAGE CENTRALE"
    assert row["closest_reference_name"] == "GARAGE CENTRAL RCI"
    assert row["closest_reference_similarity"] > 0.7
