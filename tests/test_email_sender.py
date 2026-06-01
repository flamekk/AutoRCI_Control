from __future__ import annotations

import logging
from pathlib import Path

from src.email_sender import _send_report_email, send_report


def test_send_report_email_dry_run_logs_planned_content(tmp_path, caplog) -> None:
    report_path = tmp_path / "Rapport_Reconciliation_RCI_2026-05-20_1200.xlsx"
    report_path.write_bytes(b"fake")
    config = {
        "email": {
            "enabled": True,
            "sender": "autorcicontrol@example.com",
            "recipients": ["facturation@example.com"],
            "smtp_host": "smtp.example.com",
            "smtp_port": 587,
        }
    }
    summary = {
        "generated_at": "2026-05-20T12:00:00",
        "reconciled_invoices": 10,
        "erp_matchable_invoices": 8,
        "out_of_scope_rci": 2,
        "matched_invoices": 7,
        "unmatched_erp": 2,
        "rci_out_of_period": 4,
        "gaps_detected": 3,
        "duplicates": 1,
        "total_impacted_amount": 125.5,
        "no_rci_flux_in_period_alert": True,
    }

    with caplog.at_level(logging.INFO):
        status = _send_report_email(summary, report_path, dry_run=True, config=config)

    assert status == "dry_run"
    assert "[AutoRCI] Rapport quotidien de rapprochement ERP/RCI - 2026-05-20" in caplog.text
    assert "Factures analysées : 10" in caplog.text
    assert "Factures dans le périmètre RCI : 8" in caplog.text
    assert "Factures hors périmètre RCI : 2" in caplog.text
    assert "RCI hors période : 4" in caplog.text
    assert "Écarts détectés : 3" in caplog.text
    assert "Montant impacté total : 125.50 MAD" in caplog.text
    assert "aucun flux RCI dans la période de rapprochement" in caplog.text


def test_send_report_returns_disabled_when_email_is_off(tmp_path) -> None:
    report_path = tmp_path / "Rapport_Reconciliation_RCI_2026-05-20_1200.xlsx"
    report_path.write_bytes(b"fake")

    status = send_report(
        {"email": {"enabled": False}},
        [report_path],
        dry_run=False,
        summary={"generated_at": "2026-05-20T12:00:00"},
    )

    assert status == "disabled"


def test_send_report_email_sends_smtp_message(monkeypatch, tmp_path) -> None:
    sent_messages = []

    class FakeSMTP:
        def __init__(self, host, port, timeout):
            self.host = host
            self.port = port
            self.timeout = timeout
            self.started_tls = False
            self.logged_in = None

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        def starttls(self):
            self.started_tls = True

        def login(self, username, password):
            self.logged_in = (username, password)

        def send_message(self, message):
            sent_messages.append(message)

    monkeypatch.setattr("src.email_sender.smtplib.SMTP", FakeSMTP)
    monkeypatch.setenv("AUTORCI_SMTP_PASSWORD", "secret")

    report_path = tmp_path / "Rapport_Reconciliation_RCI_2026-05-20_1200.xlsx"
    report_path.write_bytes(b"fake workbook")
    config = {
        "email": {
            "enabled": True,
            "sender": "autorcicontrol@example.com",
            "recipients": "facturation@example.com;finance@example.com",
            "smtp_host": "smtp.example.com",
            "smtp_port": 587,
            "smtp_username": "autorcicontrol",
            "use_tls": True,
        }
    }

    status = _send_report_email(
        {"generated_at": "2026-05-20T12:00:00"},
        report_path,
        dry_run=False,
        config=config,
    )

    assert status == "sent"
    assert len(sent_messages) == 1
    message = sent_messages[0]
    assert message["Subject"] == "[AutoRCI] Rapport quotidien de rapprochement ERP/RCI - 2026-05-20"
    assert message["To"] == "facturation@example.com, finance@example.com"


def test_send_report_missing_attachment() -> None:
    status = send_report({"email": {"enabled": True}}, [Path("not_a_report.csv")], dry_run=False)

    assert status == "attachment_missing"
