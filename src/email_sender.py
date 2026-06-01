from __future__ import annotations

import logging
import os
import smtplib
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from mimetypes import guess_type
from pathlib import Path
from typing import Any

try:
    from src.config import load_config
except (ModuleNotFoundError, ImportError):  # pragma: no cover - used when running python src/main.py.
    from config import load_config


LOGGER = logging.getLogger(__name__)

DEFAULT_PASSWORD_ENV_VAR = "AUTORCI_SMTP_PASSWORD"


@dataclass(frozen=True)
class EmailSettings:
    enabled: bool
    sender: str
    recipients: list[str]
    smtp_host: str
    smtp_port: int
    smtp_username: str
    smtp_password: str | None
    use_tls: bool
    use_ssl: bool


def send_report_email(summary: dict[str, Any], report_path: str | Path, dry_run: bool = False) -> str:
    config = load_config()
    return _send_report_email(summary, Path(report_path), dry_run=dry_run, config=config)


def send_report(
    config: dict[str, Any],
    artifacts: list[Path],
    dry_run: bool,
    summary: dict[str, Any] | None = None,
) -> str:
    report_path = _find_report_attachment(artifacts)
    if report_path is None:
        LOGGER.warning("Aucun rapport Excel trouve parmi les artefacts: %s", artifacts)
        return "attachment_missing"

    return _send_report_email(summary or {}, report_path, dry_run=dry_run, config=config)


def _send_report_email(
    summary: dict[str, Any],
    report_path: Path,
    dry_run: bool,
    config: dict[str, Any],
) -> str:
    settings = _load_email_settings(config)
    processing_date = _summary_date(summary)
    subject = f"[AutoRCI] Rapport quotidien de rapprochement ERP/RCI - {processing_date}"
    body = _build_email_body(summary)

    if dry_run:
        LOGGER.info("Dry-run actif: aucun email ne sera envoye.")
        LOGGER.info("Email prevu - destinataires: %s", ", ".join(settings.recipients) or "(non configures)")
        LOGGER.info("Email prevu - objet: %s", subject)
        LOGGER.info("Email prevu - piece jointe: %s", report_path)
        LOGGER.info("Email prevu - corps:\n%s", body)
        return "dry_run"

    if not settings.enabled:
        LOGGER.info("Email desactive dans la configuration.")
        return "disabled"

    validation_error = _validate_settings(settings, report_path)
    if validation_error is not None:
        LOGGER.error(validation_error)
        return "configuration_error"

    message = _build_message(settings, subject, body, report_path)
    try:
        _send_smtp_message(settings, message)
    except Exception:
        LOGGER.exception("Echec de l'envoi email SMTP.")
        return "failed"

    LOGGER.info("Email envoye a %s avec le rapport %s.", ", ".join(settings.recipients), report_path)
    return "sent"


def _load_email_settings(config: dict[str, Any]) -> EmailSettings:
    email_config = config.get("email", {})
    password_env_var = _env("AUTORCI_SMTP_PASSWORD_ENV_VAR") or str(
        email_config.get("smtp_password_env_var") or DEFAULT_PASSWORD_ENV_VAR
    )
    password = _env("AUTORCI_SMTP_PASSWORD") or os.getenv(password_env_var)

    return EmailSettings(
        enabled=_env_bool("AUTORCI_EMAIL_ENABLED", email_config.get("enabled", False)),
        sender=_env("AUTORCI_EMAIL_SENDER") or str(email_config.get("sender") or ""),
        recipients=_recipients(
            _env("AUTORCI_EMAIL_RECIPIENTS")
            or _env("AUTORCI_EMAIL_TO")
            or email_config.get("recipients", [])
        ),
        smtp_host=_env("AUTORCI_SMTP_HOST") or str(email_config.get("smtp_host") or ""),
        smtp_port=int(_env("AUTORCI_SMTP_PORT") or email_config.get("smtp_port") or 587),
        smtp_username=_env("AUTORCI_SMTP_USERNAME")
        or _env("AUTORCI_SMTP_USER")
        or str(email_config.get("smtp_username") or ""),
        smtp_password=password,
        use_tls=_env_bool("AUTORCI_SMTP_USE_TLS", email_config.get("use_tls", True)),
        use_ssl=_env_bool("AUTORCI_SMTP_USE_SSL", email_config.get("use_ssl", False)),
    )


def _validate_settings(settings: EmailSettings, report_path: Path) -> str | None:
    if not report_path.exists():
        return f"Rapport Excel introuvable pour email: {report_path}"
    if not settings.sender:
        return "Email non envoye: sender manquant."
    if not settings.recipients:
        return "Email non envoye: aucun destinataire configure."
    if not settings.smtp_host:
        return "Email non envoye: smtp_host manquant."
    if settings.smtp_username and not settings.smtp_password:
        return "Email non envoye: mot de passe SMTP manquant dans la variable d'environnement."
    return None


def _build_message(
    settings: EmailSettings,
    subject: str,
    body: str,
    report_path: Path,
) -> EmailMessage:
    message = EmailMessage()
    message["From"] = settings.sender
    message["To"] = ", ".join(settings.recipients)
    message["Subject"] = subject
    message.set_content(body)

    content_type, _ = guess_type(report_path.name)
    maintype, subtype = (content_type or "application/octet-stream").split("/", 1)
    message.add_attachment(
        report_path.read_bytes(),
        maintype=maintype,
        subtype=subtype,
        filename=report_path.name,
    )
    return message


def _send_smtp_message(settings: EmailSettings, message: EmailMessage) -> None:
    smtp_class = smtplib.SMTP_SSL if settings.use_ssl else smtplib.SMTP
    with smtp_class(settings.smtp_host, settings.smtp_port, timeout=30) as smtp:
        if settings.use_tls and not settings.use_ssl:
            smtp.starttls()
        if settings.smtp_username:
            smtp.login(settings.smtp_username, settings.smtp_password or "")
        smtp.send_message(message)


def _build_email_body(summary: dict[str, Any]) -> str:
    analyzed = _summary_int(summary, "reconciled_invoices", "erp_rows")
    in_scope = _summary_int(summary, "erp_matchable_invoices")
    out_of_scope = _summary_int(summary, "out_of_scope_rci", "status_hors_scope_rci")
    ok = _summary_int(summary, "matched_invoices", "status_ok")
    missing = _summary_int(summary, "unmatched_erp", "status_manquante_rci")
    rci_out_of_period = _summary_int(summary, "rci_out_of_period", "status_rci_hors_periode")
    gaps_detected = _summary_int(summary, "gaps_detected", "anomalies")
    total_impacted = _summary_float(summary, "total_impacted_amount", "total_amount_gap")
    no_rci_alert = _summary_bool(summary, "no_rci_flux_in_period_alert")
    no_rci_alert_text = (
        "Alerte : aucun flux RCI dans la période de rapprochement.\n\n"
        if no_rci_alert
        else ""
    )

    if in_scope == 0:
        in_scope = (
            ok
            + missing
            + _summary_int(summary, "amount_anomalies", "status_anomalie_montant")
            + _summary_int(summary, "date_anomalies", "status_anomalie_date")
            + _summary_int(summary, "duplicates", "status_doublon")
        )

    return (
        "Bonjour,\n\n"
        "Le contrôle automatique ERP Navision / RCI Banque a été exécuté.\n\n"
        "Synthèse :\n"
        f"- Factures analysées : {analyzed}\n"
        f"- Factures dans le périmètre RCI : {in_scope}\n"
        f"- Factures hors périmètre RCI : {out_of_scope}\n"
        f"- Factures OK : {ok}\n"
        f"- Factures manquantes RCI : {missing}\n"
        f"- RCI hors période : {rci_out_of_period}\n"
        f"- Écarts détectés : {gaps_detected}\n"
        f"- Montant impacté total : {total_impacted:,.2f} MAD\n\n"
        f"{no_rci_alert_text}"
        "Le rapport détaillé est disponible en pièce jointe.\n\n"
        "Cordialement,\n"
        "AutoRCI Control"
    )


def _summary_date(summary: dict[str, Any]) -> str:
    for key in ("processing_date", "date", "run_date"):
        value = summary.get(key)
        if value:
            return str(value)[:10]

    generated_at = summary.get("generated_at")
    if generated_at:
        try:
            return datetime.fromisoformat(str(generated_at).replace("Z", "+00:00")).date().isoformat()
        except ValueError:
            pass
    return datetime.now().date().isoformat()


def _summary_int(summary: dict[str, Any], *keys: str) -> int:
    for key in keys:
        value = summary.get(key)
        if value is None or value == "":
            continue
        try:
            return int(value)
        except (TypeError, ValueError):
            continue
    return 0


def _summary_float(summary: dict[str, Any], *keys: str) -> float:
    for key in keys:
        value = summary.get(key)
        if value is None or value == "":
            continue
        try:
            return float(value)
        except (TypeError, ValueError):
            continue
    return 0.0


def _summary_bool(summary: dict[str, Any], *keys: str) -> bool:
    for key in keys:
        value = summary.get(key)
        if isinstance(value, bool):
            return value
        if value is None or value == "":
            continue
        return str(value).strip().lower() in {"1", "true", "yes", "y", "on", "oui"}
    return False


def _find_report_attachment(artifacts: list[Path]) -> Path | None:
    for artifact in artifacts:
        path = Path(artifact)
        if path.suffix.lower() == ".xlsx" and path.name.startswith("Rapport_Reconciliation_RCI_"):
            return path
    for artifact in artifacts:
        path = Path(artifact)
        if path.suffix.lower() == ".xlsx":
            return path
    return None


def _recipients(value: Any) -> list[str]:
    if isinstance(value, str):
        raw_recipients = value.replace(";", ",").split(",")
    else:
        raw_recipients = list(value or [])
    return [str(recipient).strip() for recipient in raw_recipients if str(recipient).strip()]


def _env(name: str) -> str | None:
    value = os.getenv(name)
    if value is None or value.strip() == "":
        return None
    return value.strip()


def _env_bool(name: str, default: Any) -> bool:
    value = _env(name)
    if value is None:
        return _as_bool(default)
    return _as_bool(value)


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on", "oui"}
