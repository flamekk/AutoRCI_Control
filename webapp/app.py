from __future__ import annotations

import csv
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.action_plan import build_reference_suggestions, severity_for_status, sort_action_plan_records
from src.missing_rci import build_missing_rci_records

INPUT_DIRS = {
    "erp": PROJECT_ROOT / "input" / "erp",
    "rci": PROJECT_ROOT / "input" / "rci",
    "pdf": PROJECT_ROOT / "input" / "pdf",
    "reference": PROJECT_ROOT / "reference",
}
DOWNLOAD_DIRS = {
    "reports": PROJECT_ROOT / "output" / "reports",
    "anomalies": PROJECT_ROOT / "output" / "anomalies",
    "corrections": PROJECT_ROOT / "output" / "corrections",
    "powerbi": PROJECT_ROOT / "output" / "powerbi",
    "logs": PROJECT_ROOT / "logs",
}
ALLOWED_EXTENSIONS = {
    "erp": {".xlsx", ".xls", ".xlsm", ".csv"},
    "rci": {".txt", ".csv", ".xlsx", ".xls", ".xlsm"},
    "pdf": {".pdf"},
    "reference": {".xlsx", ".xls", ".xlsm"},
}
DOWNLOAD_EXTENSIONS = {
    "reports": {".xlsx"},
    "anomalies": {".csv", ".json"},
    "corrections": {".txt", ".csv"},
    "powerbi": {".csv"},
    "logs": {".log"},
}
DOWNLOAD_METADATA = {
    "reports": {"label": "Rapports Excel", "icon": "bi-file-earmark-spreadsheet"},
    "anomalies": {"label": "Audits", "icon": "bi-clipboard-data"},
    "corrections": {"label": "Corrections", "icon": "bi-file-earmark-check"},
    "powerbi": {"label": "Exports Power BI", "icon": "bi-bar-chart-line"},
    "logs": {"label": "Logs", "icon": "bi-terminal"},
}
SOURCE_METADATA = {
    "erp": {
        "label": "Export ERP Navision",
        "short_label": "ERP",
        "badge": "ERP",
        "icon": "bi-file-earmark-spreadsheet",
        "description": "Exports Excel ou CSV issus de Navision/Incadea contenant les factures et avoirs.",
    },
    "rci": {
        "label": "Flux RCI / Batch banque",
        "short_label": "RCI",
        "badge": "RCI",
        "icon": "bi-hdd-network",
        "description": "Fichiers TXT, CSV ou Excel représentant les flux transmis ou reçus côté RCI Banque.",
    },
    "pdf": {
        "label": "Rapports RCI / États de financement",
        "short_label": "PDF",
        "badge": "PDF",
        "icon": "bi-file-earmark-pdf",
        "description": "Rapports PDF quotidiens contenant les dates de facture, dates d’échéance, montants financés et origines.",
    },
    "reference": {
        "label": "Référentiel affaires couvertes RCI",
        "short_label": "Référentiel",
        "badge": "REF",
        "icon": "bi-database",
        "description": "Fichier Excel contenant la liste des affaires ou concessionnaires couverts par RCI.",
    },
}
TECHNICAL_NAMES = {".gitkeep", ".gitignore", ".ds_store", "__pycache__"}
RESULT_COLUMNS = [
    "invoice_number",
    "document_type",
    "erp_date",
    "customer_name",
    "amount_erp",
    "amount_rci",
    "amount_pdf",
    "amount_gap",
    "montant_impacte",
    "status",
    "action_recommandee",
]
ACTION_PLAN_COLUMNS = [
    "severity",
    "invoice_number",
    "erp_date",
    "customer_name",
    "amount_erp",
    "amount_rci",
    "amount_pdf",
    "montant_impacte",
    "status",
    "action_recommandee",
    "source_erp",
    "source_rci",
    "source_pdf",
]
MISSING_RCI_COLUMNS = [
    "categorie",
    "severity",
    "invoice_number",
    "erp_date",
    "due_date",
    "customer_name",
    "amount_erp",
    "montant_impacte",
    "action_recommandee",
    "included_in_corrective_batch",
]
DETAIL_STATUSES = {
    "MANQUANTE_RCI",
    "RCI_SEULEMENT",
    "ANOMALIE_MONTANT",
    "ANOMALIE_DATE",
    "DOUBLON",
}
ERP_ANALYZED_STATUSES = (
    "OK",
    "MANQUANTE_RCI",
    "HORS_SCOPE_RCI",
    "ANOMALIE_MONTANT",
    "ANOMALIE_DATE",
    "DOUBLON",
)
ERP_MATCHABLE_STATUSES = (
    "OK",
    "MANQUANTE_RCI",
    "ANOMALIE_MONTANT",
    "ANOMALIE_DATE",
    "DOUBLON",
)
GAP_STATUSES = (
    "MANQUANTE_RCI",
    "RCI_SEULEMENT",
    "ANOMALIE_MONTANT",
    "ANOMALIE_DATE",
    "DOUBLON",
)
STATUS_CHART_ORDER = (
    "OK",
    "MANQUANTE_RCI",
    "HORS_SCOPE_RCI",
    "RCI_HORS_PERIODE",
    "RCI_SEULEMENT",
    "ANOMALIE_MONTANT",
    "ANOMALIE_DATE",
    "DOUBLON",
)
SEVERITY_CHART_ORDER = (
    "CRITIQUE",
    "ELEVEE",
    "MOYENNE",
    "A_VERIFIER",
    "INFORMATION",
    "OK",
)
ACTION_SEVERITIES = {"CRITIQUE", "ELEVEE", "MOYENNE"}


app = Flask(__name__)
app.secret_key = os.getenv("AUTORCI_FLASK_SECRET", "autorcicontrol-local-dashboard")


@app.context_processor
def inject_globals() -> dict[str, Any]:
    return {
        "status_badge_class": status_badge_class,
        "severity_badge_class": severity_badge_class,
        "active_page": request.endpoint,
        "source_metadata": SOURCE_METADATA,
        "download_metadata": DOWNLOAD_METADATA,
    }


@app.template_filter("mad")
def format_money(value: Any) -> str:
    number = _to_float(value)
    return f"{number:,.2f} MAD" if number is not None else "-"


@app.template_filter("number")
def format_number(value: Any) -> str:
    number = _to_float(value)
    if number is None:
        return "-"
    if number.is_integer():
        return f"{int(number):,}"
    return f"{number:,.2f}"


@app.template_filter("percent")
def format_percent(value: Any) -> str:
    number = _to_float(value)
    return f"{number * 100:.2f} %" if number is not None else "-"


@app.template_filter("mtime")
def format_mtime(value: Any) -> str:
    try:
        return datetime.fromtimestamp(float(value)).strftime("%Y-%m-%d %H:%M")
    except (TypeError, ValueError, OSError):
        return "-"


@app.template_filter("filesize")
def format_filesize(value: Any) -> str:
    try:
        size = float(value)
    except (TypeError, ValueError):
        return "-"
    units = ["o", "Ko", "Mo", "Go"]
    index = 0
    while size >= 1024 and index < len(units) - 1:
        size /= 1024
        index += 1
    return f"{size:.1f} {units[index]}" if index else f"{int(size)} {units[index]}"


@app.route("/")
def dashboard() -> str:
    summary, summary_file = load_latest_summary()
    kpis = dashboard_kpis(summary)
    last_report = latest_file(DOWNLOAD_DIRS["reports"], "*.xlsx")
    last_log = latest_file(DOWNLOAD_DIRS["logs"], "*.log")
    return render_template(
        "dashboard.html",
        kpis=kpis,
        summary=summary,
        summary_file=summary_file,
        last_report=last_report,
        last_log=last_log,
        last_run_date=datetime.fromtimestamp(summary_file.stat().st_mtime) if summary_file else None,
        run_status=run_status_display(summary.get("status")),
        dashboard_alerts=dashboard_alerts(summary),
    )


@app.route("/upload", methods=["GET", "POST"])
def upload() -> str:
    ensure_directories()
    if request.method == "POST":
        source_type = request.form.get("source_type", "")
        uploaded_file = request.files.get("file")
        if source_type not in INPUT_DIRS:
            flash("Type de fichier inconnu.", "danger")
        elif not uploaded_file or uploaded_file.filename == "":
            flash("Aucun fichier sélectionné.", "warning")
        else:
            try:
                saved_path = save_upload(source_type, uploaded_file)
            except ValueError as exc:
                flash(str(exc), "danger")
            except OSError as exc:
                flash(f"Erreur lors de l'enregistrement : {exc}", "danger")
            else:
                flash(f"Fichier enregistré : {saved_path.name}", "success")
        return redirect(url_for("upload"))

    files = {source_type: list_files(folder) for source_type, folder in INPUT_DIRS.items()}
    return render_template("upload.html", files=files, allowed=ALLOWED_EXTENSIONS)


@app.route("/run", methods=["GET", "POST"])
def run() -> str:
    result = None
    form_data = {
        "dry_run": True,
        "no_archive": False,
        "ignore_pdf": False,
        "no_date_filter": False,
        "debug_reference": False,
        "date_from": "",
        "date_to": "",
        "debug_invoice": "",
    }
    if request.method == "POST":
        try:
            form_data = parse_run_form(request.form)
            result = run_pipeline(form_data)
        except ValueError as exc:
            result = {
                "ok": False,
                "returncode": None,
                "command": "",
                "stdout": "",
                "stderr": str(exc),
            }
        if result["ok"]:
            flash("Traitement terminé avec succès.", "success")
        else:
            flash("Traitement terminé en erreur. Consultez les logs ci-dessous.", "danger")
    return render_template("run.html", result=result, form_data=form_data)


@app.route("/results")
def results() -> str:
    rows, source_file = load_latest_reconciliation()
    filtered_rows, filters, options = filter_result_rows(rows, request.args)
    return render_template(
        "results.html",
        rows=filtered_rows,
        columns=RESULT_COLUMNS,
        source_file=source_file,
        filters=filters,
        options=options,
        title="Résultats de rapprochement",
    )


@app.route("/details")
def details() -> str:
    rows, source_file = load_latest_reconciliation()
    rows = [row for row in rows if row.get("status") in DETAIL_STATUSES]
    filtered_rows, filters, options = filter_result_rows(rows, request.args)
    return render_template(
        "details.html",
        rows=filtered_rows,
        columns=RESULT_COLUMNS,
        source_file=source_file,
        filters=filters,
        options=options,
        allowed_statuses=sorted(DETAIL_STATUSES),
    )


@app.route("/factures-absentes")
def factures_absentes() -> str:
    rows, source_file = load_latest_reconciliation()
    missing_rows = build_missing_rci_records(rows)
    filtered_rows, filters, options = filter_missing_rci_rows(missing_rows, request.args)
    run_id = run_id_from_reconciliation_file(source_file)
    latest_missing_export = file_for_run("anomalies", f"factures_avoirs_absents_RCI_{run_id}.csv") if run_id else None
    last_report = latest_file(DOWNLOAD_DIRS["reports"], "*.xlsx")
    latest_batch = file_for_run("corrections", f"batch_correctif_candidat_{run_id}.txt") if run_id else None
    return render_template(
        "factures_absentes.html",
        rows=filtered_rows,
        columns=MISSING_RCI_COLUMNS,
        source_file=source_file,
        filters=filters,
        options=options,
        kpis=missing_rci_kpis(missing_rows),
        latest_missing_export=latest_missing_export,
        last_report=last_report,
        latest_batch=latest_batch,
    )


@app.route("/action-plan")
def action_plan() -> str:
    rows, source_file = load_latest_reconciliation()
    action_rows = prepare_action_plan_rows(rows)
    filtered_rows, filters, options = filter_action_plan_rows(action_rows, request.args)
    reference_suggestions = load_latest_reference_suggestions()
    if not reference_suggestions:
        reference_suggestions = build_reference_suggestions(action_rows)
    run_id = run_id_from_reconciliation_file(source_file)
    latest_batch = file_for_run("corrections", f"batch_correctif_candidat_{run_id}.txt") if run_id else None
    latest_batch_control = (
        file_for_run("corrections", f"batch_correctif_candidat_{run_id}_control.csv")
        if run_id
        else None
    )
    return render_template(
        "action_plan.html",
        rows=filtered_rows,
        columns=ACTION_PLAN_COLUMNS,
        source_file=source_file,
        filters=filters,
        options=options,
        severity_kpis=action_plan_kpis(action_rows),
        reference_suggestions=reference_suggestions,
        latest_batch=latest_batch,
        latest_batch_control=latest_batch_control,
    )


@app.route("/historique")
def historique() -> str:
    rows, source_file = load_reconciliation_history()
    return render_template(
        "historique.html",
        source_file=source_file,
        total_rows=len(rows),
    )


@app.route("/reference-quality")
def reference_quality() -> str:
    reference_suggestions, source_file = load_reference_quality_rows()
    return render_template(
        "reference_quality.html",
        reference_suggestions=reference_suggestions,
        source_file=source_file,
    )


@app.route("/api/dashboard/charts")
def api_dashboard_charts():
    rows, _source_file = load_latest_reconciliation()
    return jsonify(build_dashboard_charts(rows))


@app.route("/api/action-plan/charts")
def api_action_plan_charts():
    rows, _source_file = load_latest_reconciliation()
    return jsonify(build_action_plan_charts(prepare_action_plan_rows(rows)))


@app.route("/api/history/charts")
def api_history_charts():
    rows, _source_file = load_reconciliation_history()
    return jsonify(build_history_charts(rows))


@app.route("/api/reference-quality/charts")
def api_reference_quality_charts():
    rows, _source_file = load_reference_quality_rows()
    return jsonify(build_reference_quality_charts(rows))


@app.route("/downloads")
def downloads() -> str:
    files = {
        category: list_downloads(category, folder)
        for category, folder in DOWNLOAD_DIRS.items()
    }
    return render_template("downloads.html", files=files)


@app.route("/download/<category>/<path:filename>")
def download_file(category: str, filename: str):
    try:
        path = resolve_download_path(category, filename)
    except FileNotFoundError:
        flash("Fichier introuvable ou non autorisé.", "danger")
        return redirect(url_for("downloads"))
    return send_file(path, as_attachment=True)


@app.route("/logs")
def logs() -> str:
    log_file = latest_file(DOWNLOAD_DIRS["logs"], "*.log")
    content = tail_file(log_file, max_lines=800) if log_file else ""
    return render_template("logs.html", log_file=log_file, content=content)


def ensure_directories() -> None:
    for folder in [*INPUT_DIRS.values(), *DOWNLOAD_DIRS.values()]:
        folder.mkdir(parents=True, exist_ok=True)


def save_upload(source_type: str, uploaded_file: Any) -> Path:
    filename = safe_original_filename(uploaded_file.filename or "")
    if not filename:
        raise ValueError("Nom de fichier invalide.")

    extension = Path(filename).suffix.lower()
    if extension not in ALLOWED_EXTENSIONS[source_type]:
        allowed = ", ".join(sorted(ALLOWED_EXTENSIONS[source_type]))
        raise ValueError(f"Extension non autorisée pour {source_type.upper()} : {extension}. Autorisées : {allowed}")

    target_dir = INPUT_DIRS[source_type]
    target_dir.mkdir(parents=True, exist_ok=True)
    target_path = unique_destination(target_dir / filename)
    uploaded_file.save(target_path)
    return target_path


def safe_original_filename(value: str) -> str:
    filename = value.replace("\\", "/").split("/")[-1].strip()
    if filename in {"", ".", ".."}:
        return ""
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", filename)


def unique_destination(path: Path) -> Path:
    if not path.exists():
        return path
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_{timestamp}{path.suffix}")


def parse_run_form(form: Any) -> dict[str, Any]:
    debug_invoice = str(form.get("debug_invoice", "") or "").strip().upper()
    if debug_invoice and not re.fullmatch(r"[A-Z0-9_-]{1,40}", debug_invoice):
        raise ValueError("Debug invoice invalide : seuls lettres, chiffres, tirets et underscores sont autorisés.")
    return {
        "date_from": str(form.get("date_from", "") or "").strip(),
        "date_to": str(form.get("date_to", "") or "").strip(),
        "dry_run": form.get("dry_run") == "on",
        "no_archive": form.get("no_archive") == "on",
        "ignore_pdf": form.get("ignore_pdf") == "on",
        "no_date_filter": form.get("no_date_filter") == "on",
        "debug_reference": form.get("debug_reference") == "on",
        "debug_invoice": debug_invoice,
    }


def run_pipeline(form_data: dict[str, Any]) -> dict[str, Any]:
    command = [sys.executable, str(PROJECT_ROOT / "src" / "main.py")]
    if form_data.get("dry_run"):
        command.append("--dry-run")
    if form_data.get("no_archive"):
        command.append("--no-archive")
    if form_data.get("ignore_pdf"):
        command.append("--ignore-pdf")
    if form_data.get("no_date_filter"):
        command.append("--no-date-filter")
    else:
        date_from = form_data.get("date_from")
        date_to = form_data.get("date_to")
        if bool(date_from) ^ bool(date_to):
            raise ValueError("Les deux dates doivent être renseignées ensemble.")
        if date_from and date_to:
            command.extend(["--date-from", date_from, "--date-to", date_to])
    if form_data.get("debug_reference"):
        command.append("--debug-reference")
    if form_data.get("debug_invoice"):
        command.extend(["--debug-invoice", form_data["debug_invoice"]])

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    try:
        completed = subprocess.run(
            command,
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=900,
            env=env,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        return {
            "ok": False,
            "returncode": None,
            "command": " ".join(command),
            "stdout": exc.stdout or "",
            "stderr": "Le traitement a dépassé le délai maximum de 15 minutes.",
        }

    return {
        "ok": completed.returncode == 0,
        "returncode": completed.returncode,
        "command": " ".join(command),
        "stdout": completed.stdout,
        "stderr": completed.stderr,
    }


def load_latest_summary() -> tuple[dict[str, Any], Path | None]:
    summary_file = latest_file(DOWNLOAD_DIRS["powerbi"], "run_summary_*.csv")
    if summary_file:
        rows = read_csv_rows(summary_file)
        summary = parse_run_summary_rows(rows)
        return normalize_dashboard_summary(summary), summary_file
    return {}, None


def parse_run_summary_rows(rows: list[dict[str, str]]) -> dict[str, Any]:
    if not rows:
        return {}

    first_row = rows[0]
    if "metric" in first_row and "value" in first_row:
        selected_rows = select_latest_run_rows(rows)
        return {
            row.get("metric", ""): row.get("value", "")
            for row in selected_rows
            if row.get("metric")
        }

    selected_row = select_latest_run_rows(rows)[-1]
    return {key: value for key, value in selected_row.items() if key}


def select_latest_run_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    if not rows:
        return []
    latest_run_id = rows[-1].get("processing_run_id")
    if latest_run_id:
        run_rows = [row for row in rows if row.get("processing_run_id") == latest_run_id]
        return run_rows or [rows[-1]]
    return rows


def normalize_dashboard_summary(summary: dict[str, Any]) -> dict[str, Any]:
    normalized = dict(summary)
    counts = {status: summary_status_count(summary, status) for status in set(ERP_ANALYZED_STATUSES) | {"RCI_SEULEMENT", "RCI_HORS_PERIODE"}}

    analyzed = sum(counts[status] for status in ERP_ANALYZED_STATUSES)
    matchable = sum(counts[status] for status in ERP_MATCHABLE_STATUSES)
    gaps = sum(counts[status] for status in GAP_STATUSES)
    ok_count = counts["OK"]
    rci_after_filter = _summary_int(summary, "rci_rows_after_date_filter")
    rci_before_filter = _summary_int(summary, "rci_rows_before_date_filter")
    no_rci_flux_in_period = _summary_bool(summary, "no_rci_flux_in_period_alert") or (
        rci_before_filter > 0 and rci_after_filter == 0
    )

    normalized.update(
        {
            "erp_analyzed_invoices": analyzed,
            "reconciled_invoices": analyzed,
            "erp_matchable_invoices": matchable,
            "matched_invoices": ok_count,
            "unmatched_erp": counts["MANQUANTE_RCI"],
            "out_of_scope_rci": counts["HORS_SCOPE_RCI"],
            "amount_anomalies": counts["ANOMALIE_MONTANT"],
            "date_anomalies": counts["ANOMALIE_DATE"],
            "duplicates": counts["DOUBLON"],
            "unmatched_rci": counts["RCI_SEULEMENT"],
            "rci_out_of_period": counts["RCI_HORS_PERIODE"],
            "gaps_detected": gaps,
            "dashboard_no_rci_flux_in_period": no_rci_flux_in_period,
            "dashboard_period_mismatch": counts["RCI_HORS_PERIODE"] > 0 and ok_count == 0,
        }
    )

    if no_rci_flux_in_period:
        normalized["matching_rate"] = None
    else:
        normalized["matching_rate"] = ok_count / matchable if matchable else 0
    return normalized


def summarize_history(path: Path) -> dict[str, Any]:
    rows = read_csv_rows(path)
    if not rows:
        return {}
    latest_run = rows[-1].get("processing_run_id", "")
    run_rows = [row for row in rows if row.get("processing_run_id") == latest_run] if latest_run else rows
    counts: dict[str, int] = {}
    for row in run_rows:
        status = row.get("status", "")
        counts[status] = counts.get(status, 0) + 1
    matchable = sum(counts.get(status, 0) for status in DETAIL_STATUSES | {"OK"})
    analyzed = sum(counts.get(status, 0) for status in ERP_ANALYZED_STATUSES)
    ok = counts.get("OK", 0)
    return {
        "erp_analyzed_invoices": analyzed,
        "reconciled_invoices": analyzed,
        "erp_matchable_invoices": matchable,
        "out_of_scope_rci": counts.get("HORS_SCOPE_RCI", 0),
        "matched_invoices": ok,
        "unmatched_erp": counts.get("MANQUANTE_RCI", 0),
        "rci_out_of_period": counts.get("RCI_HORS_PERIODE", 0),
        "gaps_detected": sum(counts.get(status, 0) for status in DETAIL_STATUSES),
        "total_impacted_amount": sum(_to_float(row.get("montant_impacte")) or 0 for row in run_rows),
        "matching_rate": ok / matchable if matchable else 0,
    }


def dashboard_kpis(summary: dict[str, Any]) -> list[dict[str, Any]]:
    matching_rate = _summary_value(summary, "matching_rate")
    rate_kpi = {
        "label": "Taux de rapprochement",
        "value": matching_rate,
        "kind": "percent",
        "icon": "bi-speedometer2",
        "description": "OK / factures dans le périmètre",
    }
    if matching_rate is None:
        rate_kpi["display"] = "N/A"

    kpis = [
        {"label": "Factures ERP analysées", "value": _summary_value(summary, "erp_analyzed_invoices"), "kind": "number", "icon": "bi-receipt", "description": "Volume ERP filtré"},
        {"label": "Factures dans le périmètre RCI", "value": _summary_value(summary, "erp_matchable_invoices"), "kind": "number", "icon": "bi-bullseye", "description": "Base contrôlable"},
        {"label": "Factures OK", "value": _summary_value(summary, "matched_invoices"), "kind": "number", "status": "OK", "icon": "bi-check-circle", "description": "Rapprochées sans écart"},
        {"label": "Factures manquantes RCI", "value": _summary_value(summary, "unmatched_erp"), "kind": "number", "status": "MANQUANTE_RCI", "icon": "bi-exclamation-triangle", "description": "ERP présentes, RCI absent"},
        {"label": "Factures hors périmètre RCI", "value": _summary_value(summary, "out_of_scope_rci"), "kind": "number", "status": "HORS_SCOPE_RCI", "icon": "bi-slash-circle", "description": "Non couvertes par le référentiel"},
        {"label": "RCI hors période", "value": _summary_value(summary, "rci_out_of_period"), "kind": "number", "status": "RCI_HORS_PERIODE", "icon": "bi-calendar-x", "description": "Flux chargés hors période"},
        {"label": "Écarts détectés", "value": _summary_value(summary, "gaps_detected"), "kind": "number", "status": "ANOMALIE_MONTANT", "icon": "bi-exclamation-octagon", "description": "Écarts métiers à traiter"},
        {"label": "Montant impacté total", "value": _summary_value(summary, "total_impacted_amount"), "kind": "money", "icon": "bi-cash-coin", "description": "Impact financier estimé"},
        rate_kpi,
    ]

    severity_kpis = [
        {"label": "Écarts critiques", "value": first_summary_int(summary, ["critical_gaps", "severity_critique"]), "kind": "number", "severity": "CRITIQUE", "icon": "bi-fire", "description": "Priorité immédiate"},
        {"label": "Écarts élevés", "value": first_summary_int(summary, ["high_gaps", "severity_elevee"]), "kind": "number", "severity": "ELEVEE", "icon": "bi-arrow-up-circle", "description": "Priorité forte"},
        {"label": "Écarts moyens", "value": first_summary_int(summary, ["medium_gaps", "severity_moyenne"]), "kind": "number", "severity": "MOYENNE", "icon": "bi-dash-circle", "description": "À suivre"},
    ]
    if any(kpi["value"] for kpi in severity_kpis):
        kpis.extend(severity_kpis)

    return kpis


def dashboard_alerts(summary: dict[str, Any]) -> list[dict[str, str]]:
    alerts = []
    if _summary_bool(summary, "dashboard_no_rci_flux_in_period"):
        alerts.append(
            {
                "class": "warning",
                "icon": "bi-exclamation-triangle-fill",
                "message": "Aucun flux RCI dans la période sélectionnée. Les lignes RCI chargées sont hors période.",
            }
        )
    if _summary_bool(summary, "dashboard_period_mismatch"):
        alerts.append(
            {
                "class": "danger",
                "icon": "bi-calendar-x-fill",
                "message": "Attention : le fichier RCI chargé ne correspond pas à la période sélectionnée.",
            }
        )
    if _summary_bool(summary, "pdf_period_mismatch_alert"):
        alerts.append(
            {
                "class": "warning",
                "icon": "bi-file-earmark-pdf",
                "message": "Attention : les PDF chargés ne correspondent pas à la période du flux RCI.",
            }
        )
    rci_out_of_period = _summary_int(summary, "rci_out_of_period")
    if rci_out_of_period > 0 and not _summary_bool(summary, "dashboard_period_mismatch"):
        alerts.append(
            {
                "class": "info",
                "icon": "bi-info-circle-fill",
                "message": f"{rci_out_of_period} ligne(s) RCI/PDF sont hors période de rapprochement.",
            }
        )
    matching_rate = _summary_value(summary, "matching_rate")
    if matching_rate is not None and (
        _summary_bool(summary, "low_matching_rate_alert")
        or ((_to_float(matching_rate) or 0) < 0.70 and _summary_int(summary, "erp_matchable_invoices") > 0)
    ):
        alerts.append(
            {
                "class": "warning",
                "icon": "bi-speedometer",
                "message": "Le taux de rapprochement est inférieur au seuil de vigilance de 70 %.",
            }
        )
    critical_gaps = first_summary_int(summary, ["critical_gaps", "severity_critique"])
    if critical_gaps > 0:
        alerts.append(
            {
                "class": "danger",
                "icon": "bi-fire",
                "message": f"{critical_gaps} écart(s) critique(s) nécessitent une action prioritaire.",
            }
        )
    return alerts


def run_status_display(status: Any) -> dict[str, str]:
    raw_status = str(status or "").strip()
    normalized = raw_status.lower()
    if not normalized:
        return {"label": "Non disponible", "class": "secondary"}
    if "completed_with_anomalies" in normalized or "anomal" in normalized or "écart" in normalized or "ecart" in normalized:
        return {"label": "Traitement avec écarts", "class": "warning"}
    if "error" in normalized or "erreur" in normalized or "failed" in normalized or "fail" in normalized:
        return {"label": "Erreur", "class": "danger"}
    if normalized in {"ok", "success", "completed", "complete"} or "success" in normalized:
        return {"label": "Traitement OK", "class": "success"}
    return {"label": raw_status, "class": "secondary"}


def load_latest_reconciliation() -> tuple[list[dict[str, Any]], Path | None]:
    path = latest_reconciliation_file()
    if not path:
        return [], None
    return read_csv_rows(path), path


def load_reconciliation_history() -> tuple[list[dict[str, Any]], Path | None]:
    path = DOWNLOAD_DIRS["powerbi"] / "reconciliation_history.csv"
    if not path.exists() or not is_visible_path(path):
        return [], None
    return read_csv_rows(path), path


def load_latest_batch_control() -> tuple[list[dict[str, Any]], Path | None]:
    path = latest_file(DOWNLOAD_DIRS["corrections"], "batch_correctif_candidat_*_control.csv")
    if not path:
        return [], None
    return read_csv_rows(path), path


def load_reference_quality_rows() -> tuple[list[dict[str, Any]], Path | None]:
    path = latest_file(DOWNLOAD_DIRS["anomalies"], "reference_suggestions_*.csv")
    if path:
        return read_csv_rows(path), path

    rows, source_file = load_latest_reconciliation()
    if not rows:
        return [], None
    return build_reference_suggestions(prepare_action_plan_rows(rows)), source_file


def latest_reconciliation_file() -> Path | None:
    folder = DOWNLOAD_DIRS["powerbi"]
    if not folder.exists():
        return None
    files = [
        path
        for path in folder.glob("reconciliation_*.csv")
        if path.name != "reconciliation_history.csv"
    ]
    files = sorted(files, key=lambda path: path.stat().st_mtime, reverse=True)
    return files[0] if files else None


def filter_result_rows(rows: list[dict[str, Any]], args: Any) -> tuple[list[dict[str, Any]], dict[str, str], dict[str, list[str]]]:
    filters = {
        "status": str(args.get("status", "") or ""),
        "customer_name": str(args.get("customer_name", "") or ""),
        "date_from": str(args.get("date_from", "") or ""),
        "date_to": str(args.get("date_to", "") or ""),
    }
    status_options = sorted({row.get("status", "") for row in rows if row.get("status")})
    customer_options = sorted({row.get("customer_name", "") for row in rows if row.get("customer_name")})
    filtered = []
    for row in rows:
        if filters["status"] and row.get("status") != filters["status"]:
            continue
        if filters["customer_name"] and filters["customer_name"].upper() not in str(row.get("customer_name", "")).upper():
            continue
        row_date = row.get("erp_date") or row.get("rci_date") or row.get("pdf_invoice_date") or ""
        if filters["date_from"] and row_date and row_date < filters["date_from"]:
            continue
        if filters["date_to"] and row_date and row_date > filters["date_to"]:
            continue
        filtered.append(row)
    return filtered, filters, {"statuses": status_options, "customers": customer_options}


def prepare_action_plan_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    prepared = []
    for row in rows:
        enriched = dict(row)
        enriched["severity"] = enriched.get("severity") or severity_for_status(
            enriched.get("status"),
            enriched.get("montant_impacte"),
        )
        prepared.append(enriched)
    return [
        row for row in sort_action_plan_records(prepared)
        if row.get("severity") != "INFORMATION"
    ]


def filter_action_plan_rows(
    rows: list[dict[str, Any]], args: Any
) -> tuple[list[dict[str, Any]], dict[str, str], dict[str, list[str]]]:
    filters = {
        "severity": str(args.get("severity", "") or ""),
        "status": str(args.get("status", "") or ""),
        "customer_name": str(args.get("customer_name", "") or ""),
    }
    severity_options = sorted(
        {row.get("severity", "") for row in rows if row.get("severity")},
        key=lambda value: {"CRITIQUE": 0, "ELEVEE": 1, "MOYENNE": 2, "A_VERIFIER": 3, "INFORMATION": 4}.get(value, 99),
    )
    status_options = sorted({row.get("status", "") for row in rows if row.get("status")})
    customer_options = sorted({row.get("customer_name", "") for row in rows if row.get("customer_name")})

    filtered = []
    for row in rows:
        if filters["severity"] and row.get("severity") != filters["severity"]:
            continue
        if filters["status"] and row.get("status") != filters["status"]:
            continue
        if filters["customer_name"] and filters["customer_name"].upper() not in str(row.get("customer_name", "")).upper():
            continue
        filtered.append(row)
    return filtered, filters, {"severities": severity_options, "statuses": status_options, "customers": customer_options}


def filter_missing_rci_rows(
    rows: list[dict[str, Any]], args: Any
) -> tuple[list[dict[str, Any]], dict[str, str], dict[str, list[str]]]:
    filters = {
        "categorie": str(args.get("categorie", "") or ""),
        "severity": str(args.get("severity", "") or ""),
        "customer_name": str(args.get("customer_name", "") or ""),
        "date_from": str(args.get("date_from", "") or ""),
        "date_to": str(args.get("date_to", "") or ""),
        "included_in_corrective_batch": str(args.get("included_in_corrective_batch", "") or ""),
    }
    category_options = sorted({row.get("categorie", "") for row in rows if row.get("categorie")})
    severity_options = sorted(
        {row.get("severity", "") for row in rows if row.get("severity")},
        key=lambda value: {"CRITIQUE": 0, "ELEVEE": 1, "MOYENNE": 2}.get(value, 99),
    )
    customer_options = sorted({row.get("customer_name", "") for row in rows if row.get("customer_name")})

    filtered = []
    for row in rows:
        if filters["categorie"] and row.get("categorie") != filters["categorie"]:
            continue
        if filters["severity"] and row.get("severity") != filters["severity"]:
            continue
        if filters["customer_name"] and filters["customer_name"].upper() not in str(row.get("customer_name", "")).upper():
            continue
        row_date = str(row.get("erp_date") or "")
        if filters["date_from"] and row_date and row_date < filters["date_from"]:
            continue
        if filters["date_to"] and row_date and row_date > filters["date_to"]:
            continue
        included_filter = filters["included_in_corrective_batch"]
        if included_filter:
            included = _as_bool(row.get("included_in_corrective_batch"))
            if included_filter == "true" and not included:
                continue
            if included_filter == "false" and included:
                continue
        filtered.append(row)
    return filtered, filters, {
        "categories": category_options,
        "severities": severity_options,
        "customers": customer_options,
    }


def missing_rci_kpis(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    facture_count = sum(1 for row in rows if row.get("categorie") == "Facture absente")
    avoir_count = sum(1 for row in rows if row.get("categorie") == "Avoir absent")
    total_amount = sum(abs(_to_float(row.get("montant_impacte")) or 0.0) for row in rows)
    severity_counts = {"CRITIQUE": 0, "ELEVEE": 0, "MOYENNE": 0}
    for row in rows:
        severity = str(row.get("severity") or "")
        if severity in severity_counts:
            severity_counts[severity] += 1

    return [
        {"label": "Factures absentes RCI", "value": facture_count, "kind": "number", "status": "MANQUANTE_RCI", "icon": "bi-receipt"},
        {"label": "Avoirs absents RCI", "value": avoir_count, "kind": "number", "status": "MANQUANTE_RCI", "icon": "bi-receipt-cutoff"},
        {"label": "Montant total absent", "value": round(total_amount, 2), "kind": "money", "icon": "bi-cash-coin"},
        {"label": "Écarts critiques", "value": severity_counts["CRITIQUE"], "kind": "number", "severity": "CRITIQUE", "icon": "bi-fire"},
        {"label": "Écarts élevés", "value": severity_counts["ELEVEE"], "kind": "number", "severity": "ELEVEE", "icon": "bi-arrow-up-circle"},
        {"label": "Écarts moyens", "value": severity_counts["MOYENNE"], "kind": "number", "severity": "MOYENNE", "icon": "bi-dash-circle"},
    ]


def action_plan_kpis(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    counts = {
        "CRITIQUE": 0,
        "ELEVEE": 0,
        "MOYENNE": 0,
        "A_VERIFIER": 0,
        "INFORMATION": 0,
    }
    for row in rows:
        severity = str(row.get("severity") or "")
        if severity in counts:
            counts[severity] += 1
    return [
        {"label": "Critiques", "value": counts["CRITIQUE"], "severity": "CRITIQUE", "icon": "bi-fire"},
        {"label": "Élevées", "value": counts["ELEVEE"], "severity": "ELEVEE", "icon": "bi-arrow-up-circle"},
        {"label": "Moyennes", "value": counts["MOYENNE"], "severity": "MOYENNE", "icon": "bi-dash-circle"},
        {"label": "À vérifier", "value": counts["A_VERIFIER"], "severity": "A_VERIFIER", "icon": "bi-search"},
        {"label": "Information", "value": counts["INFORMATION"], "severity": "INFORMATION", "icon": "bi-info-circle"},
    ]


def build_dashboard_charts(rows: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "status_distribution": _chart_from_counts(
            _ordered_counts(rows, "status", STATUS_CHART_ORDER),
            label="Nombre de lignes",
            empty_on_zero=False,
        ),
        "impacted_amount_by_status": _chart_from_values(
            _sum_by_field(rows, "status", "montant_impacte", STATUS_CHART_ORDER),
            label="Montant impacté",
            value_format="money",
        ),
        "severity_distribution": _chart_from_counts(
            _severity_counts(rows, include_ok=True),
            label="Nombre de lignes",
            empty_on_zero=False,
        ),
    }


def build_action_plan_charts(rows: list[dict[str, Any]]) -> dict[str, Any]:
    impacted_rows = [
        row
        for row in rows
        if row.get("status") == "MANQUANTE_RCI" or str(row.get("severity") or "") in ACTION_SEVERITIES
    ]
    non_ok_rows = [row for row in rows if str(row.get("severity") or "") != "OK"]
    return {
        "top_customers_impacted_amount": _chart_from_values(
            _top_items(_sum_by_customer(impacted_rows, "montant_impacte"), limit=10),
            label="Montant impacté",
            value_format="money",
        ),
        "top_customers_gap_count": _chart_from_counts(
            _top_items(_count_by_customer(non_ok_rows), limit=10),
            label="Nombre d'écarts",
            empty_on_zero=False,
        ),
        "severity_gap_distribution": _chart_from_counts(
            _severity_counts(non_ok_rows, include_ok=False),
            label="Nombre de lignes",
            empty_on_zero=False,
        ),
    }


def build_history_charts(rows: list[dict[str, Any]]) -> dict[str, Any]:
    grouped: dict[str, dict[str, Any]] = {}
    for row in rows:
        date_key = _processing_day(row.get("processing_date"))
        if not date_key:
            continue
        bucket = grouped.setdefault(
            date_key,
            {
                "impacted_amount": 0.0,
                "gap_count": 0,
                "ok_count": 0,
                "matchable_count": 0,
            },
        )
        status = str(row.get("status") or "").strip()
        bucket["impacted_amount"] += _to_float(row.get("montant_impacte")) or 0.0
        if status in GAP_STATUSES:
            bucket["gap_count"] += 1
        if status == "OK":
            bucket["ok_count"] += 1
        if status in ERP_MATCHABLE_STATUSES:
            bucket["matchable_count"] += 1

    labels = sorted(grouped)
    impacted_values = [round(grouped[label]["impacted_amount"], 2) for label in labels]
    gap_values = [grouped[label]["gap_count"] for label in labels]
    rate_values = [
        round(grouped[label]["ok_count"] / grouped[label]["matchable_count"], 4)
        if grouped[label]["matchable_count"]
        else None
        for label in labels
    ]

    return {
        "impacted_amount_trend": _chart_payload(
            labels,
            impacted_values,
            label="Montant impacté",
            value_format="money",
            empty=not labels or not any(value for value in impacted_values),
        ),
        "gap_count_trend": _chart_payload(
            labels,
            gap_values,
            label="Écarts détectés",
            value_format="number",
            empty=not labels,
        ),
        "matching_rate_trend": _chart_payload(
            labels,
            rate_values,
            label="Taux de rapprochement",
            value_format="percent",
            empty=not labels or not any(value is not None for value in rate_values),
        ),
    }


def build_reference_quality_charts(rows: list[dict[str, Any]]) -> dict[str, Any]:
    count_values: dict[str, float] = {}
    amount_values: dict[str, float] = {}
    for row in rows:
        customer_name = _customer_name(row)
        if not customer_name:
            continue
        count_values[customer_name] = count_values.get(customer_name, 0.0) + (
            _to_float(row.get("nombre_factures")) or 0.0
        )
        amount_values[customer_name] = amount_values.get(customer_name, 0.0) + (
            _to_float(row.get("montant_total_erp")) or _to_float(row.get("amount_erp")) or 0.0
        )

    return {
        "top_out_scope_count": _chart_from_counts(
            _top_items(count_values, limit=10),
            label="Nombre de lignes",
            empty_on_zero=False,
        ),
        "top_out_scope_amount": _chart_from_values(
            _top_items(amount_values, limit=10),
            label="Montant ERP",
            value_format="money",
        ),
    }


def _ordered_counts(
    rows: list[dict[str, Any]],
    field: str,
    order: tuple[str, ...],
) -> dict[str, float]:
    counts = {key: 0.0 for key in order}
    extras: dict[str, float] = {}
    for row in rows:
        key = str(row.get(field) or "").strip()
        if not key:
            continue
        target = counts if key in counts else extras
        target[key] = target.get(key, 0.0) + 1
    ordered = {key: counts[key] for key in order if counts.get(key)}
    ordered.update({key: extras[key] for key in sorted(extras) if extras[key]})
    return ordered


def _sum_by_field(
    rows: list[dict[str, Any]],
    group_field: str,
    amount_field: str,
    order: tuple[str, ...],
) -> dict[str, float]:
    values = {key: 0.0 for key in order}
    extras: dict[str, float] = {}
    for row in rows:
        key = str(row.get(group_field) or "").strip()
        if not key:
            continue
        amount = _to_float(row.get(amount_field)) or 0.0
        target = values if key in values else extras
        target[key] = target.get(key, 0.0) + amount
    ordered = {key: round(values[key], 2) for key in order if values.get(key)}
    ordered.update({key: round(extras[key], 2) for key in sorted(extras) if extras[key]})
    return ordered


def _severity_counts(rows: list[dict[str, Any]], include_ok: bool) -> dict[str, float]:
    counts = {key: 0.0 for key in SEVERITY_CHART_ORDER}
    for row in rows:
        severity = str(row.get("severity") or "").strip()
        if not severity:
            severity = severity_for_status(row.get("status"), row.get("montant_impacte"))
        if severity == "OK" and not include_ok:
            continue
        if severity in counts:
            counts[severity] += 1
    return {key: counts[key] for key in SEVERITY_CHART_ORDER if counts.get(key)}


def _sum_by_customer(rows: list[dict[str, Any]], amount_field: str) -> dict[str, float]:
    grouped: dict[str, float] = {}
    for row in rows:
        customer_name = _customer_name(row)
        if not customer_name:
            continue
        grouped[customer_name] = grouped.get(customer_name, 0.0) + (_to_float(row.get(amount_field)) or 0.0)
    return {key: round(value, 2) for key, value in grouped.items() if value}


def _count_by_customer(rows: list[dict[str, Any]]) -> dict[str, float]:
    grouped: dict[str, float] = {}
    for row in rows:
        customer_name = _customer_name(row)
        if not customer_name:
            continue
        grouped[customer_name] = grouped.get(customer_name, 0.0) + 1
    return grouped


def _top_items(values: dict[str, float], limit: int = 10) -> dict[str, float]:
    sorted_items = sorted(values.items(), key=lambda item: abs(item[1]), reverse=True)
    return dict(sorted_items[:limit])


def _chart_from_counts(
    values: dict[str, float],
    label: str,
    empty_on_zero: bool,
) -> dict[str, Any]:
    return _chart_payload(
        list(values.keys()),
        [int(value) for value in values.values()],
        label=label,
        value_format="number",
        empty=not values if not empty_on_zero else not any(values.values()),
    )


def _chart_from_values(
    values: dict[str, float],
    label: str,
    value_format: str,
) -> dict[str, Any]:
    return _chart_payload(
        list(values.keys()),
        [round(value, 2) for value in values.values()],
        label=label,
        value_format=value_format,
        empty=not values or not any(value for value in values.values()),
    )


def _chart_payload(
    labels: list[str],
    values: list[Any],
    label: str,
    value_format: str,
    empty: bool,
) -> dict[str, Any]:
    return {
        "labels": labels,
        "values": values,
        "label": label,
        "format": value_format,
        "empty": bool(empty),
    }


def _customer_name(row: dict[str, Any]) -> str:
    return str(row.get("customer_name") or row.get("normalized_customer_name") or "Non renseigné").strip()


def _processing_day(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    return text[:10]


def load_latest_reference_suggestions() -> list[dict[str, Any]]:
    path = latest_file(DOWNLOAD_DIRS["anomalies"], "reference_suggestions_*.csv")
    return read_csv_rows(path) if path else []


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    sample = text[:2048]
    delimiter = ";" if sample.count(";") >= sample.count(",") else ","
    reader = csv.DictReader(text.splitlines(), delimiter=delimiter)
    return [dict(row) for row in reader]


def latest_file(folder: Path, pattern: str) -> Path | None:
    if not folder.exists():
        return None
    files = sorted(
        [path for path in folder.glob(pattern) if is_visible_path(path)],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    return files[0] if files else None


def run_id_from_reconciliation_file(path: Path | None) -> str:
    if path is None:
        return ""
    match = re.fullmatch(r"reconciliation_(.+)\.csv", path.name)
    return match.group(1) if match else ""


def file_for_run(category: str, filename: str) -> Path | None:
    folder = DOWNLOAD_DIRS.get(category)
    if folder is None:
        return None
    path = folder / filename
    if path.exists() and path.is_file() and is_visible_path(path):
        return path
    return None


def list_files(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    return sorted(
        [path for path in folder.iterdir() if path.is_file() and is_visible_path(path)],
        key=lambda path: path.name.lower(),
    )


def list_downloads(category: str, folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    allowed = DOWNLOAD_EXTENSIONS[category]
    files = [
        path
        for path in folder.iterdir()
        if path.is_file()
        and path.suffix.lower() in allowed
        and is_visible_path(path)
    ]
    return sorted(files, key=lambda path: path.stat().st_mtime, reverse=True)[:30]


def is_visible_path(path: Path) -> bool:
    for part in path.parts:
        lowered = part.lower()
        if lowered in TECHNICAL_NAMES:
            return False
        if part.startswith("."):
            return False
    return True


def resolve_download_path(category: str, filename: str) -> Path:
    if category not in DOWNLOAD_DIRS:
        raise FileNotFoundError("Catégorie inconnue.")
    base = DOWNLOAD_DIRS[category].resolve()
    path = (base / filename).resolve()
    if not path.is_file() or base not in path.parents:
        raise FileNotFoundError("Fichier introuvable.")
    if path.suffix.lower() not in DOWNLOAD_EXTENSIONS[category]:
        raise FileNotFoundError("Extension non autorisée.")
    return path


def tail_file(path: Path, max_lines: int = 500) -> str:
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    return "\n".join(lines[-max_lines:])


def status_badge_class(status: str | None) -> str:
    mapping = {
        "OK": "success",
        "MANQUANTE_RCI": "warning",
        "ANOMALIE_MONTANT": "danger",
        "ANOMALIE_DATE": "danger",
        "DOUBLON": "danger",
        "HORS_SCOPE_RCI": "secondary",
        "RCI_HORS_PERIODE": "info",
        "RCI_SEULEMENT": "purple",
    }
    return mapping.get(str(status or ""), "secondary")


def severity_badge_class(severity: str | None) -> str:
    mapping = {
        "CRITIQUE": "danger",
        "ELEVEE": "warning",
        "MOYENNE": "medium",
        "A_VERIFIER": "secondary",
        "INFORMATION": "info",
        "OK": "success",
    }
    return mapping.get(str(severity or ""), "secondary")


def summary_status_count(summary: dict[str, Any], status: str) -> int:
    fallback_keys = {
        "OK": ["matched_invoices", "ok"],
        "MANQUANTE_RCI": ["unmatched_erp", "missing_rci", "manquante_rci"],
        "HORS_SCOPE_RCI": ["out_of_scope_rci", "hors_scope_rci"],
        "ANOMALIE_MONTANT": ["amount_anomalies", "anomalie_montant"],
        "ANOMALIE_DATE": ["date_anomalies", "anomalie_date"],
        "DOUBLON": ["duplicates", "doublon"],
        "RCI_SEULEMENT": ["unmatched_rci", "rci_seulement"],
        "RCI_HORS_PERIODE": [
            "total_rci_pdf_out_of_period",
            "rci_out_of_period",
            "status_rci_hors_periode",
            "rci_pdf_rows_excluded_by_date",
            "rci_rows_excluded_by_date",
            "rci_txt_rows_excluded_by_date",
        ],
    }
    status_key = f"status_{status.lower()}"
    keys = fallback_keys.get(status, [])
    if status != "RCI_HORS_PERIODE":
        keys = [status_key, *keys]
    return first_summary_int(summary, keys)


def first_summary_int(summary: dict[str, Any], keys: list[str]) -> int:
    for key in keys:
        if key in summary and summary.get(key) not in {None, ""}:
            return _summary_int(summary, key)
    return 0


def _summary_int(summary: dict[str, Any], key: str) -> int:
    value = _to_float(summary.get(key))
    return int(value) if value is not None else 0


def _summary_bool(summary: dict[str, Any], key: str) -> bool:
    value = summary.get(key)
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    text = str(value).strip().lower()
    if text in {"true", "1", "yes", "oui", "y"}:
        return True
    if text in {"false", "0", "no", "non", "n", ""}:
        return False
    return bool(text)


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value or "").strip().lower()
    return text in {"true", "1", "yes", "oui", "y", "on"}


def _summary_value(summary: dict[str, Any], key: str) -> Any:
    return summary.get(key, 0)


def _to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None


if __name__ == "__main__":
    ensure_directories()
    port = int(os.getenv("AUTORCI_FLASK_PORT", os.getenv("PORT", "5000")))
    app.run(host="127.0.0.1", port=port, debug=False)
