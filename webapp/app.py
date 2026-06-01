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
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
INPUT_DIRS = {
    "erp": PROJECT_ROOT / "input" / "erp",
    "rci": PROJECT_ROOT / "input" / "rci",
    "pdf": PROJECT_ROOT / "input" / "pdf",
    "reference": PROJECT_ROOT / "reference",
}
DOWNLOAD_DIRS = {
    "reports": PROJECT_ROOT / "output" / "reports",
    "anomalies": PROJECT_ROOT / "output" / "anomalies",
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
    "powerbi": {".csv"},
    "logs": {".log"},
}
DOWNLOAD_METADATA = {
    "reports": {"label": "Rapports Excel", "icon": "bi-file-earmark-spreadsheet"},
    "anomalies": {"label": "Audits", "icon": "bi-clipboard-data"},
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


app = Flask(__name__)
app.secret_key = os.getenv("AUTORCI_FLASK_SECRET", "autorcicontrol-local-dashboard")


@app.context_processor
def inject_globals() -> dict[str, Any]:
    return {
        "status_badge_class": status_badge_class,
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
        "ignore_pdf": form.get("ignore_pdf") == "on",
        "no_date_filter": form.get("no_date_filter") == "on",
        "debug_reference": form.get("debug_reference") == "on",
        "debug_invoice": debug_invoice,
    }


def run_pipeline(form_data: dict[str, Any]) -> dict[str, Any]:
    command = [sys.executable, str(PROJECT_ROOT / "src" / "main.py")]
    if form_data.get("dry_run"):
        command.append("--dry-run")
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
    ok = counts.get("OK", 0)
    return {
        "reconciled_invoices": len(run_rows),
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
    }
    if matching_rate is None:
        rate_kpi["display"] = "N/A"

    return [
        {"label": "Factures analysées", "value": _summary_value(summary, "reconciled_invoices"), "kind": "number", "icon": "bi-receipt"},
        {"label": "Factures dans le périmètre RCI", "value": _summary_value(summary, "erp_matchable_invoices"), "kind": "number", "icon": "bi-bullseye"},
        {"label": "Factures OK", "value": _summary_value(summary, "matched_invoices"), "kind": "number", "status": "OK", "icon": "bi-check-circle"},
        {"label": "Factures manquantes RCI", "value": _summary_value(summary, "unmatched_erp"), "kind": "number", "status": "MANQUANTE_RCI", "icon": "bi-exclamation-triangle"},
        {"label": "Factures hors périmètre RCI", "value": _summary_value(summary, "out_of_scope_rci"), "kind": "number", "status": "HORS_SCOPE_RCI", "icon": "bi-slash-circle"},
        {"label": "RCI hors période", "value": _summary_value(summary, "rci_out_of_period"), "kind": "number", "status": "RCI_HORS_PERIODE", "icon": "bi-calendar-x"},
        {"label": "Écarts détectés", "value": _summary_value(summary, "gaps_detected"), "kind": "number", "status": "ANOMALIE_MONTANT", "icon": "bi-exclamation-octagon"},
        {"label": "Montant impacté total", "value": _summary_value(summary, "total_impacted_amount"), "kind": "money", "icon": "bi-cash-coin"},
        rate_kpi,
    ]


def dashboard_alerts(summary: dict[str, Any]) -> list[str]:
    alerts = []
    if _summary_bool(summary, "dashboard_no_rci_flux_in_period"):
        alerts.append("Aucun flux RCI dans la période sélectionnée. Les lignes RCI chargées sont hors période.")
    if _summary_bool(summary, "dashboard_period_mismatch"):
        alerts.append("Attention : le fichier RCI chargé ne correspond pas à la période sélectionnée.")
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
            "rci_rows_excluded_by_date",
            "rci_txt_rows_excluded_by_date",
            "status_rci_hors_periode",
            "rci_out_of_period",
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
