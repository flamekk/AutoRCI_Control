from __future__ import annotations

import math
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ModuleNotFoundError:  # pragma: no cover - exercised only before dependency install.
    Workbook = None  # type: ignore[assignment]
    Alignment = Border = Font = PatternFill = Side = None  # type: ignore[assignment]
    Table = TableStyleInfo = None  # type: ignore[assignment]
    get_column_letter = None  # type: ignore[assignment]


RECONCILIATION_COLUMNS = [
    "invoice_number",
    "document_type",
    "customer_code",
    "customer_name",
    "is_rci_covered",
    "amount_erp",
    "amount_rci",
    "amount_pdf",
    "amount_gap",
    "montant_impacte",
    "erp_date",
    "rci_date",
    "pdf_invoice_date",
    "due_date",
    "origin",
    "status",
    "priority",
    "action_recommandee",
    "source_erp",
    "source_rci",
    "source_pdf",
]

RECONCILIATION_HEADERS = {
    "invoice_number": "N° facture",
    "document_type": "Type document",
    "customer_code": "Code concessionnaire",
    "customer_name": "Concessionnaire",
    "is_rci_covered": "Couvert RCI",
    "amount_erp": "Montant ERP",
    "amount_rci": "Montant RCI",
    "amount_pdf": "Montant PDF",
    "amount_gap": "Ecart montant",
    "montant_impacte": "Montant impacte",
    "erp_date": "Date ERP",
    "rci_date": "Date RCI",
    "pdf_invoice_date": "Date facture PDF",
    "due_date": "Date échéance",
    "origin": "Origine",
    "status": "Statut",
    "priority": "Priorité",
    "action_recommandee": "Action recommandée",
    "source_erp": "Source ERP",
    "source_rci": "Source RCI",
    "source_pdf": "Source PDF",
}

STATUS_FILLS = {
    "OK": "C6EFCE",
    "MANQUANTE_RCI": "FCE4D6",
    "ANOMALIE_MONTANT": "FFC7CE",
    "ANOMALIE_DATE": "FFC7CE",
    "DOUBLON": "E4DFEC",
    "RCI_SEULEMENT": "D9EAF7",
    "HORS_SCOPE_RCI": "E7E6E6",
}

STATUS_FONT_COLORS = {
    "OK": "006100",
    "MANQUANTE_RCI": "9C5700",
    "ANOMALIE_MONTANT": "9C0006",
    "ANOMALIE_DATE": "9C0006",
    "DOUBLON": "5F497A",
    "RCI_SEULEMENT": "1F4E78",
    "HORS_SCOPE_RCI": "666666",
}

MONEY_COLUMNS = {"amount_erp", "amount_rci", "amount_pdf", "amount_gap", "montant_impacte"}
SUMMARY_MONEY_LABELS = {
    "Montant total contrôlé",
    "Montant impacté total",
    "Montant manquant RCI",
}
TABLE_STYLE = "TableStyleMedium2"
MAX_TABLE_ROWS = 5000
MAX_AUTOFIT_ROWS = 1000


def write_excel_report(report: dict[str, Any], output_dir: Path, run_id: str) -> Path:
    workbook_class = _require_openpyxl()
    output_dir.mkdir(parents=True, exist_ok=True)
    generated_at = _report_datetime(report, run_id)
    workbook_path = output_dir / f"Rapport_Reconciliation_RCI_{generated_at:%Y-%m-%d_%H%M}.xlsx"

    workbook = workbook_class()
    workbook.remove(workbook.active)

    reconciliation_rows = list(report.get("reconciliation", []))

    _write_rows_sheet(workbook, "Synthèse", _summary_rows(report, generated_at), table_name="TableSynthese")
    _write_reconciliation_sheet(workbook, "Détail rapprochement", reconciliation_rows, "TableDetail")
    _write_reconciliation_sheet(
        workbook,
        "Factures manquantes RCI",
        _filter_status(reconciliation_rows, {"MANQUANTE_RCI"}),
        "TableManquantesRCI",
    )
    _write_reconciliation_sheet(
        workbook,
        "Anomalies",
        _filter_status(reconciliation_rows, {"ANOMALIE_MONTANT", "ANOMALIE_DATE"}),
        "TableAnomalies",
    )
    _write_reconciliation_sheet(
        workbook,
        "Doublons",
        _filter_status(reconciliation_rows, {"DOUBLON"}),
        "TableDoublons",
    )
    _write_reconciliation_sheet(
        workbook,
        "RCI seulement",
        _filter_status(reconciliation_rows, {"RCI_SEULEMENT"}),
        "TableRCISeulement",
    )
    _write_reconciliation_sheet(
        workbook,
        "Hors scope RCI",
        _filter_status(reconciliation_rows, {"HORS_SCOPE_RCI"}),
        "TableHorsScopeRCI",
    )
    _write_rows_sheet(
        workbook,
        "Synthèse par concessionnaire",
        _dealer_summary_rows(reconciliation_rows),
        table_name="TableSyntheseConcessionnaire",
    )

    workbook.save(workbook_path)
    return workbook_path


def _summary_rows(report: dict[str, Any], generated_at: datetime) -> list[list[Any]]:
    summary = report.get("summary", {})
    reconciliation = report.get("reconciliation", [])
    status_counts = _status_counts(reconciliation)
    erp_total = int(summary.get("erp_rows", 0) or 0)
    rci_pdf_total = int(summary.get("rci_rows", 0) or 0) + int(summary.get("pdf_rows", 0) or 0)
    ok_count = status_counts.get("OK", int(summary.get("matched_invoices", 0) or 0))
    missing_count = status_counts.get("MANQUANTE_RCI", int(summary.get("unmatched_erp", 0) or 0))
    amount_anomalies = status_counts.get("ANOMALIE_MONTANT", int(summary.get("amount_anomalies", 0) or 0))
    date_anomalies = status_counts.get("ANOMALIE_DATE", int(summary.get("date_anomalies", 0) or 0))
    duplicates = status_counts.get("DOUBLON", int(summary.get("duplicates", 0) or 0))
    rci_only = status_counts.get("RCI_SEULEMENT", int(summary.get("unmatched_rci", 0) or 0))
    out_of_scope = status_counts.get("HORS_SCOPE_RCI", int(summary.get("out_of_scope_rci", 0) or 0))
    analyzed_total = int(summary.get("reconciled_invoices", len(reconciliation)) or 0)
    total_controlled = float(summary.get("total_controlled_amount", 0) or _sum_amounts(reconciliation, "amount_erp"))
    total_impacted = float(summary.get("total_impacted_amount", 0) or _sum_abs_amounts(reconciliation, "montant_impacte"))
    missing_rci_amount = float(summary.get("missing_rci_amount", 0) or 0)
    erp_matchable = int(summary.get("erp_matchable_invoices", 0) or (
        ok_count + missing_count + amount_anomalies + date_anomalies + duplicates
    ))
    gaps_detected = int(summary.get("gaps_detected", 0) or (
        missing_count + amount_anomalies + date_anomalies + duplicates + rci_only
    ))
    matching_rate = float(summary.get("matching_rate", 0) or (ok_count / erp_matchable if erp_matchable else 0))

    return [
        ["Indicateur", "Valeur"],
        ["Date de traitement", generated_at.strftime("%Y-%m-%d %H:%M")],
        ["Période de rapprochement", summary.get("reconciliation_period", "filtre desactive")],
        ["ERP avant filtre", int(summary.get("erp_rows_before_date_filter", erp_total) or 0)],
        ["ERP après filtre", int(summary.get("erp_rows_after_date_filter", erp_total) or 0)],
        ["Lignes ERP exclues par date", int(summary.get("erp_rows_excluded_by_date", 0) or 0)],
        ["Factures analysées", analyzed_total],
        ["Factures dans le périmètre RCI", erp_matchable],
        ["Factures hors périmètre RCI", out_of_scope],
        ["Factures OK", ok_count],
        ["Factures manquantes RCI", missing_count],
        ["Écarts détectés", gaps_detected],
        ["Montant impacté total", total_impacted],
        ["Nombre total RCI/PDF", rci_pdf_total],
        ["Nombre ANOMALIE_MONTANT", amount_anomalies],
        ["Nombre ANOMALIE_DATE", date_anomalies],
        ["Nombre DOUBLON", duplicates],
        ["Nombre RCI_SEULEMENT", rci_only],
        ["Nombre factures couvertes RCI", int(summary.get("rci_covered_invoices", 0) or 0)],
        ["Montant total contrôlé", total_controlled],
        ["Montant manquant RCI", missing_rci_amount],
        ["Taux de rapprochement", matching_rate],
    ]


def _write_reconciliation_sheet(
    workbook: Any, sheet_name: str, records: list[dict[str, Any]], table_name: str
) -> None:
    rows = [[RECONCILIATION_HEADERS[column] for column in RECONCILIATION_COLUMNS]]
    for record in records:
        rows.append([_clean_cell_value(record.get(column)) for column in RECONCILIATION_COLUMNS])

    worksheet = _write_rows_sheet(workbook, sheet_name, rows, table_name=table_name)
    status_column = RECONCILIATION_COLUMNS.index("status") + 1
    for row_index in range(2, worksheet.max_row + 1):
        _apply_status_style(worksheet, row_index, status_column)
    _format_money_columns(worksheet, RECONCILIATION_COLUMNS)


def _dealer_summary_rows(records: list[dict[str, Any]]) -> list[list[Any]]:
    headers = [
        "Code concessionnaire",
        "Concessionnaire",
        "Total factures",
        "OK",
        "MANQUANTE_RCI",
        "ANOMALIE_MONTANT",
        "ANOMALIE_DATE",
        "DOUBLON",
        "RCI_SEULEMENT",
        "HORS_SCOPE_RCI",
        "Montant ERP",
        "Montant RCI",
        "Montant PDF",
        "Montant écarts",
        "Montant impacte",
    ]
    groups: dict[tuple[str, str], dict[str, Any]] = {}
    for record in records:
        code = str(record.get("customer_code") or "NON RENSEIGNE")
        name = str(record.get("customer_name") or "NON RENSEIGNE")
        key = (code, name)
        if key not in groups:
            groups[key] = {
                "total": 0,
                "OK": 0,
                "MANQUANTE_RCI": 0,
                "ANOMALIE_MONTANT": 0,
                "ANOMALIE_DATE": 0,
                "DOUBLON": 0,
                "RCI_SEULEMENT": 0,
                "HORS_SCOPE_RCI": 0,
                "amount_erp": 0.0,
                "amount_rci": 0.0,
                "amount_pdf": 0.0,
                "amount_gap": 0.0,
                "montant_impacte": 0.0,
            }
        group = groups[key]
        status = str(record.get("status") or "")
        group["total"] += 1
        if status in STATUS_FILLS:
            group[status] += 1
        group["amount_erp"] += abs(_number(record.get("amount_erp")) or 0)
        group["amount_rci"] += abs(_number(record.get("amount_rci")) or 0)
        group["amount_pdf"] += abs(_number(record.get("amount_pdf")) or 0)
        group["amount_gap"] += abs(_number(record.get("amount_gap")) or 0)
        group["montant_impacte"] += abs(_number(record.get("montant_impacte")) or 0)

    rows = [headers]
    for (code, name), values in sorted(groups.items(), key=lambda item: item[0]):
        rows.append(
            [
                code,
                name,
                values["total"],
                values["OK"],
                values["MANQUANTE_RCI"],
                values["ANOMALIE_MONTANT"],
                values["ANOMALIE_DATE"],
                values["DOUBLON"],
                values["RCI_SEULEMENT"],
                values["HORS_SCOPE_RCI"],
                round(values["amount_erp"], 2),
                round(values["amount_rci"], 2),
                round(values["amount_pdf"], 2),
                round(values["amount_gap"], 2),
                round(values["montant_impacte"], 2),
            ]
        )
    return rows


def _write_rows_sheet(
    workbook: Any, sheet_name: str, rows: list[list[Any]], table_name: str
) -> Any:
    worksheet = workbook.create_sheet(title=sheet_name)
    for row in rows:
        worksheet.append([_clean_cell_value(value) for value in row])

    _style_sheet(worksheet, table_name)
    return worksheet


def _style_sheet(worksheet: Any, table_name: str) -> None:
    worksheet.freeze_panes = "A2"
    if worksheet.max_row == 0 or worksheet.max_column == 0:
        return

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    thin_border = Border(bottom=Side(style="thin", color="D9E2F3"))

    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=False)

    ref = f"A1:{get_column_letter(worksheet.max_column)}{max(worksheet.max_row, 1)}"
    worksheet.auto_filter.ref = ref
    if worksheet.max_row <= MAX_TABLE_ROWS:
        table = Table(displayName=_safe_table_name(table_name), ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name=TABLE_STYLE,
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        worksheet.add_table(table)

    _format_summary_sheet(worksheet)
    _auto_fit_columns(worksheet)


def _format_summary_sheet(worksheet: Any) -> None:
    if worksheet.title != "Synthèse":
        if worksheet.title == "Synthèse par concessionnaire":
            for column_index in range(11, 16):
                for cell in worksheet.iter_cols(
                    min_col=column_index,
                    max_col=column_index,
                    min_row=2,
                    max_row=worksheet.max_row,
                ):
                    for item in cell:
                        item.number_format = '#,##0.00'
        return

    for row_index in range(2, worksheet.max_row + 1):
        label = worksheet.cell(row=row_index, column=1).value
        value_cell = worksheet.cell(row=row_index, column=2)
        if label in SUMMARY_MONEY_LABELS:
            value_cell.number_format = '#,##0.00'
        elif label == "Taux de rapprochement":
            value_cell.number_format = '0.00%'
        elif isinstance(value_cell.value, int):
            value_cell.number_format = '#,##0'
        worksheet.cell(row=row_index, column=1).font = Font(bold=True)


def _format_money_columns(worksheet: Any, columns: list[str]) -> None:
    for index, column in enumerate(columns, start=1):
        if column not in MONEY_COLUMNS:
            continue
        for row in range(2, worksheet.max_row + 1):
            worksheet.cell(row=row, column=index).number_format = '#,##0.00'


def _apply_status_style(worksheet: Any, row_index: int, status_column: int) -> None:
    status = worksheet.cell(row=row_index, column=status_column).value
    if not status:
        return
    fill_color = STATUS_FILLS.get(str(status))
    font_color = STATUS_FONT_COLORS.get(str(status), "000000")
    if fill_color is None:
        return

    status_cell = worksheet.cell(row=row_index, column=status_column)
    status_cell.fill = PatternFill("solid", fgColor=fill_color)
    status_cell.font = Font(color=font_color, bold=True)


def _auto_fit_columns(worksheet: Any) -> None:
    max_row = min(worksheet.max_row, MAX_AUTOFIT_ROWS)
    for column_index in range(1, worksheet.max_column + 1):
        column_letter = get_column_letter(column_index)
        max_length = 0
        for row_index in range(1, max_row + 1):
            cell = worksheet.cell(row=row_index, column=column_index)
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 55)


def _filter_status(records: list[dict[str, Any]], statuses: set[str]) -> list[dict[str, Any]]:
    return [record for record in records if record.get("status") in statuses]


def _status_counts(records: list[dict[str, Any]]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for record in records:
        status = str(record.get("status") or "")
        if not status:
            continue
        counts[status] = counts.get(status, 0) + 1
    return counts


def _sum_amounts(records: list[dict[str, Any]], field: str) -> float:
    total = sum(abs(_number(record.get(field)) or 0) for record in records)
    return round(total, 2)


def _sum_abs_amounts(records: list[dict[str, Any]], field: str) -> float:
    total = sum(abs(_number(record.get(field)) or 0) for record in records)
    return round(total, 2)


def _number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return number


def _report_datetime(report: dict[str, Any], run_id: str) -> datetime:
    for date_format in ("%Y%m%d_%H%M%S", "%Y%m%d_%H%M"):
        try:
            return datetime.strptime(run_id, date_format)
        except ValueError:
            continue

    generated_at = report.get("generated_at")
    if generated_at:
        try:
            return datetime.fromisoformat(str(generated_at).replace("Z", "+00:00")).replace(tzinfo=None)
        except ValueError:
            pass
    return datetime.now()


def _clean_cell_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    return value


def _safe_table_name(value: str) -> str:
    cleaned = "".join(character for character in value if character.isalnum())
    if not cleaned:
        return "TableRapport"
    if cleaned[0].isdigit():
        return f"Table{cleaned}"
    return cleaned[:254]


def _require_openpyxl() -> Any:
    if Workbook is None:
        raise RuntimeError(
            "Le module openpyxl est requis pour generer le rapport Excel. "
            "Installez les dependances avec: python -m pip install -r requirements.txt"
        )
    return Workbook
