from __future__ import annotations

import logging
import re
import unicodedata
from pathlib import Path
from typing import Any, Iterable

try:
    import pandas as pd
except ModuleNotFoundError:  # pragma: no cover - exercised only before dependency install.
    pd = None  # type: ignore[assignment]

try:
    from normalize import (
        detect_document_type,
        normalize_date,
        normalize_invoice_number,
        normalize_text,
        parse_french_amount,
    )
except ModuleNotFoundError:  # pragma: no cover - useful when imported as src.extract_erp in tests.
    from src.normalize import (
        detect_document_type,
        normalize_date,
        normalize_invoice_number,
        normalize_text,
        parse_french_amount,
    )


LOGGER = logging.getLogger(__name__)

ERP_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv"}
EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
CSV_EXTENSIONS = {".csv"}

STANDARD_COLUMNS = [
    "source_file",
    "source_sheet",
    "invoice_number",
    "document_type",
    "erp_date",
    "customer_code",
    "customer_name",
    "amount_erp",
    "department_code",
    "brand_code",
    "sales_order_type",
]

COLUMN_ALIASES = {
    "invoice_number": [
        "No facture",
        "N facture",
        "Numero facture",
        "Num facture",
        "Facture",
        "Numero",
        "No",
        "N",
    ],
    "erp_date": [
        "Date comptabilisation",
        "Date facture",
        "Date de comptabilisation",
        "Date document",
        "Date",
    ],
    "amount_erp": [
        "Montant TTC",
        "Total TTC",
        "Montant",
        "Total",
    ],
    "customer_code": [
        "No donneur d'ordre",
        "Numero donneur d'ordre",
        "Code donneur d'ordre",
        "No concession",
        "Numero concession",
        "Numero CNC",
        "No CNC",
        "CNC",
        "Numero client facture",
        "No client facture",
    ],
    "customer_name": [
        "Nom du donneur d'ordre",
        "Nom donneur d'ordre",
        "Nom",
        "Client",
        "Concessionnaire",
        "Concession",
    ],
    "department_code": [
        "Code departement",
        "Departement",
        "CTG",
    ],
    "brand_code": [
        "Code marque",
        "Marque",
        "MQ",
    ],
    "sales_order_type": [
        "Type Commande Vente",
        "Type commande",
        "Commande vente",
        "Type vente",
    ],
}

def extract_erp_folder(folder_path: str | Path) -> "pd.DataFrame":
    pandas = _require_pandas()
    folder = Path(folder_path)
    if not folder.exists():
        LOGGER.warning("Dossier ERP inexistant: %s", folder)
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    files = sorted(
        path for path in folder.iterdir() if path.is_file() and path.suffix.lower() in ERP_EXTENSIONS
    )
    return extract_erp_files(files)


def extract_erp_directory(folder_path: str | Path) -> "pd.DataFrame":
    return extract_erp_folder(folder_path)


def extract_erp_files(paths: Iterable[str | Path]) -> "pd.DataFrame":
    pandas = _require_pandas()
    files = _collect_erp_files(paths)
    if not files:
        LOGGER.info("Aucun fichier ERP Excel/CSV detecte.")
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    frames = []
    total_input_rows = 0

    for path in files:
        try:
            file_frames, file_input_rows = _extract_file(path)
        except Exception:
            LOGGER.exception("Erreur de lecture ERP pour %s", path)
            continue

        frames.extend(file_frames)
        total_input_rows += file_input_rows

    frames = [frame for frame in frames if not frame.empty]
    if not frames:
        LOGGER.warning("Aucune ligne ERP valide extraite depuis %s fichier(s).", len(files))
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    result = pandas.concat([frame.dropna(axis=1, how="all") for frame in frames], ignore_index=True)
    result = result.reindex(columns=STANDARD_COLUMNS)
    LOGGER.info(
        "Extraction ERP terminee: %s ligne(s) lue(s), %s ligne(s) valide(s), %s fichier(s).",
        total_input_rows,
        len(result),
        len(files),
    )
    return result


def _collect_erp_files(paths: Iterable[str | Path]) -> list[Path]:
    files: list[Path] = []
    for raw_path in paths:
        path = Path(raw_path)
        if path.is_dir():
            files.extend(
                child
                for child in path.iterdir()
                if child.is_file() and child.suffix.lower() in ERP_EXTENSIONS
            )
        elif path.is_file() and path.suffix.lower() in ERP_EXTENSIONS:
            files.append(path)

    return sorted(set(files), key=lambda item: (str(item.parent).lower(), item.name.lower()))


def _extract_file(path: Path) -> tuple[list["pd.DataFrame"], int]:
    suffix = path.suffix.lower()
    if suffix in EXCEL_EXTENSIONS:
        return _extract_excel_file(path)
    if suffix in CSV_EXTENSIONS:
        raw_frame = _read_csv(path)
        standardized = _standardize_sheet(raw_frame, path, "CSV")
        LOGGER.info(
            "ERP %s [CSV]: %s ligne(s) lue(s), %s ligne(s) conservee(s).",
            path.name,
            len(raw_frame),
            len(standardized),
        )
        return [standardized], len(raw_frame)
    LOGGER.warning("Extension ERP ignoree: %s", path)
    return [], 0


def _extract_excel_file(path: Path) -> tuple[list["pd.DataFrame"], int]:
    pandas = _require_pandas()
    frames = []
    total_rows = 0

    workbook = pandas.ExcelFile(path)
    for sheet_name in workbook.sheet_names:
        raw_frame = pandas.read_excel(workbook, sheet_name=sheet_name, dtype=object)
        total_rows += len(raw_frame)
        standardized = _standardize_sheet(raw_frame, path, sheet_name)
        LOGGER.info(
            "ERP %s [%s]: %s ligne(s) lue(s), %s ligne(s) conservee(s).",
            path.name,
            sheet_name,
            len(raw_frame),
            len(standardized),
        )
        frames.append(standardized)

    return frames, total_rows


def _read_csv(path: Path) -> "pd.DataFrame":
    pandas = _require_pandas()
    last_error: Exception | None = None

    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return pandas.read_csv(
                path,
                sep=None,
                engine="python",
                dtype=object,
                encoding=encoding,
            )
        except UnicodeDecodeError as exc:
            last_error = exc
            continue

    if last_error is not None:
        raise last_error
    raise RuntimeError(f"Impossible de lire le CSV ERP: {path}")


def _standardize_sheet(raw_frame: "pd.DataFrame", source_file: Path, source_sheet: str) -> "pd.DataFrame":
    pandas = _require_pandas()
    cleaned = _drop_empty_rows_and_columns(raw_frame)
    if cleaned.empty:
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    if _has_only_unnamed_columns(cleaned):
        unnamed_output = _standardize_unnamed_sheet(cleaned, source_file, source_sheet)
        if not unnamed_output.empty:
            return unnamed_output

    column_map = _build_column_map(cleaned)
    column_map.update(_infer_missing_columns(cleaned, column_map))

    output = pandas.DataFrame(index=cleaned.index)
    output["source_file"] = source_file.name
    output["source_sheet"] = source_sheet

    raw_invoice = _series_for(cleaned, column_map.get("invoice_number"))
    output["invoice_number"] = raw_invoice.map(normalize_invoice_number)
    output = output[output["invoice_number"].notna()].copy()

    if output.empty:
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    output["document_type"] = output["invoice_number"].map(detect_document_type)
    output["erp_date"] = _series_for(cleaned, column_map.get("erp_date")).loc[output.index].map(
        normalize_date
    )
    output["customer_code"] = _series_for(
        cleaned, column_map.get("customer_code")
    ).loc[output.index].map(_normalize_optional_text)
    output["customer_name"] = _series_for(
        cleaned, column_map.get("customer_name")
    ).loc[output.index].map(_normalize_optional_text)
    output["amount_erp"] = _series_for(cleaned, column_map.get("amount_erp")).loc[
        output.index
    ].map(parse_french_amount)
    output["department_code"] = _series_for(
        cleaned, column_map.get("department_code")
    ).loc[output.index].map(_normalize_optional_text)
    output["brand_code"] = _series_for(cleaned, column_map.get("brand_code")).loc[
        output.index
    ].map(_normalize_optional_text)
    output["sales_order_type"] = _series_for(
        cleaned, column_map.get("sales_order_type")
    ).loc[output.index].map(_normalize_optional_text)

    return _keep_business_rows(output.reindex(columns=STANDARD_COLUMNS)).reset_index(drop=True)


def _standardize_unnamed_sheet(
    frame: "pd.DataFrame", source_file: Path, source_sheet: str
) -> "pd.DataFrame":
    pandas = _require_pandas()
    rows: list[dict[str, Any]] = []
    columns = list(frame.columns)

    for _, row in frame.iterrows():
        invoice_index = _find_invoice_cell_index(row, columns)
        if invoice_index is None:
            continue

        invoice_number = normalize_invoice_number(row[columns[invoice_index]])
        if invoice_number is None:
            continue

        document_type = detect_document_type(invoice_number)
        rows.append(
            _standardize_unnamed_row(
                row=row,
                columns=columns,
                source_file=source_file.name,
                source_sheet=source_sheet,
                invoice_number=invoice_number,
                document_type=document_type,
            )
        )

    if not rows:
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    return _keep_business_rows(pandas.DataFrame(rows).reindex(columns=STANDARD_COLUMNS))


def _keep_business_rows(frame: "pd.DataFrame") -> "pd.DataFrame":
    if frame.empty:
        return frame
    return frame[frame["erp_date"].notna() | frame["amount_erp"].notna()].copy()


def _standardize_unnamed_row(
    row: Any,
    columns: list[Any],
    source_file: str,
    source_sheet: str,
    invoice_number: str,
    document_type: str,
) -> dict[str, Any]:
    if document_type == "AVOIR":
        return {
            "source_file": source_file,
            "source_sheet": source_sheet,
            "invoice_number": invoice_number,
            "document_type": document_type,
            "erp_date": normalize_date(_value_at(row, columns, 3)),
            "customer_code": _normalize_optional_text(_value_at(row, columns, 5)),
            "customer_name": _normalize_optional_text(_value_at(row, columns, 6)),
            "amount_erp": _first_amount(row, columns, [10, 9]),
            "department_code": None,
            "brand_code": None,
            "sales_order_type": _normalize_optional_text(_value_at(row, columns, 7)),
        }

    return {
        "source_file": source_file,
        "source_sheet": source_sheet,
        "invoice_number": invoice_number,
        "document_type": document_type,
        "erp_date": normalize_date(_value_at(row, columns, 0)),
        "customer_code": _normalize_optional_text(_value_at(row, columns, 3)),
        "customer_name": _normalize_optional_text(_value_at(row, columns, 4)),
        "amount_erp": _first_amount(row, columns, [8]),
        "department_code": _normalize_optional_text(_value_at(row, columns, 9)),
        "brand_code": _normalize_optional_text(_value_at(row, columns, 10)),
        "sales_order_type": _normalize_optional_text(_value_at(row, columns, 12)),
    }


def _drop_empty_rows_and_columns(frame: "pd.DataFrame") -> "pd.DataFrame":
    cleaned = frame.dropna(how="all")
    cleaned = cleaned.dropna(axis=1, how="all")
    return cleaned


def _has_only_unnamed_columns(frame: "pd.DataFrame") -> bool:
    return all(str(column).lower().startswith("unnamed") for column in frame.columns)


def _find_invoice_cell_index(row: Any, columns: list[Any]) -> int | None:
    for index, column in enumerate(columns):
        if normalize_invoice_number(row[column]) is not None:
            return index
    return None


def _value_at(row: Any, columns: list[Any], index: int) -> Any:
    if index >= len(columns):
        return None
    return row[columns[index]]


def _first_amount(row: Any, columns: list[Any], indexes: list[int]) -> float | None:
    for index in indexes:
        amount = parse_french_amount(_value_at(row, columns, index))
        if amount is not None:
            return amount
    return None


def _build_column_map(frame: "pd.DataFrame") -> dict[str, Any]:
    normalized_columns = {
        column: _canonicalize_column_name(str(column))
        for column in frame.columns
    }
    used_columns: set[Any] = set()
    column_map: dict[str, Any] = {}

    for target, aliases in COLUMN_ALIASES.items():
        for alias in (_canonicalize_column_name(alias) for alias in aliases):
            matched_column = _find_exact_column(normalized_columns, alias, used_columns)
            if matched_column is not None:
                column_map[target] = matched_column
                used_columns.add(matched_column)
                break

    return column_map


def _find_exact_column(
    normalized_columns: dict[Any, str], alias: str, used_columns: set[Any]
) -> Any | None:
    for column, normalized_column in normalized_columns.items():
        if column in used_columns:
            continue
        if normalized_column == alias:
            return column
    return None


def _infer_missing_columns(frame: "pd.DataFrame", existing_map: dict[str, Any]) -> dict[str, Any]:
    inferred: dict[str, Any] = {}
    used_columns = set(existing_map.values())
    columns = list(frame.columns)

    if "invoice_number" not in existing_map:
        invoice_column = _best_scored_column(
            frame,
            used_columns,
            lambda series: _invoice_score(series),
            minimum_score=1,
        )
        if invoice_column is not None:
            inferred["invoice_number"] = invoice_column
            used_columns.add(invoice_column)

    if "erp_date" not in existing_map:
        date_column = _best_scored_column(
            frame,
            used_columns,
            lambda series: _date_score(series),
            minimum_score=1,
        )
        if date_column is not None:
            inferred["erp_date"] = date_column
            used_columns.add(date_column)

    if "amount_erp" not in existing_map:
        amount_column = _best_scored_column(
            frame,
            used_columns,
            lambda series: _amount_score(series),
            minimum_score=2,
        )
        if amount_column is not None:
            inferred["amount_erp"] = amount_column
            used_columns.add(amount_column)

    invoice_column = inferred.get("invoice_number") or existing_map.get("invoice_number")
    amount_column = inferred.get("amount_erp") or existing_map.get("amount_erp")

    if invoice_column in columns:
        invoice_index = columns.index(invoice_column)
        if "customer_code" not in existing_map:
            customer_code = _first_column_after(
                frame,
                invoice_index,
                used_columns,
                lambda series: _code_score(series) >= 1,
            )
            if customer_code is not None:
                inferred["customer_code"] = customer_code
                used_columns.add(customer_code)

        if "customer_name" not in existing_map:
            name_start = columns.index(inferred["customer_code"]) if "customer_code" in inferred else invoice_index
            customer_name = _first_column_after(
                frame,
                name_start,
                used_columns,
                lambda series: _name_score(series) >= 1,
            )
            if customer_name is not None:
                inferred["customer_name"] = customer_name
                used_columns.add(customer_name)

    if amount_column in columns:
        amount_index = columns.index(amount_column)
        if "department_code" not in existing_map:
            department_code = _first_column_after(
                frame,
                amount_index,
                used_columns,
                lambda series: _short_code_score(series) >= 1,
            )
            if department_code is not None:
                inferred["department_code"] = department_code
                used_columns.add(department_code)

        if "brand_code" not in existing_map:
            brand_start = (
                columns.index(inferred["department_code"])
                if "department_code" in inferred
                else amount_index
            )
            brand_code = _first_column_after(
                frame,
                brand_start,
                used_columns,
                lambda series: _name_score(series) >= 1 or _short_code_score(series) >= 1,
            )
            if brand_code is not None:
                inferred["brand_code"] = brand_code
                used_columns.add(brand_code)

        if "sales_order_type" not in existing_map:
            order_start = columns.index(inferred["brand_code"]) if "brand_code" in inferred else amount_index
            sales_order_type = _first_column_after(
                frame,
                order_start,
                used_columns,
                lambda series: _name_score(series) >= 1,
            )
            if sales_order_type is not None:
                inferred["sales_order_type"] = sales_order_type

    return inferred


def _best_scored_column(
    frame: "pd.DataFrame",
    used_columns: set[Any],
    scorer: Any,
    minimum_score: int,
) -> Any | None:
    best_column = None
    best_score = minimum_score - 1

    for column in frame.columns:
        if column in used_columns:
            continue
        score = scorer(frame[column])
        if score > best_score:
            best_column = column
            best_score = score

    return best_column


def _first_column_after(
    frame: "pd.DataFrame",
    start_index: int,
    used_columns: set[Any],
    predicate: Any,
) -> Any | None:
    columns = list(frame.columns)
    for column in columns[start_index + 1 :]:
        if column in used_columns:
            continue
        if predicate(frame[column]):
            return column
    return None


def _invoice_score(series: "pd.Series") -> int:
    return sum(1 for value in _sample_values(series) if normalize_invoice_number(value) is not None)


def _date_score(series: "pd.Series") -> int:
    score = 0
    for value in _sample_values(series):
        normalized = normalize_date(value)
        if normalized and re.match(r"20\d{2}-\d{2}-\d{2}$", normalized):
            score += 1
    return score


def _amount_score(series: "pd.Series") -> int:
    score = 0
    for value in _sample_values(series):
        amount = parse_french_amount(value)
        if amount is None:
            continue

        text = str(value)
        score += 1
        if re.search(r"[\d][\s\u00a0\u202f.]*\d+,\d{1,2}-?$", text):
            score += 2
        elif re.search(r"\d+\.\d{1,2}$", text):
            score += 1
        if abs(amount) >= 10_000_000 and not re.search(r"[,.]", text):
            score -= 1
    return score


def _code_score(series: "pd.Series") -> int:
    score = 0
    for value in _sample_values(series):
        text = _normalize_optional_text(value)
        if text and re.fullmatch(r"[A-Z0-9]{2,20}", text):
            score += 1
    return score


def _short_code_score(series: "pd.Series") -> int:
    score = 0
    for value in _sample_values(series):
        text = _normalize_optional_text(value)
        if text and re.fullmatch(r"[A-Z0-9]{1,8}", text):
            score += 1
    return score


def _name_score(series: "pd.Series") -> int:
    score = 0
    for value in _sample_values(series):
        text = _normalize_optional_text(value)
        if text and re.search(r"[A-Z]", text) and len(text) > 2:
            score += 1
    return score


def _sample_values(series: "pd.Series", limit: int = 50) -> list[Any]:
    values = []
    for value in series:
        if _is_blank(value):
            continue
        values.append(value)
        if len(values) >= limit:
            break
    return values


def _series_for(frame: "pd.DataFrame", column: Any | None) -> "pd.Series":
    pandas = _require_pandas()
    if column is None or column not in frame.columns:
        return pandas.Series([None] * len(frame), index=frame.index, dtype=object)
    return frame[column]


def _normalize_optional_text(value: Any) -> str | None:
    if _is_blank(value):
        return None

    text = normalize_text(value)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text or None


def _is_blank(value: Any) -> bool:
    pandas = _require_pandas()
    try:
        if pandas.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ""


def _canonicalize_column_name(value: str) -> str:
    text = unicodedata.normalize("NFKD", value)
    text = "".join(character for character in text if not unicodedata.combining(character))
    text = text.lower()
    text = text.replace("n°", "numero")
    text = text.replace("nº", "numero")
    text = re.sub(r"\bno\.?\b", "numero", text)
    text = re.sub(r"\bnum\.?\b", "numero", text)
    text = re.sub(r"\bn\b", "numero", text)
    return re.sub(r"[^a-z0-9]+", "", text)


def _require_pandas() -> Any:
    if pd is None:
        raise RuntimeError(
            "Le module pandas est requis pour lire les exports ERP. "
            "Installez les dependances avec: python -m pip install -r requirements.txt"
        )
    return pd
