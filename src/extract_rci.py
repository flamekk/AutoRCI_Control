from __future__ import annotations

import logging
import re
import unicodedata
from decimal import Decimal, InvalidOperation
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
except ModuleNotFoundError:  # pragma: no cover - useful when imported as src.extract_rci in tests.
    from src.normalize import (
        detect_document_type,
        normalize_date,
        normalize_invoice_number,
        normalize_text,
        parse_french_amount,
    )


LOGGER = logging.getLogger(__name__)

RCI_EXTENSIONS = {".txt", ".csv", ".xlsx", ".xls", ".xlsm"}
TEXT_EXTENSIONS = {".txt"}
CSV_EXTENSIONS = {".csv"}
EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}

STANDARD_COLUMNS = [
    "source_file",
    "invoice_number",
    "document_type",
    "rci_date",
    "dealer_code",
    "amount_rci",
    "raw_line",
]

INVOICE_RE = re.compile(r"\b(?:FVF|VF|AAF)\s*\d+\b", re.IGNORECASE)
DEALER_RE = re.compile(r"\bSAP\s*0*(\d{5,12})\b", re.IGNORECASE)
DATE_TOKEN_RE = re.compile(r"\d{8}")
LONG_NUMBER_RE = re.compile(r"[+-]?\d{10,}")
DECIMAL_NUMBER_RE = re.compile(r"[+-]?\d{1,3}(?:[\s\u00a0\u202f.]?\d{3})*(?:,\d{1,2})-?|[+-]?\d+\.\d{1,2}")

COLUMN_ALIASES = {
    "invoice_number": [
        "No facture",
        "N facture",
        "Numero facture",
        "Num facture",
        "Facture",
        "Invoice",
        "Invoice number",
        "Document",
    ],
    "rci_date": [
        "Date",
        "Date RCI",
        "Date operation",
        "Date comptabilisation",
        "Date facture",
        "Accounting date",
    ],
    "dealer_code": [
        "Dealer",
        "Dealer code",
        "Code dealer",
        "Code concession",
        "Concession",
        "Donneur ordre",
        "Code donneur ordre",
    ],
    "amount_rci": [
        "Montant",
        "Montant RCI",
        "Montant TTC",
        "Total TTC",
        "Amount",
        "Total",
    ],
}


def extract_rci_folder(folder_path: str | Path) -> "pd.DataFrame":
    pandas = _require_pandas()
    folder = Path(folder_path)
    if not folder.exists():
        LOGGER.warning("Dossier RCI inexistant: %s", folder)
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    files = sorted(
        path for path in folder.iterdir() if path.is_file() and path.suffix.lower() in RCI_EXTENSIONS
    )
    return extract_rci_files(files)


def extract_rci_directory(folder_path: str | Path) -> "pd.DataFrame":
    return extract_rci_folder(folder_path)


def extract_rci_files(paths: Iterable[str | Path]) -> "pd.DataFrame":
    pandas = _require_pandas()
    files = _collect_rci_files(paths)
    if not files:
        LOGGER.info("Aucun fichier RCI TXT/CSV/Excel detecte.")
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    frames = []
    total_input_rows = 0

    for path in files:
        try:
            file_frames, file_input_rows = _extract_file(path)
        except Exception:
            LOGGER.exception("Erreur de lecture RCI pour %s", path)
            continue

        frames.extend(file_frames)
        total_input_rows += file_input_rows

    frames = [frame for frame in frames if not frame.empty]
    if not frames:
        LOGGER.warning("Aucune ligne RCI valide extraite depuis %s fichier(s).", len(files))
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    result = pandas.concat([frame.dropna(axis=1, how="all") for frame in frames], ignore_index=True)
    result = result.reindex(columns=STANDARD_COLUMNS)
    LOGGER.info(
        "Extraction RCI terminee: %s ligne(s) lue(s), %s ligne(s) valide(s), %s fichier(s).",
        total_input_rows,
        len(result),
        len(files),
    )
    return result


def _collect_rci_files(paths: Iterable[str | Path]) -> list[Path]:
    files: list[Path] = []
    for raw_path in paths:
        path = Path(raw_path)
        if path.is_dir():
            files.extend(
                child
                for child in path.iterdir()
                if child.is_file() and child.suffix.lower() in RCI_EXTENSIONS
            )
        elif path.is_file() and path.suffix.lower() in RCI_EXTENSIONS:
            files.append(path)

    return sorted(set(files), key=lambda item: (str(item.parent).lower(), item.name.lower()))


def _extract_file(path: Path) -> tuple[list["pd.DataFrame"], int]:
    suffix = path.suffix.lower()
    if suffix in TEXT_EXTENSIONS:
        frame, input_rows = _extract_text_file(path)
        return [frame], input_rows
    if suffix in CSV_EXTENSIONS:
        raw_frame = _read_csv(path)
        standardized = _standardize_table(raw_frame, path.name, "CSV")
        LOGGER.info(
            "RCI %s [CSV]: %s ligne(s) lue(s), %s ligne(s) conservee(s).",
            path.name,
            len(raw_frame),
            len(standardized),
        )
        return [standardized], len(raw_frame)
    if suffix in EXCEL_EXTENSIONS:
        return _extract_excel_file(path)
    LOGGER.warning("Extension RCI ignoree: %s", path)
    return [], 0


def _extract_text_file(path: Path) -> tuple["pd.DataFrame", int]:
    pandas = _require_pandas()
    lines = _read_text_lines(path)
    records = []
    detail_lines = 0

    for line_number, line in enumerate(lines, start=1):
        if not line.startswith("D"):
            continue

        detail_lines += 1
        try:
            record = _parse_detail_line(line, path.name)
        except Exception as exc:
            LOGGER.warning(
                "Ligne RCI ignoree dans %s:%s (%s): %s",
                path.name,
                line_number,
                exc,
                line[:180].rstrip(),
            )
            continue

        if record is None:
            LOGGER.warning(
                "Ligne RCI detail sans facture valide ignoree dans %s:%s: %s",
                path.name,
                line_number,
                line[:180].rstrip(),
            )
            continue

        records.append(record)

    frame = pandas.DataFrame(records).reindex(columns=STANDARD_COLUMNS)
    LOGGER.info(
        "RCI %s [TXT]: %s ligne(s) lue(s), %s detail(s), %s ligne(s) conservee(s).",
        path.name,
        len(lines),
        detail_lines,
        len(frame),
    )
    return frame, len(lines)


def _read_text_lines(path: Path) -> list[str]:
    last_error: UnicodeDecodeError | None = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return path.read_text(encoding=encoding).splitlines()
        except UnicodeDecodeError as exc:
            last_error = exc
            continue

    if last_error is not None:
        raise last_error
    return []


def _parse_detail_line(line: str, source_file: str) -> dict[str, Any] | None:
    invoice_match = INVOICE_RE.search(line)
    if invoice_match is None:
        return None

    invoice_number = normalize_invoice_number(invoice_match.group(0))
    if invoice_number is None:
        return None

    return {
        "source_file": source_file,
        "invoice_number": invoice_number,
        "document_type": detect_document_type(invoice_number),
        "rci_date": _extract_rci_date(line, invoice_match.start()),
        "dealer_code": _extract_dealer_code(line),
        "amount_rci": _extract_amount_near_invoice(line, invoice_match.start()),
        "raw_line": line.rstrip(),
    }


def _extract_rci_date(line: str, invoice_position: int | None = None) -> str | None:
    matches = list(DATE_TOKEN_RE.finditer(line))
    if invoice_position is not None:
        before_invoice = [match for match in matches if match.start() < invoice_position]
        matches = before_invoice or matches

    for match in matches:
        normalized = _normalize_rci_date_token(match.group(0))
        if normalized is not None:
            return normalized
    return None


def _normalize_rci_date_token(token: str) -> str | None:
    if len(token) != 8:
        return None

    if token.startswith(("19", "20")):
        return _plausible_date(normalize_date(f"{token[0:4]}-{token[4:6]}-{token[6:8]}"))

    if not token[4:8].startswith(("19", "20")):
        return None
    return _plausible_date(normalize_date(f"{token[0:2]}/{token[2:4]}/{token[4:8]}"))


def _plausible_date(value: str | None) -> str | None:
    if value is None:
        return None
    year = int(value[:4])
    if 1990 <= year <= 2100:
        return value
    return None


def _extract_dealer_code(line: str) -> str | None:
    match = DEALER_RE.search(line)
    if match:
        return _normalize_dealer_code(match.group(1))

    return None


def _extract_amount_near_invoice(line: str, invoice_position: int) -> float | None:
    amount = _extract_fixed_width_amount(line, invoice_position)
    if amount is not None:
        return amount

    search_start = max(0, invoice_position - 180)
    search_end = min(len(line), invoice_position + 80)
    segment = line[search_start:search_end]
    candidates = []
    for match in DECIMAL_NUMBER_RE.finditer(segment):
        parsed = parse_french_amount(match.group(0))
        if parsed is None:
            continue
        distance = abs((search_start + match.start()) - invoice_position)
        candidates.append((distance, parsed))

    if not candidates:
        return None
    return sorted(candidates, key=lambda item: item[0])[0][1]


def _extract_fixed_width_amount(line: str, invoice_position: int) -> float | None:
    date_boundary = 0
    for match in DATE_TOKEN_RE.finditer(line[:invoice_position]):
        normalized = _normalize_rci_date_token(match.group(0))
        if normalized is not None:
            date_boundary = max(date_boundary, match.end())

    candidates = []
    for match in LONG_NUMBER_RE.finditer(line[:invoice_position]):
        if match.end() <= date_boundary:
            continue

        chunks = _split_fixed_amount_chunks(match.group(0))
        for chunk_index, chunk in enumerate(chunks):
            amount = _fixed_amount_to_float(chunk)
            if amount is None or amount == 0:
                continue
            candidates.append((match.start(), chunk_index, amount))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def _split_fixed_amount_chunks(token: str) -> list[str]:
    sign = ""
    digits = token
    if token[0] in "+-":
        sign = token[0]
        digits = token[1:]

    if len(digits) >= 26 and len(digits) % 13 == 0:
        return [f"{sign}{digits[index:index + 13]}" for index in range(0, len(digits), 13)]
    if len(digits) >= 13:
        return [f"{sign}{digits[:13]}"]
    return []


def _fixed_amount_to_float(token: str) -> float | None:
    try:
        sign = -1 if token.startswith("-") else 1
        digits = token.lstrip("+-")
        amount = Decimal(digits) / Decimal("100")
    except (InvalidOperation, ValueError):
        return None
    return float(amount * sign)


def _extract_excel_file(path: Path) -> tuple[list["pd.DataFrame"], int]:
    pandas = _require_pandas()
    frames = []
    total_rows = 0

    workbook = pandas.ExcelFile(path)
    for sheet_name in workbook.sheet_names:
        raw_frame = pandas.read_excel(workbook, sheet_name=sheet_name, dtype=object)
        total_rows += len(raw_frame)
        standardized = _standardize_table(raw_frame, path.name, str(sheet_name))
        LOGGER.info(
            "RCI %s [%s]: %s ligne(s) lue(s), %s ligne(s) conservee(s).",
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
    raise RuntimeError(f"Impossible de lire le CSV RCI: {path}")


def _standardize_table(raw_frame: "pd.DataFrame", source_file: str, source_sheet: str) -> "pd.DataFrame":
    pandas = _require_pandas()
    cleaned = raw_frame.dropna(how="all").dropna(axis=1, how="all")
    if cleaned.empty:
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    column_map = _build_column_map(cleaned)
    column_map.update(_infer_missing_columns(cleaned, column_map))

    output = pandas.DataFrame(index=cleaned.index)
    output["source_file"] = source_file

    raw_invoice = _series_for(cleaned, column_map.get("invoice_number"))
    output["invoice_number"] = raw_invoice.map(normalize_invoice_number)
    output = output[output["invoice_number"].notna()].copy()
    if output.empty:
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    output["document_type"] = output["invoice_number"].map(detect_document_type)
    output["rci_date"] = _series_for(cleaned, column_map.get("rci_date")).loc[output.index].map(
        _normalize_table_date
    )
    output["dealer_code"] = _series_for(cleaned, column_map.get("dealer_code")).loc[
        output.index
    ].map(_normalize_dealer_code)
    output["amount_rci"] = _series_for(cleaned, column_map.get("amount_rci")).loc[
        output.index
    ].map(parse_french_amount)
    output["raw_line"] = cleaned.loc[output.index].apply(_row_to_raw_line, axis=1)

    return output.reindex(columns=STANDARD_COLUMNS).reset_index(drop=True)


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

    if "rci_date" not in existing_map:
        date_column = _best_scored_column(
            frame,
            used_columns,
            lambda series: _date_score(series),
            minimum_score=1,
        )
        if date_column is not None:
            inferred["rci_date"] = date_column
            used_columns.add(date_column)

    if "amount_rci" not in existing_map:
        amount_column = _best_scored_column(
            frame,
            used_columns,
            lambda series: _amount_score(series),
            minimum_score=1,
        )
        if amount_column is not None:
            inferred["amount_rci"] = amount_column
            used_columns.add(amount_column)

    if "dealer_code" not in existing_map:
        dealer_column = _best_scored_column(
            frame,
            used_columns,
            lambda series: _dealer_score(series),
            minimum_score=1,
        )
        if dealer_column is not None:
            inferred["dealer_code"] = dealer_column

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


def _invoice_score(series: "pd.Series") -> int:
    return sum(1 for value in _sample_values(series) if normalize_invoice_number(value) is not None)


def _date_score(series: "pd.Series") -> int:
    score = 0
    for value in _sample_values(series):
        if normalize_date(value) is not None:
            score += 1
            continue

        text = str(value).strip()
        if _normalize_rci_date_token(text) is not None:
            score += 1
    return score


def _amount_score(series: "pd.Series") -> int:
    return sum(1 for value in _sample_values(series) if parse_french_amount(value) is not None)


def _dealer_score(series: "pd.Series") -> int:
    score = 0
    for value in _sample_values(series):
        normalized = _normalize_dealer_code(value)
        if normalized and re.fullmatch(r"[A-Z0-9]{2,20}", normalized):
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


def _row_to_raw_line(row: Any) -> str:
    values = []
    for value in row.tolist():
        if _is_blank(value):
            continue
        values.append(str(value))
    return " | ".join(values)


def _normalize_dealer_code(value: Any) -> str | None:
    if _is_blank(value):
        return None

    text = normalize_text(value)
    text = text.replace("SAP", "")
    text = re.sub(r"[^A-Z0-9]", "", text)
    if not text:
        return None

    if text.isdigit():
        text = text.lstrip("0") or "0"
    return text


def _normalize_table_date(value: Any) -> str | None:
    normalized = normalize_date(value)
    if normalized is not None:
        return normalized

    text = str(value).strip()
    if DATE_TOKEN_RE.fullmatch(text):
        return _normalize_rci_date_token(text)
    return None


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
            "Le module pandas est requis pour lire les fichiers RCI. "
            "Installez les dependances avec: python -m pip install -r requirements.txt"
        )
    return pd
