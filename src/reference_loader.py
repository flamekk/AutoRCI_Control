from __future__ import annotations

import csv
import json
import logging
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

try:
    import pandas as pd
except ModuleNotFoundError:  # pragma: no cover - exercised only before dependency install.
    pd = None  # type: ignore[assignment]


LOGGER = logging.getLogger(__name__)

EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}

CODE_KEYWORDS = [
    "code",
    "code affaire",
    "code client",
    "code concessionnaire",
    "n client",
    "no client",
    "numero",
    "n cnc",
    "cnc",
    "dealer code",
]

NAME_KEYWORDS = [
    "affaire",
    "affaires",
    "concessionnaire",
    "client",
    "nom",
    "nom client",
    "raison sociale",
    "societe",
    "libelle",
    "dealer",
    "dealer name",
]

REFERENCE_DEBUG_COLUMNS = [
    "source_file",
    "sheet",
    "row_number",
    "raw_values",
    "detected_code",
    "detected_name",
    "normalized_code",
    "normalized_name",
]


@dataclass(frozen=True)
class RciCoverageReference:
    loaded: bool
    code_values: set[str]
    name_values: set[str]
    files: list[Path]
    rows: int = 0

    @property
    def empty(self) -> bool:
        return not self.code_values and not self.name_values


@dataclass(frozen=True)
class SheetDetection:
    frame: "pd.DataFrame"
    code_columns: list[Any]
    name_columns: list[Any]
    header_row: int | None
    strategy: str


def load_rci_coverage_reference(reference_path: str | Path) -> RciCoverageReference:
    folder = Path(reference_path)
    files = _reference_files(folder)
    if files is None:
        return RciCoverageReference(False, set(), set(), [], 0)

    code_values: set[str] = set()
    name_values: set[str] = set()
    total_rows = 0

    for path in files:
        try:
            file_codes, file_names, file_rows, _debug_rows = _load_reference_file(path)
        except Exception:
            LOGGER.exception("Erreur de lecture du referentiel RCI %s. Fichier ignore.", path)
            continue
        code_values.update(file_codes)
        name_values.update(file_names)
        total_rows += file_rows

    loaded = bool(code_values or name_values)
    if not loaded:
        LOGGER.warning(
            "Referentiel RCI vide ou sans colonnes reconnues dans %s. Le controle continue sans filtrage de couverture.",
            folder,
        )
        return RciCoverageReference(False, set(), set(), files, total_rows)

    LOGGER.info(
        "Referentiel RCI charge: %s fichier(s), %s ligne(s) utiles, %s code(s), %s nom(s).",
        len(files),
        total_rows,
        len(code_values),
        len(name_values),
    )
    LOGGER.info("Exemples codes RCI normalises: %s", _sample_values(code_values))
    LOGGER.info("Exemples noms RCI normalises: %s", _sample_values(name_values))
    return RciCoverageReference(True, code_values, name_values, files, total_rows)


def inspect_reference_file(
    reference_path: str | Path,
    output_dir: str | Path | None = None,
    run_id: str | None = None,
) -> Path | None:
    """Log workbook diagnostics and optionally write row-level reference debug CSV."""

    folder = Path(reference_path)
    files = _reference_files(folder)
    if files is None:
        return None

    debug_rows: list[dict[str, Any]] = []
    for path in files:
        try:
            workbook = _require_pandas().ExcelFile(path)
        except Exception:
            LOGGER.exception("Inspection referentiel impossible pour %s.", path)
            continue

        LOGGER.info("Inspection referentiel RCI - fichier: %s", path.name)
        LOGGER.info("Inspection referentiel RCI - feuilles: %s", workbook.sheet_names)
        for sheet_name in workbook.sheet_names:
            raw = _read_raw_sheet(workbook, sheet_name)
            LOGGER.info(
                "Inspection referentiel RCI %s [%s]: dimensions brutes=%s x %s.",
                path.name,
                sheet_name,
                raw.shape[0],
                raw.shape[1],
            )
            LOGGER.info(
                "Inspection referentiel RCI %s [%s] - 5 premieres lignes brutes: %s",
                path.name,
                sheet_name,
                _first_rows_as_json(raw),
            )
            detection = _detect_sheet_table(workbook, sheet_name)
            LOGGER.info(
                "Inspection referentiel RCI %s [%s]: strategie=%s, header_row=%s, colonnes=%s.",
                path.name,
                sheet_name,
                detection.strategy,
                detection.header_row,
                [str(column) for column in detection.frame.columns],
            )
            LOGGER.info(
                "Inspection referentiel RCI %s [%s]: colonnes candidates code=%s, nom=%s.",
                path.name,
                sheet_name,
                [str(column) for column in detection.code_columns],
                [str(column) for column in detection.name_columns],
            )
            debug_rows.extend(_debug_rows_for_sheet(path.name, sheet_name, detection))

    if output_dir is None or run_id is None:
        return None

    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    debug_path = output / f"reference_debug_{run_id}.csv"
    with debug_path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=REFERENCE_DEBUG_COLUMNS, delimiter=";")
        writer.writeheader()
        writer.writerows(debug_rows)
    LOGGER.info("Debug referentiel RCI genere: %s (%s ligne(s)).", debug_path, len(debug_rows))
    return debug_path


def enrich_erp_with_rci_coverage(
    erp_records: Any,
    coverage: RciCoverageReference | None,
    *,
    enabled: bool = True,
) -> "pd.DataFrame":
    pandas = _require_pandas()
    erp = _to_dataframe(erp_records)
    if erp.empty:
        erp["is_rci_covered"] = []
        return erp

    enriched = erp.copy()
    if not enabled:
        enriched["is_rci_covered"] = True
        LOGGER.info("Filtrage couverture RCI desactive: toutes les lignes ERP sont considerees couvertes.")
        return enriched

    if coverage is None or not coverage.loaded or coverage.empty:
        enriched["is_rci_covered"] = True
        LOGGER.warning("Referentiel RCI absent ou vide: toutes les lignes ERP restent dans le scope de controle.")
        return enriched

    customer_codes = (
        enriched["customer_code"]
        if "customer_code" in enriched.columns
        else pandas.Series([None] * len(enriched), index=enriched.index)
    )
    customer_names = (
        enriched["customer_name"]
        if "customer_name" in enriched.columns
        else pandas.Series([None] * len(enriched), index=enriched.index)
    )

    enriched["is_rci_covered"] = [
        _is_covered(code, name, coverage)
        for code, name in zip(customer_codes, customer_names)
    ]
    LOGGER.info(
        "Couverture RCI appliquee: %s ligne(s) ERP couvertes, %s hors scope.",
        int(enriched["is_rci_covered"].sum()),
        int((~enriched["is_rci_covered"]).sum()),
    )
    return enriched


def _reference_files(folder: Path) -> list[Path] | None:
    if not folder.exists():
        LOGGER.warning("Referentiel RCI introuvable: %s. Le controle continue sans filtrage de couverture.", folder)
        return None

    files = sorted(path for path in folder.iterdir() if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS)
    if not files:
        LOGGER.warning(
            "Aucun fichier Excel de referentiel RCI trouve dans %s. Le controle continue sans filtrage de couverture.",
            folder,
        )
        return None
    return files


def _load_reference_file(path: Path) -> tuple[set[str], set[str], int, list[dict[str, Any]]]:
    workbook = _require_pandas().ExcelFile(path)
    code_values: set[str] = set()
    name_values: set[str] = set()
    total_rows = 0
    debug_rows: list[dict[str, Any]] = []

    LOGGER.info("Referentiel RCI %s: feuilles detectees=%s.", path.name, workbook.sheet_names)
    for sheet_name in workbook.sheet_names:
        detection = _detect_sheet_table(workbook, sheet_name)
        frame = detection.frame
        useful_rows = _useful_row_count(frame, detection.code_columns, detection.name_columns)
        total_rows += useful_rows

        for column in detection.code_columns:
            code_values.update(_normalized_codes(frame[column]))
        for column in detection.name_columns:
            name_values.update(_normalized_names(frame[column]))

        LOGGER.info(
            "Referentiel RCI %s [%s]: %s ligne(s) utiles, colonnes code=%s, nom=%s, strategie=%s.",
            path.name,
            sheet_name,
            useful_rows,
            [str(column) for column in detection.code_columns],
            [str(column) for column in detection.name_columns],
            detection.strategy,
        )
        debug_rows.extend(_debug_rows_for_sheet(path.name, sheet_name, detection))

    return code_values, name_values, total_rows, debug_rows


def _detect_sheet_table(workbook: Any, sheet_name: str) -> SheetDetection:
    raw = _read_raw_sheet(workbook, sheet_name)
    header_row = _detect_header_row(raw)
    if header_row is not None and header_row > 0:
        shifted_frame = _frame_from_header_row(raw, header_row)
        code_columns, name_columns = _detect_reference_columns(shifted_frame)
        if code_columns or name_columns:
            return SheetDetection(shifted_frame, code_columns, name_columns, header_row, "detected_header_row")

        fallback_shifted = _single_text_column_fallback(shifted_frame)
        if fallback_shifted is not None:
            return SheetDetection(shifted_frame, [], [fallback_shifted], header_row, "single_text_column_shifted")

        code_columns, name_columns = _detect_columns_by_content(shifted_frame)
        if code_columns or name_columns:
            return SheetDetection(shifted_frame, code_columns, name_columns, header_row, "content_detection_shifted")

    default_frame = _read_default_sheet(workbook, sheet_name)
    code_columns, name_columns = _detect_reference_columns(default_frame)
    if code_columns or name_columns:
        return SheetDetection(default_frame, code_columns, name_columns, None, "default_header")

    if header_row is not None:
        shifted_frame = _frame_from_header_row(raw, header_row)
        code_columns, name_columns = _detect_reference_columns(shifted_frame)
        if code_columns or name_columns:
            return SheetDetection(shifted_frame, code_columns, name_columns, header_row, "detected_header_row")

        fallback_shifted = _single_text_column_fallback(shifted_frame)
        if fallback_shifted is not None:
            return SheetDetection(shifted_frame, [], [fallback_shifted], header_row, "single_text_column_shifted")

        code_columns, name_columns = _detect_columns_by_content(shifted_frame)
        if code_columns or name_columns:
            return SheetDetection(shifted_frame, code_columns, name_columns, header_row, "content_detection_shifted")

    raw_data = _frame_from_single_text_raw_column(raw)
    if raw_data is not None:
        return SheetDetection(raw_data, [], [raw_data.columns[0]], None, "single_text_column_raw")

    fallback_default = _single_text_column_fallback(default_frame)
    if fallback_default is not None:
        return SheetDetection(default_frame, [], [fallback_default], None, "single_text_column_default")

    code_columns, name_columns = _detect_columns_by_content(default_frame)
    if code_columns or name_columns:
        return SheetDetection(default_frame, code_columns, name_columns, None, "content_detection_default")

    return SheetDetection(default_frame, [], [], None, "unresolved")


def _read_default_sheet(workbook: Any, sheet_name: str) -> "pd.DataFrame":
    frame = _require_pandas().read_excel(workbook, sheet_name=sheet_name, dtype=object)
    return frame.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)


def _read_raw_sheet(workbook: Any, sheet_name: str) -> "pd.DataFrame":
    frame = _require_pandas().read_excel(workbook, sheet_name=sheet_name, dtype=object, header=None)
    return frame.dropna(how="all").dropna(axis=1, how="all")


def _detect_reference_columns(frame: "pd.DataFrame") -> tuple[list[Any], list[Any]]:
    if frame.empty:
        return [], []

    canonical_columns = {column: _canonicalize(column) for column in frame.columns}
    code_keywords = {_canonicalize(keyword) for keyword in CODE_KEYWORDS}
    name_keywords = {_canonicalize(keyword) for keyword in NAME_KEYWORDS}

    code_columns = [
        column
        for column, canonical in canonical_columns.items()
        if canonical and any(keyword in canonical or canonical in keyword for keyword in code_keywords)
    ]
    name_columns = [
        column
        for column, canonical in canonical_columns.items()
        if column not in code_columns
        and canonical
        and any(keyword in canonical or canonical in keyword for keyword in name_keywords)
    ]

    return code_columns, name_columns


def _detect_columns_by_content(frame: "pd.DataFrame") -> tuple[list[Any], list[Any]]:
    code_candidates: list[Any] = []
    name_candidates: list[Any] = []
    for column in frame.columns:
        values = _non_empty_values(frame[column])
        if not values:
            continue
        text_count = sum(1 for value in values if _looks_like_name(value))
        code_count = sum(1 for value in values if _looks_like_code(value))
        if code_count >= max(2, len(values) // 3):
            code_candidates.append(column)
        if text_count >= max(2, len(values) // 2):
            name_candidates.append(column)

    name_candidates = [column for column in name_candidates if column not in code_candidates]
    return code_candidates, name_candidates


def _single_text_column_fallback(frame: "pd.DataFrame") -> Any | None:
    if frame.empty:
        return None

    non_empty_columns = [
        column for column in frame.columns if _non_empty_values(frame[column])
    ]
    if len(non_empty_columns) != 1:
        return None

    column = non_empty_columns[0]
    values = _non_empty_values(frame[column])
    if sum(1 for value in values if _looks_like_name(value)) >= max(1, len(values) // 2):
        LOGGER.info("Fallback referentiel RCI: colonne unique texte utilisee comme nom (%s).", column)
        return column
    return None


def _frame_from_single_text_raw_column(raw: "pd.DataFrame") -> "pd.DataFrame | None":
    if raw.empty:
        return None
    fallback_column = _single_text_column_fallback(raw)
    if fallback_column is None:
        return None
    values = [value for value in raw[fallback_column] if normalize_reference_name(value)]
    if not values:
        return None
    return _require_pandas().DataFrame({"reference_name": values})


def _detect_header_row(raw: "pd.DataFrame") -> int | None:
    best_row = None
    best_score = 0
    for row_index, row in raw.iterrows():
        values = [value for value in row.tolist() if normalize_reference_name(value)]
        if not values:
            continue
        if len(values) == 1 and _has_later_multi_cell_row(raw, int(row_index)):
            continue
        keyword_score = sum(_header_keyword_score(value) for value in values)
        score = keyword_score * 10 + len(values)
        if score > best_score and keyword_score > 0:
            best_score = score
            best_row = int(row_index)
    return best_row


def _has_later_multi_cell_row(raw: "pd.DataFrame", row_index: int) -> bool:
    for later_index, row in raw.iterrows():
        if int(later_index) <= row_index:
            continue
        values = [value for value in row.tolist() if normalize_reference_name(value)]
        if len(values) > 1:
            return True
    return False


def _frame_from_header_row(raw: "pd.DataFrame", header_row: int) -> "pd.DataFrame":
    pandas = _require_pandas()
    header_values = raw.loc[header_row].tolist()
    columns = [_header_value(value, index) for index, value in enumerate(header_values)]
    data = raw.loc[raw.index > header_row].copy()
    data.columns = columns
    data = data.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if data.empty:
        return pandas.DataFrame(columns=columns)
    return data


def _debug_rows_for_sheet(source_file: str, sheet_name: str, detection: SheetDetection) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    frame = detection.frame
    code_column = detection.code_columns[0] if detection.code_columns else None
    name_column = detection.name_columns[0] if detection.name_columns else None
    row_offset = (detection.header_row + 2) if detection.header_row is not None else 2

    for index, row in frame.iterrows():
        detected_code = row.get(code_column) if code_column is not None else None
        detected_name = row.get(name_column) if name_column is not None else None
        normalized_code = normalize_reference_code(detected_code)
        normalized_name = normalize_reference_name(detected_name)
        if not normalized_code and not normalized_name:
            continue
        rows.append(
            {
                "source_file": source_file,
                "sheet": sheet_name,
                "row_number": row_offset + int(index),
                "raw_values": json.dumps([_clean_cell(value) for value in row.tolist()], ensure_ascii=False),
                "detected_code": _clean_cell(detected_code),
                "detected_name": _clean_cell(detected_name),
                "normalized_code": normalized_code,
                "normalized_name": normalized_name,
            }
        )
    return rows


def _useful_row_count(frame: "pd.DataFrame", code_columns: list[Any], name_columns: list[Any]) -> int:
    count = 0
    for _, row in frame.iterrows():
        if any(normalize_reference_code(row.get(column)) for column in code_columns):
            count += 1
            continue
        if any(normalize_reference_name(row.get(column)) for column in name_columns):
            count += 1
    return count


def _is_covered(code: Any, name: Any, coverage: RciCoverageReference) -> bool:
    normalized_code = normalize_reference_code(code)
    if normalized_code and normalized_code in coverage.code_values:
        return True

    normalized_name = normalize_reference_name(name)
    return bool(normalized_name and normalized_name in coverage.name_values)


def normalize_reference_code(value: Any) -> str:
    text = normalize_reference_name(value)
    text = re.sub(r"[^A-Z0-9]", "", text)
    if text.isdigit():
        text = text.lstrip("0") or "0"
    return text


def normalize_reference_name(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd is not None and pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(character for character in text if not unicodedata.combining(character))
    text = re.sub(r"\s+", " ", text).strip().upper()
    return text


def _normalized_codes(values: Iterable[Any]) -> set[str]:
    return {normalized for value in values if (normalized := normalize_reference_code(value))}


def _normalized_names(values: Iterable[Any]) -> set[str]:
    return {normalized for value in values if (normalized := normalize_reference_name(value))}


def _non_empty_values(values: Iterable[Any]) -> list[Any]:
    return [value for value in values if normalize_reference_name(value)]


def _looks_like_name(value: Any) -> bool:
    text = normalize_reference_name(value)
    if len(text) < 2:
        return False
    return bool(re.search(r"[A-Z]", text))


def _looks_like_code(value: Any) -> bool:
    text = normalize_reference_code(value)
    if not text:
        return False
    name = normalize_reference_name(value)
    if re.search(r"\s", name):
        return False
    if any(character.isdigit() for character in text):
        return bool(re.fullmatch(r"[A-Z0-9]{2,20}", text))
    return bool(re.fullmatch(r"[A-Z]{2,5}", text))


def _header_keyword_score(value: Any) -> int:
    canonical = _canonicalize(value)
    if not canonical:
        return 0
    keywords = [_canonicalize(keyword) for keyword in [*CODE_KEYWORDS, *NAME_KEYWORDS]]
    return sum(1 for keyword in keywords if keyword and (keyword in canonical or canonical in keyword))


def _header_value(value: Any, index: int) -> str:
    normalized = normalize_reference_name(value)
    return normalized if normalized else f"column_{index + 1}"


def _canonicalize(value: Any) -> str:
    text = normalize_reference_name(value).lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def _first_rows_as_json(frame: "pd.DataFrame", limit: int = 5) -> str:
    rows = []
    for _, row in frame.head(limit).iterrows():
        rows.append([_clean_cell(value) for value in row.tolist()])
    return json.dumps(rows, ensure_ascii=False)


def _clean_cell(value: Any) -> Any:
    try:
        if pd is not None and pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return value


def _sample_values(values: set[str], limit: int = 8) -> str:
    sample = sorted(values)[:limit]
    return ", ".join(sample) if sample else "(aucun)"


def _to_dataframe(records: Any) -> "pd.DataFrame":
    pandas = _require_pandas()
    if records is None:
        return pandas.DataFrame()
    if hasattr(records, "copy") and hasattr(records, "columns"):
        return records.copy()
    if isinstance(records, list):
        return pandas.DataFrame(records)
    return pandas.DataFrame(list(records or []))


def _require_pandas() -> Any:
    if pd is None:
        raise RuntimeError(
            "Le module pandas est requis pour lire le referentiel RCI. "
            "Installez les dependances avec: python -m pip install -r requirements.txt"
        )
    return pd
