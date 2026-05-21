from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Iterable

try:
    import pandas as pd
except ModuleNotFoundError:  # pragma: no cover - exercised only before dependency install.
    pd = None  # type: ignore[assignment]

try:
    import pdfplumber
except ModuleNotFoundError:  # pragma: no cover - exercised only before dependency install.
    pdfplumber = None  # type: ignore[assignment]

try:
    from normalize import (
        detect_document_type,
        normalize_date,
        normalize_invoice_number,
        normalize_text,
        parse_french_amount,
    )
except ModuleNotFoundError:  # pragma: no cover - useful when imported as src.extract_pdf in tests.
    from src.normalize import (
        detect_document_type,
        normalize_date,
        normalize_invoice_number,
        normalize_text,
        parse_french_amount,
    )


LOGGER = logging.getLogger(__name__)

PDF_EXTENSIONS = {".pdf"}

STANDARD_COLUMNS = [
    "source_file",
    "page",
    "dealer_code",
    "invoice_number",
    "document_type",
    "cf_code",
    "pdf_invoice_date",
    "due_date",
    "model",
    "chassis_number",
    "amount_pdf",
    "origin",
    "raw_line",
]

INVOICE_RE = re.compile(r"\b(?:FVF|VF|AAF)\s*\d{4,}\b", re.IGNORECASE)
LINE_PREFIX_RE = re.compile(
    r"^\s*"
    r"(?P<dealer>\S+)\s+"
    r"(?P<invoice>(?:FVF|VF|AAF)\s*\d{4,})\s+"
    r"(?P<cf>\S+)\s+"
    r"(?P<invoice_date>\d{2}/\d{2}/\d{2,4})\s+"
    r"(?P<due_date>\d{2}/\d{2}/\d{2,4})\s+"
    r"(?P<rest>.+?)\s*$",
    re.IGNORECASE,
)
AMOUNT_RE = re.compile(r"(?<!\S)[+-]?\d+(?:[\s\u00a0\u202f.]\d{3})*,\d{1,2}-?(?!\S)")


def extract_pdf_folder(folder_path: str | Path) -> "pd.DataFrame":
    pandas = _require_pandas()
    folder = Path(folder_path)
    if not folder.exists():
        LOGGER.warning("Dossier PDF inexistant: %s", folder)
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    files = sorted(
        path for path in folder.iterdir() if path.is_file() and path.suffix.lower() in PDF_EXTENSIONS
    )
    return extract_pdf_files(files)


def extract_pdf_directory(folder_path: str | Path) -> "pd.DataFrame":
    return extract_pdf_folder(folder_path)


def extract_pdf_files(paths: Iterable[str | Path]) -> "pd.DataFrame":
    pandas = _require_pandas()
    _require_pdfplumber()

    files = _collect_pdf_files(paths)
    if not files:
        LOGGER.info("Aucun fichier PDF detecte.")
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    frames = []
    total_pages = 0

    for path in files:
        try:
            frame, page_count = _extract_pdf_file(path)
        except Exception:
            LOGGER.exception("Erreur de lecture PDF pour %s", path)
            continue

        frames.append(frame)
        total_pages += page_count

    frames = [frame for frame in frames if not frame.empty]
    if not frames:
        LOGGER.warning("Aucune ligne PDF exploitable extraite depuis %s fichier(s).", len(files))
        return pandas.DataFrame(columns=STANDARD_COLUMNS)

    result = pandas.concat([frame.dropna(axis=1, how="all") for frame in frames], ignore_index=True)
    result = result.reindex(columns=STANDARD_COLUMNS)
    LOGGER.info(
        "Extraction PDF terminee: %s page(s) lue(s), %s ligne(s) valide(s), %s fichier(s).",
        total_pages,
        len(result),
        len(files),
    )
    return result


def _collect_pdf_files(paths: Iterable[str | Path]) -> list[Path]:
    files: list[Path] = []
    for raw_path in paths:
        path = Path(raw_path)
        if path.is_dir():
            files.extend(
                child
                for child in path.iterdir()
                if child.is_file() and child.suffix.lower() in PDF_EXTENSIONS
            )
        elif path.is_file() and path.suffix.lower() in PDF_EXTENSIONS:
            files.append(path)

    return sorted(set(files), key=lambda item: (str(item.parent).lower(), item.name.lower()))


def _extract_pdf_file(path: Path) -> tuple["pd.DataFrame", int]:
    pandas = _require_pandas()
    pdfplumber_module = _require_pdfplumber()

    records = []
    with pdfplumber_module.open(path) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            page_records = []
            for raw_line in text.splitlines():
                if not _is_useful_line(raw_line):
                    continue

                record = _parse_pdf_line(raw_line, path.name, page_index)
                if record is None:
                    LOGGER.warning(
                        "Ligne PDF ignoree dans %s page %s: %s",
                        path.name,
                        page_index,
                        raw_line[:220],
                    )
                    continue
                page_records.append(record)

            records.extend(page_records)
            LOGGER.debug(
                "PDF %s page %s: %s ligne(s) exploitable(s).",
                path.name,
                page_index,
                len(page_records),
            )

        page_count = len(pdf.pages)

    frame = pandas.DataFrame(records).reindex(columns=STANDARD_COLUMNS)
    if frame.empty:
        LOGGER.warning("PDF %s: aucune ligne exploitable detectee.", path.name)
    else:
        LOGGER.info(
            "PDF %s: %s page(s), %s ligne(s) extraite(s).",
            path.name,
            page_count,
            len(frame),
        )
    return frame, page_count


def _is_useful_line(line: str) -> bool:
    if not INVOICE_RE.search(line):
        return False

    upper_line = line.upper()
    if upper_line.startswith("TOTAL "):
        return False
    return bool(AMOUNT_RE.search(line))


def _parse_pdf_line(line: str, source_file: str, page: int) -> dict[str, Any] | None:
    normalized_line = " ".join(line.strip().split())
    match = LINE_PREFIX_RE.match(normalized_line)
    if match is None:
        return _parse_pdf_line_fallback(normalized_line, source_file, page)

    invoice_number = normalize_invoice_number(match.group("invoice"))
    if invoice_number is None:
        return None

    rest = match.group("rest")
    amount_match = _last_amount_match(rest)
    if amount_match is None:
        return None

    before_amount = rest[: amount_match.start()].strip()
    after_amount = rest[amount_match.end() :].strip()
    model, chassis_number = _parse_model_and_chassis(before_amount)

    return {
        "source_file": source_file,
        "page": page,
        "dealer_code": _normalize_dealer_code(match.group("dealer")),
        "invoice_number": invoice_number,
        "document_type": detect_document_type(invoice_number),
        "cf_code": normalize_text(match.group("cf")) or None,
        "pdf_invoice_date": normalize_date(match.group("invoice_date")),
        "due_date": normalize_date(match.group("due_date")),
        "model": model,
        "chassis_number": chassis_number,
        "amount_pdf": parse_french_amount(amount_match.group(0)),
        "origin": _normalize_origin(after_amount),
        "raw_line": normalized_line,
    }


def _parse_pdf_line_fallback(line: str, source_file: str, page: int) -> dict[str, Any] | None:
    invoice_match = INVOICE_RE.search(line)
    if invoice_match is None:
        return None

    invoice_number = normalize_invoice_number(invoice_match.group(0))
    if invoice_number is None:
        return None

    amount_match = _last_amount_match(line)
    if amount_match is None:
        return None

    before_invoice = line[: invoice_match.start()].strip().split()
    after_invoice = line[invoice_match.end() : amount_match.start()].strip().split()
    dealer_code = before_invoice[-1] if before_invoice else None
    cf_code = after_invoice[0] if len(after_invoice) >= 1 else None
    invoice_date = after_invoice[1] if len(after_invoice) >= 2 else None
    due_date = after_invoice[2] if len(after_invoice) >= 3 else None
    model, chassis_number = _parse_model_and_chassis(" ".join(after_invoice[3:]))

    return {
        "source_file": source_file,
        "page": page,
        "dealer_code": _normalize_dealer_code(dealer_code),
        "invoice_number": invoice_number,
        "document_type": detect_document_type(invoice_number),
        "cf_code": normalize_text(cf_code) or None,
        "pdf_invoice_date": normalize_date(invoice_date),
        "due_date": normalize_date(due_date),
        "model": model,
        "chassis_number": chassis_number,
        "amount_pdf": parse_french_amount(amount_match.group(0)),
        "origin": _normalize_origin(line[amount_match.end() :]),
        "raw_line": line,
    }


def _last_amount_match(text: str) -> re.Match[str] | None:
    matches = list(AMOUNT_RE.finditer(text))
    if not matches:
        return None
    return matches[-1]


def _parse_model_and_chassis(value: str) -> tuple[str | None, str | None]:
    tokens = value.strip().split()
    if not tokens:
        return None, None

    model = normalize_text(tokens[0]) or None
    chassis_tokens = tokens[1:]
    chassis_number = normalize_text(" ".join(chassis_tokens)) if chassis_tokens else None
    return model, chassis_number or None


def _normalize_dealer_code(value: Any) -> str | None:
    if value is None:
        return None

    text = normalize_text(value)
    if not text:
        return None
    return re.sub(r"[^A-Z0-9]", "", text) or None


def _normalize_origin(value: Any) -> str | None:
    if value is None:
        return None

    text = normalize_text(value)
    if not text:
        return None
    return text


def _require_pandas() -> Any:
    if pd is None:
        raise RuntimeError(
            "Le module pandas est requis pour produire les DataFrames PDF. "
            "Installez les dependances avec: python -m pip install -r requirements.txt"
        )
    return pd


def _require_pdfplumber() -> Any:
    if pdfplumber is None:
        raise RuntimeError(
            "Le module pdfplumber est requis pour lire les PDF. "
            "Installez les dependances avec: python -m pip install -r requirements.txt"
        )
    return pdfplumber
