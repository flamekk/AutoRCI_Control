from __future__ import annotations

import argparse
import logging
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    from src.audit import enrich_report_with_audits
    from src.action_plan import build_reference_suggestions, write_reference_suggestions
    from src.config import ensure_project_structure, get_project_root, load_config
    from src.corrective_batch import write_corrective_batch_outputs
    from src.date_filter import apply_reconciliation_date_filter
    from src.debug_matching import log_debug_invoice, log_invoice_presence, write_matching_debug
    from src.email_sender import send_report
    from src.extract_erp import extract_erp_files
    from src.extract_pdf import extract_pdf_files
    from src.extract_rci import extract_rci_files
    from src.file_detector import SOURCE_TYPES, FileInventory, detect_files
    from src.missing_rci import write_missing_rci_export
    from src.powerbi_export import write_powerbi_exports
    from src.reference_loader import (
        enrich_erp_with_rci_coverage,
        inspect_reference_file,
        load_rci_coverage_reference,
    )
    from src.reconcile import reconcile
    from src.report_excel import write_excel_report
except ModuleNotFoundError:  # pragma: no cover - used when running python src/main.py.
    from audit import enrich_report_with_audits
    from action_plan import build_reference_suggestions, write_reference_suggestions
    from config import ensure_project_structure, get_project_root, load_config
    from corrective_batch import write_corrective_batch_outputs
    from date_filter import apply_reconciliation_date_filter
    from debug_matching import log_debug_invoice, log_invoice_presence, write_matching_debug
    from email_sender import send_report
    from extract_erp import extract_erp_files
    from extract_pdf import extract_pdf_files
    from extract_rci import extract_rci_files
    from file_detector import SOURCE_TYPES, FileInventory, detect_files
    from missing_rci import write_missing_rci_export
    from powerbi_export import write_powerbi_exports
    from reference_loader import enrich_erp_with_rci_coverage, inspect_reference_file, load_rci_coverage_reference
    from reconcile import reconcile
    from report_excel import write_excel_report


LOGGER = logging.getLogger("autorcicontrol")


class PipelineError(RuntimeError):
    """Raised for clean, user-actionable pipeline failures."""


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automatisation du rapprochement ERP Navision/Incadea avec RCI Banque."
    )
    parser.add_argument(
        "--use-samples",
        action="store_true",
        help="Utilise samples/ au lieu de input/ pour tester le developpement.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="N'envoie pas d'email et n'archive pas les fichiers source.",
    )
    parser.add_argument(
        "--no-archive",
        action="store_true",
        help="Execute le traitement reel sans archiver les fichiers input apres succes.",
    )
    parser.add_argument(
        "--ignore-pdf",
        action="store_true",
        help="Ignore les PDF et rapproche uniquement ERP avec le flux RCI TXT/CSV/Excel.",
    )
    parser.add_argument(
        "--debug-invoice",
        default=None,
        help="Journalise toutes les informations ERP/RCI/PDF et le statut final pour une facture donnee.",
    )
    parser.add_argument(
        "--debug-reference",
        action="store_true",
        help="Inspecte le referentiel RCI et genere output/anomalies/reference_debug_*.csv.",
    )
    parser.add_argument(
        "--no-date-filter",
        action="store_true",
        help="Desactive le filtre de periode de rapprochement ERP.",
    )
    parser.add_argument(
        "--date-from",
        default=None,
        help="Date de debut forcee pour le rapprochement au format YYYY-MM-DD.",
    )
    parser.add_argument(
        "--date-to",
        default=None,
        help="Date de fin forcee pour le rapprochement au format YYYY-MM-DD.",
    )
    parser.add_argument(
        "--config",
        default=None,
        help="Chemin optionnel vers un fichier de configuration YAML.",
    )
    parser.add_argument(
        "--log-level",
        default=None,
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Niveau de logs console/fichier.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    log_file: Path | None = None

    try:
        project_root = get_project_root()
        ensure_project_structure(project_root)

        config_path = _resolve_config_path(project_root, args.config)
        config = load_config(config_path)

        run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = _setup_logging(project_root, config, run_id, args.log_level)

        LOGGER.info("Demarrage AutoRCI_Control")
        LOGGER.info("Racine projet: %s", project_root)
        LOGGER.info("Configuration: %s", config_path)
        LOGGER.info("Mode source: %s", "samples" if args.use_samples else "input")
        LOGGER.info("Dry-run: %s", "oui" if args.dry_run else "non")
        LOGGER.info("Archivage desactive par option: %s", "oui" if args.no_archive else "non")
        LOGGER.info("Ignore PDF: %s", "oui" if args.ignore_pdf else "non")
        LOGGER.info("Filtre date: %s", "non" if args.no_date_filter else "oui")
        if args.date_from or args.date_to:
            LOGGER.info("Periode forcee demandee: %s -> %s", args.date_from, args.date_to)
        if args.debug_invoice:
            LOGGER.info("Debug facture demande: %s", args.debug_invoice)
        if args.debug_reference:
            LOGGER.info("Debug referentiel RCI demande.")

        inventory = detect_files(project_root, config, use_samples=args.use_samples)
        _log_inventory(inventory)
        _validate_inventory(inventory, ignore_pdf=args.ignore_pdf)

        LOGGER.info("Etape 4/12 - Extraction ERP")
        erp_records = extract_erp_files(inventory.erp_files)

        LOGGER.info("Etape 4b/12 - Chargement referentiel couverture RCI")
        reference_config = config.get("reference", {}).get("rci_coverage", {})
        reference_enabled = _as_bool(reference_config.get("enabled", True))
        reference_path = _resolve_project_path(
            project_root,
            reference_config.get("path", config.get("paths", {}).get("reference_root", "reference")),
        )
        reference_debug_path: Path | None = None
        if args.debug_reference:
            reference_debug_path = inspect_reference_file(
                reference_path,
                project_root / config["paths"]["anomalies_dir"],
                run_id,
            )
        coverage_reference = load_rci_coverage_reference(reference_path) if reference_enabled else None
        erp_records = enrich_erp_with_rci_coverage(
            erp_records,
            coverage_reference,
            enabled=reference_enabled,
        )

        LOGGER.info("Etape 5/12 - Extraction RCI")
        rci_records = extract_rci_files(inventory.rci_files)

        LOGGER.info("Etape 6/12 - Extraction PDF")
        if args.ignore_pdf:
            LOGGER.info("Option --ignore-pdf active: extraction et integration PDF ignorees.")
            pdf_records = extract_pdf_files([])
        else:
            pdf_records = extract_pdf_files(inventory.pdf_files)

        LOGGER.info("Etape 6b/12 - Filtrage periode ERP")
        try:
            date_filter_result = apply_reconciliation_date_filter(
                erp_records,
                rci_records,
                pdf_records,
                config.get("reconciliation", {}).get("date_filter", {}),
                disable=args.no_date_filter,
                date_from=args.date_from,
                date_to=args.date_to,
            )
        except ValueError as exc:
            raise PipelineError(str(exc)) from exc
        filtered_erp_records = date_filter_result.erp_records
        filtered_rci_records = date_filter_result.rci_records
        filtered_pdf_records = date_filter_result.pdf_records

        LOGGER.info("Etape 7/12 - Reconciliation ERP / RCI / PDF")
        reconciliation_config = config.get("reconciliation", {})
        amount_tolerance = float(
            reconciliation_config.get("amount_tolerance", reconciliation_config.get("amount_tolerance_mad", 1.0))
        )
        report = reconcile(
            filtered_erp_records,
            filtered_rci_records,
            filtered_pdf_records,
            {},
            amount_tolerance=amount_tolerance,
            rci_out_of_period_records=date_filter_result.rci_out_of_period_records,
            pdf_out_of_period_records=date_filter_result.pdf_out_of_period_records,
        )
        report.setdefault("summary", {}).update(date_filter_result.summary)
        reference_names = (
            coverage_reference.name_values
            if coverage_reference is not None and coverage_reference.loaded
            else []
        )
        report = enrich_report_with_audits(report, reference_names)
        reference_quality_source = (
            report.get("audits", {}).get("out_of_scope_rci")
            or report.get("reconciliation", [])
        )
        report["reference_quality"] = build_reference_suggestions(reference_quality_source)

        anomalies_dir = project_root / config["paths"]["anomalies_dir"]
        matching_debug_path = write_matching_debug(
            filtered_erp_records,
            filtered_rci_records,
            filtered_pdf_records,
            anomalies_dir,
            run_id,
        )
        log_invoice_presence(erp_records, rci_records, pdf_records)
        if args.debug_invoice:
            log_debug_invoice(
                args.debug_invoice,
                erp_records,
                rci_records,
                pdf_records,
                report.get("reconciliation", []),
            )

        LOGGER.info("Etapes 8-9/12 - Generation rapport Excel et exports Power BI")
        artifacts = [
            artifact
            for artifact in [
                reference_debug_path,
                matching_debug_path,
                *_write_outputs(project_root, config, report, run_id),
            ]
            if artifact is not None
        ]

        LOGGER.info("Etape 10/12 - Email")
        email_summary = {**report.get("summary", {}), "generated_at": report.get("generated_at")}
        email_status = send_report(config, artifacts, dry_run=args.dry_run, summary=email_summary)
        LOGGER.info("Statut email: %s", email_status)

        LOGGER.info("Etape 11/12 - Archivage")
        archived_files: list[Path] = []
        if args.dry_run:
            LOGGER.info("Dry-run actif: archivage ignore.")
        elif args.no_archive:
            LOGGER.info("Archivage desactive par option --no-archive.")
        elif args.use_samples:
            LOGGER.info("Mode samples: les exemples ne sont jamais archives.")
        else:
            archived_files = _archive_inputs(project_root, config, inventory, run_id)
            LOGGER.info("Fichiers archives: %s", len(archived_files))

        LOGGER.info("Etape 12/12 - Synthese finale")
        _log_final_summary(report, artifacts, email_status, archived_files, log_file)
        LOGGER.info("Execution terminee avec succes.")
        return 0

    except PipelineError as exc:
        LOGGER.error("Execution interrompue: %s", exc)
        if log_file is not None:
            LOGGER.error("Rapport de logs: %s", log_file)
        return 2
    except Exception as exc:
        if logging.getLogger().handlers:
            LOGGER.exception("Erreur inattendue pendant l'execution: %s", exc)
        else:
            print(f"Erreur inattendue pendant l'execution: {exc}", file=sys.stderr)
        if log_file is not None:
            LOGGER.error("Rapport de logs: %s", log_file)
        return 1


def _resolve_config_path(project_root: Path, config_arg: str | None) -> Path:
    if not config_arg:
        return project_root / "config" / "config.yaml"

    config_path = Path(config_arg)
    if config_path.is_absolute():
        return config_path
    return project_root / config_path


def _resolve_project_path(project_root: Path, value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return project_root / path


def _setup_logging(
    project_root: Path, config: dict[str, Any], run_id: str, log_level_arg: str | None
) -> Path:
    logs_root = project_root / config["paths"]["logs_root"]
    logs_root.mkdir(parents=True, exist_ok=True)
    log_file = logs_root / f"autorcicontrol_{run_id}.log"
    level_name = log_level_arg or config.get("logging", {}).get("level", "INFO")
    level = getattr(logging, level_name.upper(), logging.INFO)

    logging.basicConfig(
        level=level,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,
    )
    return log_file


def _log_inventory(inventory: FileInventory) -> None:
    LOGGER.info("Dossier analyse: %s", inventory.base_dir)
    for source_type in SOURCE_TYPES:
        files = inventory.files_for(source_type)
        LOGGER.info("%s: %s fichier(s) detecte(s)", source_type.upper(), len(files))
        for path in files:
            LOGGER.info(" - %s", path.name)

    if inventory.missing_required:
        LOGGER.warning("Sources incompletes: %s", inventory.missing_required)


def _validate_inventory(inventory: FileInventory, ignore_pdf: bool = False) -> None:
    if not inventory.erp_files:
        raise PipelineError(
            f"Aucun fichier ERP trouve dans {inventory.base_dir / 'erp'}. "
            "Deposez au moins un fichier Excel ou CSV ERP avant de relancer."
        )

    if ignore_pdf and not inventory.rci_files:
        raise PipelineError(
            f"Aucun fichier RCI trouve dans {inventory.base_dir / 'rci'} alors que --ignore-pdf est actif. "
            "Deposez au moins un fichier RCI TXT/CSV/Excel avant de relancer."
        )

    if not inventory.rci_files and not inventory.pdf_files:
        raise PipelineError(
            f"Aucun fichier RCI/PDF trouve dans {inventory.base_dir / 'rci'} et "
            f"{inventory.base_dir / 'pdf'}. Deposez au moins un fichier RCI TXT/CSV/Excel "
            "ou un PDF banque/RCI avant de relancer."
        )

    if ignore_pdf:
        LOGGER.info(
            "Option --ignore-pdf active: les fichiers PDF detectes ne seront pas utilises pour le rapprochement."
        )
    elif not inventory.pdf_files:
        LOGGER.warning(
            "Aucun PDF detecte dans %s. Le pipeline continue avec ERP/RCI; "
            "les dates d'echeance et origines PDF seront absentes.",
            inventory.base_dir / "pdf",
        )

    if not inventory.rci_files:
        LOGGER.warning(
            "Aucun fichier RCI TXT/CSV/Excel detecte dans %s. Le pipeline continue avec ERP/PDF; "
            "la presence dans le flux RCI ne pourra pas etre confirmee.",
            inventory.base_dir / "rci",
        )


def _write_outputs(
    project_root: Path, config: dict[str, Any], report: dict[str, Any], run_id: str
) -> list[Path]:
    if not config.get("processing", {}).get("output_reports", True):
        LOGGER.info("Generation des rapports desactivee.")
        return []

    reports_dir = project_root / config["paths"]["reports_dir"]
    powerbi_dir = project_root / config["paths"]["powerbi_dir"]
    anomalies_dir = project_root / config["paths"]["anomalies_dir"]
    corrections_dir = project_root / config["paths"].get("corrections_dir", "output/corrections")

    corrective_artifacts = write_corrective_batch_outputs(report, corrections_dir, run_id)
    missing_rci_export = write_missing_rci_export(report, anomalies_dir, run_id)

    artifacts = [
        write_excel_report(report, reports_dir, run_id),
        *write_powerbi_exports(report, powerbi_dir, run_id),
        missing_rci_export,
        _write_anomaly_export(report, anomalies_dir, run_id),
        write_reference_suggestions(
            report.get("audits", {}).get("out_of_scope_rci") or report.get("reconciliation", []),
            anomalies_dir,
            run_id,
        ),
        *corrective_artifacts,
    ]

    for artifact in artifacts:
        LOGGER.info("Artefact genere: %s", artifact)
    return artifacts


def _log_final_summary(
    report: dict[str, Any],
    artifacts: list[Path],
    email_status: str,
    archived_files: list[Path],
    log_file: Path,
) -> None:
    summary = report.get("summary", {})
    report_path = _first_artifact(artifacts, ".xlsx")
    history_path = next(
        (artifact for artifact in artifacts if artifact.name == "reconciliation_history.csv"),
        None,
    )

    LOGGER.info("")
    LOGGER.info("===== SYNTHESE FINALE AutoRCI =====")
    LOGGER.info("Statut global: %s", report.get("status", ""))
    LOGGER.info("Factures ERP analysees: %s", summary.get("erp_analyzed_invoices", summary.get("reconciled_invoices", 0)))
    LOGGER.info("Periode de rapprochement: %s", summary.get("reconciliation_period", "non renseignee"))
    LOGGER.info("ERP avant filtre date: %s", summary.get("erp_rows_before_date_filter", summary.get("erp_rows", 0)))
    LOGGER.info("ERP apres filtre date: %s", summary.get("erp_rows_after_date_filter", summary.get("erp_rows", 0)))
    LOGGER.info("ERP exclu par date: %s", summary.get("erp_rows_excluded_by_date", 0))
    LOGGER.info("RCI avant filtre date: %s", summary.get("rci_rows_before_date_filter", summary.get("rci_rows", 0)))
    LOGGER.info("RCI apres filtre date: %s", summary.get("rci_rows_after_date_filter", summary.get("rci_rows", 0)))
    LOGGER.info("RCI exclu par periode: %s", summary.get("rci_rows_excluded_by_date", 0))
    LOGGER.info("OK: %s", summary.get("matched_invoices", 0))
    LOGGER.info("Manquantes RCI: %s", summary.get("unmatched_erp", 0))
    LOGGER.info("Factures absentes RCI: %s", summary.get("missing_rci_invoice_count", 0))
    LOGGER.info("Avoirs absents RCI: %s", summary.get("missing_rci_credit_note_count", 0))
    LOGGER.info(
        "Montant total absent RCI: %s MAD",
        summary.get("missing_rci_total_amount", summary.get("missing_rci_amount", 0)),
    )
    LOGGER.info("Hors scope RCI: %s", summary.get("out_of_scope_rci", 0))
    LOGGER.info("RCI seulement: %s", summary.get("unmatched_rci", 0))
    LOGGER.info("RCI hors periode: %s", summary.get("rci_out_of_period", 0))
    LOGGER.info(
        "Total lignes RCI/PDF hors periode: %s",
        summary.get("total_rci_pdf_out_of_period", summary.get("rci_pdf_rows_excluded_by_date", summary.get("rci_out_of_period", 0))),
    )
    LOGGER.info("Anomalies montant: %s", summary.get("amount_anomalies", 0))
    LOGGER.info("Anomalies date: %s", summary.get("date_anomalies", 0))
    LOGGER.info("Doublons: %s", summary.get("duplicates", 0))
    LOGGER.info("Ecarts detectes: %s", summary.get("gaps_detected", summary.get("anomalies", 0)))
    LOGGER.info("Ecarts critiques: %s", summary.get("critical_gaps", summary.get("severity_critique", 0)))
    LOGGER.info("Ecarts eleves: %s", summary.get("high_gaps", summary.get("severity_elevee", 0)))
    LOGGER.info("Ecarts moyens: %s", summary.get("medium_gaps", summary.get("severity_moyenne", 0)))
    LOGGER.info("Lignes information: %s", summary.get("information_lines", summary.get("severity_information", 0)))
    LOGGER.info(
        "Lignes a verifier referentiel: %s",
        summary.get("reference_review_lines", summary.get("severity_a_verifier", 0)),
    )
    LOGGER.info("Batch correctif candidat: %s", "oui" if summary.get("corrective_batch_generated") else "non")
    LOGGER.info("Factures incluses batch correctif: %s", summary.get("corrective_batch_invoice_count", 0))
    LOGGER.info("Montant batch correctif: %s MAD", summary.get("corrective_batch_total_amount", 0))
    LOGGER.info("Montant impacte total: %s MAD", summary.get("total_impacted_amount", summary.get("total_amount_gap", 0)))
    LOGGER.info("Montant manquant RCI: %s MAD", summary.get("missing_rci_amount", 0))
    LOGGER.info("Taux de rapprochement: %s", summary.get("matching_rate", 0))
    if summary.get("no_rci_flux_in_period_alert"):
        LOGGER.warning("Attention : aucun flux RCI dans la période de rapprochement.")
    if summary.get("pdf_period_mismatch_alert"):
        LOGGER.warning("Attention : les PDF chargés ne correspondent pas à la période du flux RCI.")
    if not summary.get("corrective_batch_generated") and report.get("corrective_batch", {}).get("warning"):
        LOGGER.info("%s", report["corrective_batch"]["warning"])
    LOGGER.info("Rapport Excel: %s", report_path or "non genere")
    LOGGER.info("Historique Power BI: %s", history_path or "non genere")
    LOGGER.info("Email: %s", email_status)
    LOGGER.info("Fichiers archives: %s", len(archived_files))
    LOGGER.info("Logs: %s", log_file)
    LOGGER.info("===================================")


def _first_artifact(artifacts: list[Path], suffix: str) -> Path | None:
    for artifact in artifacts:
        if artifact.suffix.lower() == suffix:
            return artifact
    return None


def _write_anomaly_export(report: dict[str, Any], anomalies_dir: Path, run_id: str) -> Path:
    import json

    anomalies_dir.mkdir(parents=True, exist_ok=True)
    path = anomalies_dir / f"anomalies_{run_id}.json"
    path.write_text(json.dumps(report["anomalies"], ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _archive_inputs(
    project_root: Path, config: dict[str, Any], inventory: FileInventory, run_id: str
) -> list[Path]:
    if not config.get("processing", {}).get("archive_after_success", True):
        LOGGER.info("Archivage desactive dans la configuration.")
        return []

    archive_root = project_root / config["paths"]["archive_root"]
    archived_files: list[Path] = []

    for source_type in SOURCE_TYPES:
        destination_dir = archive_root / source_type / run_id
        destination_dir.mkdir(parents=True, exist_ok=True)
        for source_file in inventory.files_for(source_type):
            destination = destination_dir / source_file.name
            destination = _deduplicate_destination(destination)
            shutil.move(str(source_file), str(destination))
            archived_files.append(destination)
            LOGGER.info("Archive: %s -> %s", source_file, destination)

    return archived_files


def _deduplicate_destination(destination: Path) -> Path:
    if not destination.exists():
        return destination

    counter = 1
    while True:
        candidate = destination.with_name(f"{destination.stem}_{counter}{destination.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on", "oui"}


if __name__ == "__main__":
    raise SystemExit(main())
