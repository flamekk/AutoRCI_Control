from __future__ import annotations

import logging
from types import SimpleNamespace

import pytest

from src.file_detector import FileInventory
from src.main import PipelineError, _validate_inventory, parse_args
import src.main as main_module


def test_validate_inventory_blocks_when_no_erp(tmp_path) -> None:
    inventory = FileInventory(
        mode="input",
        base_dir=tmp_path,
        erp_files=[],
        rci_files=[tmp_path / "rci.txt"],
        pdf_files=[],
        missing_required={},
    )

    with pytest.raises(PipelineError, match="Aucun fichier ERP"):
        _validate_inventory(inventory)


def test_validate_inventory_blocks_when_no_rci_or_pdf(tmp_path) -> None:
    erp_file = tmp_path / "erp.xlsx"
    erp_file.write_text("fake", encoding="utf-8")
    inventory = FileInventory(
        mode="input",
        base_dir=tmp_path,
        erp_files=[erp_file],
        rci_files=[],
        pdf_files=[],
        missing_required={},
    )

    with pytest.raises(PipelineError, match="Aucun fichier RCI/PDF"):
        _validate_inventory(inventory)


def test_validate_inventory_allows_missing_pdf_with_warning(tmp_path, caplog) -> None:
    erp_file = tmp_path / "erp.xlsx"
    rci_file = tmp_path / "rci.txt"
    erp_file.write_text("fake", encoding="utf-8")
    rci_file.write_text("fake", encoding="utf-8")
    inventory = FileInventory(
        mode="input",
        base_dir=tmp_path,
        erp_files=[erp_file],
        rci_files=[rci_file],
        pdf_files=[],
        missing_required={},
    )

    with caplog.at_level(logging.WARNING):
        _validate_inventory(inventory)

    assert "Aucun PDF detecte" in caplog.text


def test_ignore_pdf_option() -> None:
    args = parse_args(["--use-samples", "--dry-run", "--ignore-pdf"])

    assert args.use_samples is True
    assert args.dry_run is True
    assert args.ignore_pdf is True


def test_no_archive_option_is_accepted_by_parser() -> None:
    args = parse_args(["--no-archive", "--date-from", "2026-04-29", "--date-to", "2026-05-05"])

    assert args.no_archive is True
    assert args.dry_run is False


def test_no_archive_option_skips_archive_step(monkeypatch, tmp_path) -> None:
    config_path = tmp_path / "config.yaml"
    config_path.write_text(
        """
paths:
  input_root: input
  samples_root: samples
  reference_root: reference
  output_root: output
  archive_root: archive
  logs_root: logs
  reports_dir: output/reports
  powerbi_dir: output/powerbi
  anomalies_dir: output/anomalies
processing:
  archive_after_success: true
  output_reports: true
email:
  enabled: false
""".strip(),
        encoding="utf-8",
    )
    erp_file = tmp_path / "input" / "erp" / "erp.xlsx"
    rci_file = tmp_path / "input" / "rci" / "rci.txt"
    erp_file.parent.mkdir(parents=True)
    rci_file.parent.mkdir(parents=True)
    erp_file.write_text("fake", encoding="utf-8")
    rci_file.write_text("fake", encoding="utf-8")
    archive_called = {"value": False}

    monkeypatch.setattr(main_module, "get_project_root", lambda: tmp_path)
    monkeypatch.setattr(
        main_module,
        "detect_files",
        lambda project_root, config, use_samples=False: FileInventory(
            mode="input",
            base_dir=tmp_path / "input",
            erp_files=[erp_file],
            rci_files=[rci_file],
            pdf_files=[],
            missing_required={},
        ),
    )
    monkeypatch.setattr(main_module, "extract_erp_files", lambda files: [])
    monkeypatch.setattr(main_module, "extract_rci_files", lambda files: [])
    monkeypatch.setattr(main_module, "extract_pdf_files", lambda files: [])
    monkeypatch.setattr(main_module, "load_rci_coverage_reference", lambda path: None)
    monkeypatch.setattr(main_module, "enrich_erp_with_rci_coverage", lambda records, reference, enabled=True: records)
    monkeypatch.setattr(
        main_module,
        "apply_reconciliation_date_filter",
        lambda *args, **kwargs: SimpleNamespace(
            erp_records=[],
            rci_records=[],
            pdf_records=[],
            rci_out_of_period_records=[],
            pdf_out_of_period_records=[],
            summary={},
        ),
    )
    monkeypatch.setattr(
        main_module,
        "reconcile",
        lambda *args, **kwargs: {
            "generated_at": "2026-06-01T00:00:00+00:00",
            "status": "ok",
            "summary": {},
            "reconciliation": [],
            "source_files": [],
            "anomalies": [],
        },
    )
    monkeypatch.setattr(main_module, "enrich_report_with_audits", lambda report, reference_names: report)
    monkeypatch.setattr(main_module, "write_matching_debug", lambda *args, **kwargs: tmp_path / "matching_debug.csv")
    monkeypatch.setattr(main_module, "log_invoice_presence", lambda *args, **kwargs: None)
    report_path = tmp_path / "output" / "reports" / "Rapport_Reconciliation_RCI_2026-06-01_0000.xlsx"
    report_path.parent.mkdir(parents=True)
    report_path.write_bytes(b"fake")
    monkeypatch.setattr(main_module, "_write_outputs", lambda *args, **kwargs: [report_path])
    monkeypatch.setattr(main_module, "send_report", lambda *args, **kwargs: "disabled")
    monkeypatch.setattr(main_module, "_log_final_summary", lambda *args, **kwargs: None)

    def fake_archive(*args, **kwargs):
        archive_called["value"] = True
        return []

    monkeypatch.setattr(main_module, "_archive_inputs", fake_archive)

    result = main_module.main(["--no-archive", "--config", str(config_path)])

    assert result == 0
    assert archive_called["value"] is False
