from __future__ import annotations

import logging

import pytest

from src.file_detector import FileInventory
from src.main import PipelineError, _validate_inventory, parse_args


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
