from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


SOURCE_TYPES = ("erp", "rci", "pdf")


@dataclass(frozen=True)
class FileInventory:
    mode: str
    base_dir: Path
    erp_files: list[Path]
    rci_files: list[Path]
    pdf_files: list[Path]
    missing_required: dict[str, int]

    @property
    def all_files(self) -> list[Path]:
        return [*self.erp_files, *self.rci_files, *self.pdf_files]

    def files_for(self, source_type: str) -> list[Path]:
        return list(getattr(self, f"{source_type}_files"))

    def counts(self) -> dict[str, int]:
        return {
            "erp": len(self.erp_files),
            "rci": len(self.rci_files),
            "pdf": len(self.pdf_files),
        }


def detect_files(project_root: Path, config: dict[str, Any], use_samples: bool) -> FileInventory:
    mode = "samples" if use_samples else "input"
    base_root_key = "samples_root" if use_samples else "input_root"
    base_dir = project_root / config["paths"][base_root_key]

    extensions = config["file_detection"]["extensions"]
    min_required = config["file_detection"].get("min_required", {})

    detected: dict[str, list[Path]] = {}
    missing_required: dict[str, int] = {}

    for source_type in SOURCE_TYPES:
        source_dir = base_dir / source_type
        detected[source_type] = _scan_source_dir(source_dir, extensions.get(source_type, []))

        required_count = int(min_required.get(source_type, 0))
        if len(detected[source_type]) < required_count:
            missing_required[source_type] = required_count - len(detected[source_type])

    return FileInventory(
        mode=mode,
        base_dir=base_dir,
        erp_files=detected["erp"],
        rci_files=detected["rci"],
        pdf_files=detected["pdf"],
        missing_required=missing_required,
    )


def _scan_source_dir(source_dir: Path, allowed_extensions: list[str]) -> list[Path]:
    if not source_dir.exists():
        return []

    normalized_extensions = {extension.lower() for extension in allowed_extensions}
    candidates = [
        path
        for path in source_dir.iterdir()
        if path.is_file() and path.suffix.lower() in normalized_extensions
    ]
    return sorted(candidates, key=lambda path: (path.stat().st_mtime, path.name.lower()))
