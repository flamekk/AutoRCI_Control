from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]

DEFAULT_CONFIG: dict[str, Any] = {
    "project": {"name": "AutoRCI_Control"},
    "paths": {
        "input_root": "input",
        "samples_root": "samples",
        "reference_root": "reference",
        "output_root": "output",
        "archive_root": "archive",
        "logs_root": "logs",
        "reports_dir": "output/reports",
        "powerbi_dir": "output/powerbi",
        "anomalies_dir": "output/anomalies",
        "corrections_dir": "output/corrections",
    },
    "file_detection": {
        "extensions": {
            "erp": [".xlsx", ".xls", ".csv"],
            "rci": [".txt", ".csv", ".xlsx", ".xls"],
            "pdf": [".pdf"],
        },
        "min_required": {"erp": 0, "rci": 0, "pdf": 0},
    },
    "processing": {
        "archive_after_success": True,
        "output_reports": True,
        "run_timezone": "Africa/Casablanca",
    },
    "reconciliation": {
        "amount_tolerance": 1.0,
        "amount_tolerance_mad": 1.0,
        "date_filter": {
            "enabled": True,
            "mode": "auto",
            "days_before": 3,
            "days_after": 1,
            "fallback_days": 7,
        },
    },
    "reference": {
        "rci_coverage": {
            "enabled": True,
            "path": "reference",
        },
    },
    "email": {
        "enabled": False,
        "sender": "",
        "recipients": [],
        "smtp_host": "",
        "smtp_server": "",
        "smtp_port": 587,
        "username": "",
        "smtp_username": "",
        "password_env_var": "AUTORCI_EMAIL_PASSWORD",
        "smtp_password_env_var": "AUTORCI_SMTP_PASSWORD",
        "use_tls": True,
        "use_ssl": False,
    },
}

STRUCTURE_DIRECTORIES = [
    "input/erp",
    "input/rci",
    "input/pdf",
    "samples/erp",
    "samples/rci",
    "samples/pdf",
    "reference",
    "output/reports",
    "output/powerbi",
    "output/anomalies",
    "output/corrections",
    "archive/erp",
    "archive/rci",
    "archive/pdf",
    "logs",
    "src",
    "tests",
    "config",
]


class ConfigError(RuntimeError):
    """Raised when the application configuration cannot be loaded."""


def get_project_root() -> Path:
    return PROJECT_ROOT


def resolve_project_path(project_root: Path, value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return project_root / path


def ensure_project_structure(project_root: Path) -> None:
    for directory in STRUCTURE_DIRECTORIES:
        (project_root / directory).mkdir(parents=True, exist_ok=True)


def load_config(config_path: Path | None = None) -> dict[str, Any]:
    config_path = config_path or PROJECT_ROOT / "config" / "config.yaml"
    config = deepcopy(DEFAULT_CONFIG)

    if not config_path.exists():
        return config

    raw = config_path.read_text(encoding="utf-8")
    loaded = _load_yaml(raw, config_path)
    if not isinstance(loaded, dict):
        raise ConfigError(f"Configuration invalide dans {config_path}: racine YAML attendue.")

    return _deep_merge(config, loaded)


def _load_yaml(raw: str, config_path: Path) -> Any:
    try:
        import yaml  # type: ignore
    except ModuleNotFoundError:
        return _load_yaml_without_dependency(raw)

    try:
        return yaml.safe_load(raw) or {}
    except Exception as exc:  # pragma: no cover - defensive around optional dependency.
        raise ConfigError(f"Impossible de lire {config_path}: {exc}") from exc


def _load_yaml_without_dependency(raw: str) -> dict[str, Any]:
    """Small YAML subset loader for this project's simple config file.

    The production dependency is PyYAML. This fallback keeps the bootstrap command
    usable before dependencies are installed.
    """

    root: dict[str, Any] = {}
    stack: list[tuple[int, dict[str, Any]]] = [(-1, root)]

    for line_number, raw_line in enumerate(raw.splitlines(), start=1):
        line = raw_line.split("#", 1)[0].rstrip()
        if not line.strip():
            continue
        if ":" not in line:
            raise ConfigError(f"YAML non supporte ligne {line_number}: {raw_line}")

        indent = len(line) - len(line.lstrip(" "))
        key, value = line.strip().split(":", 1)
        key = key.strip()
        value = value.strip()

        while stack and indent <= stack[-1][0]:
            stack.pop()
        if not stack:
            raise ConfigError(f"Indentation YAML invalide ligne {line_number}: {raw_line}")

        parent = stack[-1][1]
        if value == "":
            child: dict[str, Any] = {}
            parent[key] = child
            stack.append((indent, child))
        else:
            parent[key] = _parse_scalar(value)

    return root


def _parse_scalar(value: str) -> Any:
    lower_value = value.lower()
    if lower_value == "true":
        return True
    if lower_value == "false":
        return False
    if lower_value in {"null", "none", "~"}:
        return None
    if value.startswith("[") and value.endswith("]"):
        return json.loads(value)
    if value.startswith('"') and value.endswith('"'):
        return value[1:-1]
    if value.startswith("'") and value.endswith("'"):
        return value[1:-1]
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value)
    except ValueError:
        return value


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(base.get(key), dict):
            base[key] = _deep_merge(base[key], value)
        else:
            base[key] = value
    return base
