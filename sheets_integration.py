"""Read-only Google Sheets integration scaffolding for ExhibitLedger THB.

This module intentionally performs **no write-back**. It is designed to let the bot
check configuration, read sheet headers, and preview how exhibition-related tabs
will be mapped before any finance data is imported into SQLite.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List


DEFAULT_MAPPING_FILE = os.environ.get("SHEETS_MAPPING_FILE", "./sheets_mapping.example.json")


@dataclass
class SheetsConfig:
    spreadsheet_id: str | None
    credentials_path: str | None
    mapping_file: str

    @property
    def is_configured(self) -> bool:
        return bool(self.spreadsheet_id and self.credentials_path and Path(self.credentials_path).exists())


def get_sheets_config() -> SheetsConfig:
    return SheetsConfig(
        spreadsheet_id=os.environ.get("GOOGLE_SHEETS_SPREADSHEET_ID") or None,
        credentials_path=os.environ.get("GOOGLE_APPLICATION_CREDENTIALS") or None,
        mapping_file=os.environ.get("SHEETS_MAPPING_FILE", DEFAULT_MAPPING_FILE),
    )


def load_mapping(path: str | None = None) -> Dict[str, Any]:
    mapping_path = Path(path or DEFAULT_MAPPING_FILE)
    if not mapping_path.exists():
        raise FileNotFoundError(f"Mapping file not found: {mapping_path}")
    return json.loads(mapping_path.read_text(encoding="utf-8"))


def validate_mapping(mapping: Dict[str, Any]) -> List[str]:
    warnings: List[str] = []
    required_top = ["version", "currency", "exhibition_key_column", "sheets"]
    for key in required_top:
        if key not in mapping:
            warnings.append(f"Missing mapping key: {key}")
    if mapping.get("currency") != "THB":
        warnings.append("Mapping currency must be THB for this bot.")
    sheets = mapping.get("sheets", {})
    for logical_name in ["cash_book", "stock_sales", "commission", "estimated_pnl"]:
        if logical_name not in sheets:
            warnings.append(f"Missing sheet mapping: {logical_name}")
            continue
        sheet_cfg = sheets[logical_name]
        if not sheet_cfg.get("tab_name"):
            warnings.append(f"Sheet mapping {logical_name} is missing tab_name.")
        columns = sheet_cfg.get("columns", {})
        if not isinstance(columns, dict) or not columns:
            warnings.append(f"Sheet mapping {logical_name} has no columns dictionary.")
    if not warnings:
        warnings.append("Mapping structure looks valid for read-only preview.")
    return warnings


def _open_spreadsheet(config: SheetsConfig):
    try:
        import gspread  # type: ignore
    except ImportError as exc:
        raise RuntimeError("gspread is not installed. Run: pip install -r requirements.txt") from exc

    if not config.spreadsheet_id:
        raise RuntimeError("GOOGLE_SHEETS_SPREADSHEET_ID is not set.")
    if not config.credentials_path:
        raise RuntimeError("GOOGLE_APPLICATION_CREDENTIALS is not set.")
    if not Path(config.credentials_path).exists():
        raise RuntimeError(f"Google credentials file does not exist: {config.credentials_path}")

    client = gspread.service_account(filename=config.credentials_path)
    return client.open_by_key(config.spreadsheet_id)


def configuration_status() -> Dict[str, Any]:
    config = get_sheets_config()
    status = {
        "spreadsheet_id_set": bool(config.spreadsheet_id),
        "credentials_path_set": bool(config.credentials_path),
        "credentials_file_exists": bool(config.credentials_path and Path(config.credentials_path).exists()),
        "mapping_file": config.mapping_file,
        "mapping_file_exists": Path(config.mapping_file).exists(),
        "read_only_mode": True,
    }
    if status["mapping_file_exists"]:
        try:
            status["mapping_warnings"] = validate_mapping(load_mapping(config.mapping_file))
        except Exception as exc:  # pragma: no cover - shown to user as setup feedback
            status["mapping_warnings"] = [f"Mapping could not be read: {exc}"]
    else:
        status["mapping_warnings"] = ["Mapping file is missing."]
    return status


def preview_workbook(limit_rows: int = 5) -> Dict[str, Any]:
    """Return a read-only workbook preview using the mapping file.

    The preview includes worksheet presence, headers, and a small row count/sample.
    It does not update SQLite and does not write to Google Sheets.
    """
    config = get_sheets_config()
    mapping = load_mapping(config.mapping_file)
    mapping_warnings = validate_mapping(mapping)
    workbook = _open_spreadsheet(config)
    existing_titles = {ws.title: ws for ws in workbook.worksheets()}

    preview: Dict[str, Any] = {
        "spreadsheet_title": workbook.title,
        "read_only_mode": True,
        "mapping_warnings": mapping_warnings,
        "sheets": [],
    }

    for logical_name, sheet_cfg in mapping.get("sheets", {}).items():
        tab_name = sheet_cfg.get("tab_name")
        item: Dict[str, Any] = {"logical_name": logical_name, "tab_name": tab_name, "exists": tab_name in existing_titles}
        if tab_name in existing_titles:
            ws = existing_titles[tab_name]
            values = ws.get_all_values()
            item["row_count"] = max(len(values) - 1, 0)
            item["headers"] = values[0] if values else []
            item["sample_rows"] = values[1 : 1 + limit_rows] if len(values) > 1 else []
        else:
            item["row_count"] = 0
            item["headers"] = []
            item["sample_rows"] = []
        preview["sheets"].append(item)
    return preview


def format_sheets_status_markdown() -> str:
    status = configuration_status()
    lines = ["*Google Sheets Read-Only Status*", ""]
    lines.append(f"Spreadsheet ID set: {'YES' if status['spreadsheet_id_set'] else 'NO'}")
    lines.append(f"Credentials path set: {'YES' if status['credentials_path_set'] else 'NO'}")
    lines.append(f"Credentials file exists: {'YES' if status['credentials_file_exists'] else 'NO'}")
    lines.append(f"Mapping file: `{status['mapping_file']}`")
    lines.append(f"Mapping file exists: {'YES' if status['mapping_file_exists'] else 'NO'}")
    lines.append("Mode: *READ-ONLY*. No sheet write-back is enabled.")
    lines.append("")
    lines.append("Mapping checks:")
    lines.extend(f"• {warning}" for warning in status.get("mapping_warnings", []))
    lines.append("")
    lines.append("To activate preview, set `GOOGLE_SHEETS_SPREADSHEET_ID`, `GOOGLE_APPLICATION_CREDENTIALS`, and share the workbook with the service-account email.")
    return "\n".join(lines)


def format_sync_preview_markdown(limit_rows: int = 3) -> str:
    preview = preview_workbook(limit_rows=limit_rows)
    lines = ["*Google Sheets Read-Only Preview*", "", f"Workbook: *{preview['spreadsheet_title']}*", "Mode: *READ-ONLY*", ""]
    if preview.get("mapping_warnings"):
        lines.append("Mapping checks:")
        lines.extend(f"• {warning}" for warning in preview["mapping_warnings"])
        lines.append("")
    for item in preview["sheets"]:
        exists = "YES" if item["exists"] else "NO"
        lines.append(f"*{item['logical_name']}* → `{item['tab_name']}` — exists: {exists}; rows: {item['row_count']}")
        if item["headers"]:
            header_text = ", ".join(str(h) for h in item["headers"][:10])
            lines.append(f"Headers: {header_text}")
    lines.append("")
    lines.append("This command only previews workbook structure. It does not import, overwrite, or edit finance data.")
    return "\n".join(lines)
