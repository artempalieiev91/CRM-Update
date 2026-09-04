"""Збереження робочого стану CRM Update між оновленнями сторінки."""

from __future__ import annotations

import json
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd

from ids_runtime import log_disk_failure, log_disk_write

WORKSPACE_VERSION = 1
WORKSPACE_DIR_NAME = ".crm_update_workspace"
ORGANIZATIONS_CRM_FILENAME = "organizations_crm.csv"
PERSONS_CRM_FILENAME = "persons_crm.csv"

# Згенеровані файли в корені проєкту (див. _export_paths у app.py).
EXPORT_ARTIFACT_GLOBS: tuple[str, ...] = (
    "* - comp.csv",
    "* - linkedin.csv",
    "* - companies.csv",
    "* - companies-duplicates.csv",
    "* - companies-org-issues.csv",
    "* - companies-review.xlsx",
    "* - leads-person-issues.csv",
    "* - leads-person-issues.xlsx",
    "* - research-duplicates.csv",
    "* - activities all highlighted.xlsx",
    "* - activities issues.xlsx",
    "* - activities delete.xlsx",
    "* - activities filtered.xlsx",
    "* - activities CRM.xlsx",
)

_DATAFRAME_KEYS: tuple[str, ...] = (
    "full_df",
    "pending_df",
    "comp_df",
    "linkedin_df",
    "companies_df",
    "companies_merge_preview",
)

_META_KEYS: tuple[str, ...] = (
    "active_tab",
    "file_stem",
    "data_source_label",
    "pending_only",
    "crm_status_filter",
    "load_error",
    "crm_leads_name",
    "crm_companies_name",
    "crm_leads_key",
    "crm_companies_key",
    "research_csv_key",
    "crm_organizations_name",
    "crm_organizations_key",
    "organizations_row_count",
    "organizations_load_error",
    "crm_persons_name",
    "crm_persons_key",
    "persons_row_count",
    "persons_load_error",
    "fill_notion_campaign_id",
    "leads_default_activity_type",
)


def organizations_crm_path(root: Path) -> Path:
    return workspace_dir(root) / ORGANIZATIONS_CRM_FILENAME


def save_organizations_crm(root: Path, df: pd.DataFrame) -> Path:
    path = organizations_crm_path(root)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(path, index=False, encoding="utf-8-sig")
    except OSError as exc:
        log_disk_failure(path, exc, action="write organizations CRM")
        raise
    log_disk_write(path, action="write organizations CRM")
    return path


def load_organizations_crm(root: Path) -> pd.DataFrame | None:
    path = organizations_crm_path(root)
    if not path.is_file():
        return None
    return _read_df(path)


def organizations_crm_exists(root: Path) -> bool:
    return organizations_crm_path(root).is_file()


def workspace_dir(root: Path) -> Path:
    return root / WORKSPACE_DIR_NAME


def workspace_exists(root: Path) -> bool:
    return (workspace_dir(root) / "meta.json").is_file()


def clear_workspace(root: Path) -> None:
    path = workspace_dir(root)
    if path.exists():
        shutil.rmtree(path)


def clear_export_artifacts(root: Path) -> int:
    """Видаляє згенеровані CSV/XLSX лідів, компаній, activities у корені проєкту."""
    removed = 0
    for pattern in EXPORT_ARTIFACT_GLOBS:
        for path in root.glob(pattern):
            if path.is_file():
                path.unlink(missing_ok=True)
                removed += 1
    return removed


def clear_organizations_crm(root: Path) -> None:
    path = organizations_crm_path(root)
    if path.is_file():
        path.unlink()


def persons_crm_path(root: Path) -> Path:
    return workspace_dir(root) / PERSONS_CRM_FILENAME


def save_persons_crm(root: Path, df: pd.DataFrame) -> Path:
    path = persons_crm_path(root)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(path, index=False, encoding="utf-8-sig")
    except OSError as exc:
        log_disk_failure(path, exc, action="write persons CRM")
        raise
    log_disk_write(path, action="write persons CRM")
    return path


def persons_crm_exists(root: Path) -> bool:
    return persons_crm_path(root).is_file()


def clear_persons_crm(root: Path) -> None:
    path = persons_crm_path(root)
    if path.is_file():
        path.unlink()


def _df_path(base: Path, key: str) -> Path:
    return base / f"{key}.csv"


def _read_df(path: Path) -> pd.DataFrame | None:
    if not path.is_file():
        return None
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    df.columns = [str(c).strip().lstrip("\ufeff") for c in df.columns]
    return df


def _write_df(path: Path, df: pd.DataFrame | None) -> None:
    if df is None or df.empty:
        if path.is_file():
            path.unlink()
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(path, index=False, encoding="utf-8-sig")
    except OSError as exc:
        log_disk_failure(path, exc, action="write workspace frame")
        raise


def _session_value(session: Any, key: str, default: Any = None) -> Any:
    if hasattr(session, "get") and callable(session.get):
        return session.get(key, default)
    return getattr(session, key, default)


def _workspace_meta(session: Any) -> dict[str, Any]:
    meta: dict[str, Any] = {"version": WORKSPACE_VERSION}
    for key in _META_KEYS:
        value = _session_value(session, key)
        if key == "pending_only":
            meta[key] = bool(value) if value is not None else True
        elif key == "load_error":
            meta[key] = str(value) if value else None
        elif value is not None:
            meta[key] = value

    loaded_at = _session_value(session, "loaded_at")
    if isinstance(loaded_at, datetime):
        meta["loaded_at"] = loaded_at.astimezone(timezone.utc).isoformat()
    elif loaded_at:
        meta["loaded_at"] = str(loaded_at)
    return meta


def _write_meta_json(base: Path, session: Any) -> None:
    meta_path = base / "meta.json"
    try:
        meta_path.write_text(
            json.dumps(_workspace_meta(session), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError as exc:
        log_disk_failure(meta_path, exc, action="write workspace meta")
        raise


def save_workspace_meta(root: Path, session: Any) -> None:
    """Лише meta.json (вкладка, імена файлів) — без великих CSV."""
    base = workspace_dir(root)
    base.mkdir(parents=True, exist_ok=True)
    _write_meta_json(base, session)


def save_workspace_frames(root: Path, session: Any, keys: tuple[str, ...]) -> None:
    """Зберігає лише обрані DataFrame + meta (швидше після оновлення lookup)."""
    base = workspace_dir(root)
    base.mkdir(parents=True, exist_ok=True)
    _write_meta_json(base, session)
    for key in keys:
        if key in _DATAFRAME_KEYS:
            _write_df(_df_path(base, key), _session_value(session, key))


def save_workspace(root: Path, session: Any) -> None:
    """Зберігає session_state на диск."""
    base = workspace_dir(root)
    base.mkdir(parents=True, exist_ok=True)
    _write_meta_json(base, session)

    for key in _DATAFRAME_KEYS:
        _write_df(_df_path(base, key), _session_value(session, key))


def load_workspace(root: Path) -> dict[str, Any] | None:
    """Читає збережений стан; None якщо немає workspace."""
    base = workspace_dir(root)
    meta_path = base / "meta.json"
    if not meta_path.is_file():
        return None

    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None

    state: dict[str, Any] = {}
    for key in _META_KEYS:
        if key in meta:
            state[key] = meta[key]

    loaded_raw = meta.get("loaded_at")
    if loaded_raw:
        try:
            state["loaded_at"] = datetime.fromisoformat(str(loaded_raw))
            if state["loaded_at"].tzinfo is None:
                state["loaded_at"] = state["loaded_at"].replace(tzinfo=timezone.utc)
        except ValueError:
            state["loaded_at"] = None

    for key in _DATAFRAME_KEYS:
        df = _read_df(_df_path(base, key))
        if df is not None:
            state[key] = df

    return state


def apply_workspace(state: dict[str, Any], session: Any) -> None:
    for key, value in state.items():
        session[key] = value


def has_persisted_data(session: Any) -> bool:
    return any(
        (df := _session_value(session, key)) is not None
        and (not hasattr(df, "empty") or not df.empty)
        for key in _DATAFRAME_KEYS
    )
