"""Завантаження Google Таблиці та фільтрація рядків для CRM."""

from __future__ import annotations

import io
import re
from pathlib import Path

import pandas as pd
import requests

from ids_runtime import log_key_action

DEFAULT_SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1PykN0zXd5xO51R2x-kW1_a_U-quNrQmOG7ESE3niYt4/edit?gid=0#gid=0"
)
PPL_CRM_TEMPLATE_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1VmDcr8_qYJYDVa-6qdN_XFk1E6_II2aveIyLKjGqedg/edit?gid=0#gid=0"
)
CRM_STATUS_COLUMN = "CRM status"
SOURCE_SHEET_COLUMN = "Source sheet"
CRM_STATUS_FILTER_PENDING = "pending"
CRM_STATUS_FILTER_VERIFICATION = "verification"
CRM_STATUS_FILTER_PENDING_AND_VERIFICATION = "pending+verification"
CRM_STATUS_FILTER_ALL = "all"
DEFAULT_OUTPUT_PATH = Path(__file__).resolve().parent / "pending_crm_upload.csv"


def extract_spreadsheet_id(url_or_id: str) -> str:
    s = (url_or_id or "").strip()
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", s)
    if m:
        return m.group(1)
    if re.fullmatch(r"[a-zA-Z0-9-_]+", s):
        return s
    raise ValueError("Не вдалося визначити ID таблиці з посилання або рядка.")


def extract_gid(url_or_id: str, *, default: str = "0") -> str:
    s = (url_or_id or "").strip()
    m = re.search(r"[?&#]gid=(\d+)", s)
    return m.group(1) if m else default


_SPREADSHEET_URL_RE = re.compile(
    r"https?://docs\.google\.com/spreadsheets/d/[a-zA-Z0-9-_]+[^\s]*",
    re.IGNORECASE,
)


def parse_spreadsheet_urls(text: str) -> list[str]:
    """Одне або кілька посилань Google Таблиці (кожне з нового рядка або через пробіл)."""
    raw = (text or "").strip()
    if not raw:
        return []

    found = _SPREADSHEET_URL_RE.findall(raw)
    candidates = found if found else [line.strip() for line in raw.splitlines() if line.strip()]

    out: list[str] = []
    seen_ids: set[str] = set()
    for item in candidates:
        url = str(item).strip()
        if not url:
            continue
        try:
            sid = extract_spreadsheet_id(url)
        except ValueError:
            continue
        if sid in seen_ids:
            continue
        seen_ids.add(sid)
        out.append(url)
    return out


def google_sheet_csv_export_url(spreadsheet_url_or_id: str, gid: str | None = None) -> str:
    sid = extract_spreadsheet_id(spreadsheet_url_or_id)
    base = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv"
    if gid:
        base += f"&gid={gid}"
    return base


def _parse_csv_text(text: str) -> pd.DataFrame:
    sample = text[:8192]
    sep = ","
    if sample.count(";") > sample.count(",") and ";" in sample.split("\n", 1)[0]:
        sep = ";"
    df = pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False, sep=sep)
    if len(df.columns) == 1 and sep == "," and ";" in str(df.columns[0]):
        df = pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False, sep=";")
    return df.replace({"": pd.NA})


def load_source_csv(
    source: str | bytes | Path | io.BytesIO,
    *,
    encoding: str = "utf-8-sig",
) -> pd.DataFrame:
    """CSV з файлу, шляху або bytes (той самий формат, що експорт Google Таблиці)."""
    if isinstance(source, Path):
        text = source.read_text(encoding=encoding)
    elif isinstance(source, bytes):
        for enc in (encoding, "utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return _parse_csv_text(source.decode(enc))
            except UnicodeDecodeError:
                continue
        text = source.decode(encoding, errors="replace")
    elif isinstance(source, io.BytesIO):
        raw = source.getvalue()
        return load_source_csv(raw, encoding=encoding)
    else:
        text = str(source)
    return _parse_csv_text(text)


def load_sheet_csv(
    spreadsheet_url_or_id: str,
    *,
    gid: str | None = None,
    timeout: int = 120,
) -> pd.DataFrame:
    if gid is None:
        gid = extract_gid(spreadsheet_url_or_id)
    url = google_sheet_csv_export_url(spreadsheet_url_or_id, gid)
    r = requests.get(url, timeout=timeout)
    r.raise_for_status()
    r.encoding = "utf-8-sig"
    return _parse_csv_text(r.text)


def _resolve_status_column(df: pd.DataFrame) -> str:
    for col in df.columns:
        if str(col).strip().lower() == CRM_STATUS_COLUMN.lower():
            return col
    raise ValueError(f"У таблиці немає колонки «{CRM_STATUS_COLUMN}».")


def has_crm_status_column(df: pd.DataFrame) -> bool:
    for col in df.columns:
        if str(col).strip().lower() == CRM_STATUS_COLUMN.lower():
            return True
    return False


_LEAD_DATA_COLUMNS = ("Company", "First name", "Last name", "Email", "Linkedin Person")


def _has_lead_data(df: pd.DataFrame) -> "pd.Series[bool]":
    """True якщо хоч одне з ключових полів ліда не порожнє."""
    cols = [c for c in _LEAD_DATA_COLUMNS if c in df.columns]
    if not cols:
        return pd.Series(True, index=df.index)
    result = pd.Series(False, index=df.index)
    for col in cols:
        filled = df[col].notna() & (df[col].astype(str).str.strip().str.casefold() != "<na>") & (df[col].astype(str).str.strip() != "")
        result = result | filled
    return result


def filter_pending_crm_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    status_col = _resolve_status_column(df)
    mask = (df[status_col].isna() | (df[status_col].astype(str).str.strip() == "")) & _has_lead_data(df)
    pending = df.loc[mask].copy()
    pending[status_col] = ""
    return pending, status_col


def filter_pending_and_verification_rows(df: pd.DataFrame) -> pd.DataFrame:
    pending, _status_col = filter_pending_crm_rows(df)
    verification = filter_rows_by_crm_status(df, "verification")
    if pending.empty:
        return verification
    if verification.empty:
        return pending
    return pd.concat([pending, verification], ignore_index=True)


def select_rows_for_export(
    df: pd.DataFrame,
    *,
    crm_status_filter: str = CRM_STATUS_FILTER_PENDING,
    pending_only: bool | None = None,
) -> pd.DataFrame:
    """
    Фільтр лідів за CRM status.

    - pending — порожній CRM status
    - verification — значення «verification»
    - pending+verification — порожній або verification
    - all — усі рядки
    """
    if pending_only is not None:
        crm_status_filter = (
            CRM_STATUS_FILTER_PENDING if pending_only else CRM_STATUS_FILTER_ALL
        )

    key = str(crm_status_filter).strip().casefold()
    if key == CRM_STATUS_FILTER_PENDING:
        pending, _status_col = filter_pending_crm_rows(df)
        return pending
    if key == CRM_STATUS_FILTER_VERIFICATION:
        return filter_rows_by_crm_status(df, "verification")
    if key == CRM_STATUS_FILTER_PENDING_AND_VERIFICATION:
        return filter_pending_and_verification_rows(df)
    if key == CRM_STATUS_FILTER_ALL:
        return df.copy()
    raise ValueError(
        f"Невідомий фільтр CRM status: {crm_status_filter!r}. "
        f"Очікується: pending, verification, pending+verification або all."
    )


def crm_status_filter_label(crm_status_filter: str) -> str:
    labels = {
        CRM_STATUS_FILTER_PENDING: "порожній CRM status (pending)",
        CRM_STATUS_FILTER_VERIFICATION: "CRM status = verification",
        CRM_STATUS_FILTER_PENDING_AND_VERIFICATION: "порожній + verification",
        CRM_STATUS_FILTER_ALL: "усі рядки",
    }
    return labels.get(str(crm_status_filter).strip().casefold(), crm_status_filter)


def save_pending_csv(df: pd.DataFrame, path: Path | str = DEFAULT_OUTPUT_PATH) -> Path:
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(out, index=False, encoding="utf-8-sig")
    return out


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def google_sheet_gviz_csv_url(spreadsheet_url_or_id: str, *, sheet_name: str) -> str:
    sid = extract_spreadsheet_id(spreadsheet_url_or_id)
    from urllib.parse import quote

    return (
        f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq"
        f"?tqx=out:csv&sheet={quote(sheet_name)}"
    )


def discover_all_sheet_names(
    spreadsheet_url_or_id: str,
    *,
    timeout: int = 120,
) -> list[str]:
    """Усі листи таблиці (публічний export xlsx)."""
    sid = extract_spreadsheet_id(spreadsheet_url_or_id)
    url = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx"
    r = requests.get(url, timeout=timeout)
    r.raise_for_status()
    xl = pd.ExcelFile(io.BytesIO(r.content))
    return list(xl.sheet_names)


def discover_sheet_tab_names(
    spreadsheet_url_or_id: str,
    *,
    timeout: int = 120,
) -> list[str]:
    """Імена листів дослідження (regex у HTML або всі листи з xlsx)."""
    sid = extract_spreadsheet_id(spreadsheet_url_or_id)
    url = f"https://docs.google.com/spreadsheets/d/{sid}/edit"
    r = requests.get(url, timeout=timeout)
    r.raise_for_status()
    names = sorted(
        set(
            re.findall(
                r"0[56]\.26(?:GEN|DE|AI|SS|EVENT|SALES|NON)[A-Z0-9\-]*",
                r.text,
            )
        )
    )
    if names:
        return names
    return discover_all_sheet_names(spreadsheet_url_or_id, timeout=timeout)


def load_sheet_by_name(
    spreadsheet_url_or_id: str,
    sheet_name: str,
    *,
    timeout: int = 120,
) -> pd.DataFrame:
    url = google_sheet_gviz_csv_url(spreadsheet_url_or_id, sheet_name=sheet_name)
    r = requests.get(url, timeout=timeout)
    r.raise_for_status()
    r.encoding = "utf-8-sig"
    df = pd.read_csv(io.StringIO(r.text), dtype=str, keep_default_na=False)
    return df.replace({"": pd.NA})


def load_all_sheets(
    spreadsheet_url_or_id: str,
    *,
    timeout: int = 120,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """
    Завантажити та обʼєднати всі листи таблиці.
    Додає колонку Source sheet; порожні листи пропускає.
    """
    names = discover_all_sheet_names(spreadsheet_url_or_id, timeout=timeout)
    if not names:
        raise ValueError("У таблиці немає листів.")

    parts: list[pd.DataFrame] = []
    counts: dict[str, int] = {}
    errors: list[str] = []

    for name in names:
        try:
            sheet_df = load_sheet_by_name(
                spreadsheet_url_or_id, name, timeout=timeout
            )
        except requests.RequestException as exc:
            errors.append(f"{name}: {exc}")
            counts[name] = 0
            continue

        sheet_df = _drop_empty_columns(sheet_df)
        if sheet_df.empty or len(sheet_df.columns) == 0:
            counts[name] = 0
            continue

        chunk = sheet_df.copy()
        if SOURCE_SHEET_COLUMN not in chunk.columns:
            chunk.insert(0, SOURCE_SHEET_COLUMN, name)
        else:
            chunk[SOURCE_SHEET_COLUMN] = chunk[SOURCE_SHEET_COLUMN].fillna(name)
            chunk.loc[
                chunk[SOURCE_SHEET_COLUMN].astype(str).str.strip() == "",
                SOURCE_SHEET_COLUMN,
            ] = name

        parts.append(chunk)
        counts[name] = len(chunk)

    if not parts:
        detail = "; ".join(errors[:3]) if errors else "усі листи порожні"
        raise ValueError(f"Не вдалося завантажити дані з листів таблиці ({detail}).")

    combined = pd.concat(parts, ignore_index=True, sort=False)
    return combined, counts


def load_all_sheets_from_urls(
    spreadsheet_urls: list[str],
    *,
    timeout: int = 120,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Завантажити та обʼєднати всі листи з однієї або кількох Google Таблиць."""
    urls = [str(url or "").strip() for url in spreadsheet_urls if str(url or "").strip()]
    if not urls:
        raise ValueError("Потрібно хоча б одне посилання на Google Таблицю.")

    log_key_action(
        "Google Sheets fetch start",
        ok=True,
        detail=f"{len(urls)} URL",
        update_marker=False,
    )
    try:
        parts: list[pd.DataFrame] = []
        counts: dict[str, int] = {}
        multi = len(urls) > 1

        for index, url in enumerate(urls, start=1):
            sheet_df, sheet_counts = load_all_sheets(url, timeout=timeout)
            if multi and SOURCE_SHEET_COLUMN in sheet_df.columns:
                prefix = f"Таблиця {index}"
                sheet_df = sheet_df.copy()
                sheet_df[SOURCE_SHEET_COLUMN] = (
                    prefix + " · " + sheet_df[SOURCE_SHEET_COLUMN].astype(str)
                )
                for name, row_count in sheet_counts.items():
                    counts[f"{prefix} · {name}"] = row_count
            else:
                counts.update(sheet_counts)
            parts.append(sheet_df)

        combined = pd.concat(parts, ignore_index=True, sort=False)
    except Exception as exc:
        log_key_action("Google Sheets fetch", ok=False, detail=str(exc))
        raise
    log_key_action(
        "Google Sheets fetch",
        ok=True,
        detail=f"{len(urls)} tables, {len(combined)} rows, {len(counts)} sheets",
    )
    return combined, counts


def _drop_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    drop: list[str] = []
    for col in out.columns:
        name = str(col).strip()
        if name.startswith("Unnamed:") or not name:
            drop.append(col)
            continue
        # CRM status часто повністю порожній (pending) — колонку не прибираємо.
        if name.casefold() == CRM_STATUS_COLUMN.casefold():
            continue
        if out[col].isna().all() or (out[col].astype(str).str.strip() == "").all():
            drop.append(col)
    return out.drop(columns=drop, errors="ignore")


def filter_rows_by_crm_status(
    df: pd.DataFrame,
    status_value: str,
    *,
    column: str = CRM_STATUS_COLUMN,
) -> pd.DataFrame:
    status_col = _resolve_status_column(df)
    target = status_value.strip().casefold()
    mask = df[status_col].astype(str).str.strip().str.casefold() == target
    return df.loc[mask].copy()


def export_verification_rows_all_sheets(
    spreadsheet_url_or_id: str,
    output_path: Path | str,
    *,
    status_value: str = "verification",
    sheet_names: list[str] | None = None,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Усі листи → рядки з CRM status = verification (+ колонка Source sheet)."""
    names = sheet_names or discover_sheet_tab_names(spreadsheet_url_or_id)
    parts: list[pd.DataFrame] = []
    counts: dict[str, int] = {}
    for name in names:
        sheet_df = load_sheet_by_name(spreadsheet_url_or_id, name)
        filtered = filter_rows_by_crm_status(sheet_df, status_value)
        counts[name] = len(filtered)
        if len(filtered):
            chunk = _drop_empty_columns(filtered)
            chunk.insert(0, SOURCE_SHEET_COLUMN, name)
            parts.append(chunk)

    if parts:
        combined = pd.concat(parts, ignore_index=True)
    else:
        combined = pd.DataFrame()

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(out, index=False, encoding="utf-8-sig")
    return combined, counts
