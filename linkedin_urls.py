"""Нормалізація LinkedIn Person → повний URL https://www.linkedin.com/..."""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from ids_runtime import log_disk_failure, log_disk_write

LINKEDIN_PERSON_COLUMN = "Linkedin Person"
LINKEDIN_HELPER_ID_COLUMN = "Person - Linked Helper Id"
LINKEDIN_BASE = "https://www.linkedin.com/"
ICP_NAME_EXPORT_COLUMN = "ICP Name"
ICP_NAME_SOURCE_COLUMN = "ICP name"
ACTIVITY_SUBJECT_COLUMN = "Activity Subject"
LINKEDIN_EXPORT_COLUMNS: tuple[str, ...] = (
    ICP_NAME_EXPORT_COLUMN,
    LINKEDIN_PERSON_COLUMN,
)

_STRIP_PREFIX = re.compile(r"^https?://", re.IGNORECASE)
_STRIP_WWW = re.compile(r"^www\.", re.IGNORECASE)


def normalize_linkedin_person_url(value: object) -> str:
    """Єдиний ключ для злиття: CRM «linkedin.com/in/…» і тулза «http(s)://www.linkedin.com/in/…»."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = text.split("?")[0].split("#")[0].strip().rstrip("/")
    text = _STRIP_PREFIX.sub("", text)
    text = _STRIP_WWW.sub("", text)
    lower = text.casefold()
    if lower.startswith("in/"):
        text = f"linkedin.com/{text.lstrip('/')}"
    elif lower.startswith("/in/"):
        text = f"linkedin.com{lower}"
    if not text.casefold().startswith("linkedin.com/"):
        return ""
    path = text[len("linkedin.com/") :].lstrip("/").rstrip("/").casefold()
    if not path:
        return ""
    return f"{LINKEDIN_BASE}{path}"


def linkedin_person_match_key(value: object) -> str:
    """Канонічний ключ злиття (шлях профілю в нижньому регістрі)."""
    url = normalize_linkedin_person_url(value)
    if not url:
        return ""
    return url[len(LINKEDIN_BASE) :].casefold()


def linkedin_helper_id_from_person(value: object) -> str:
    """
    Як PPL на Google: REGEXREPLACE(Linkedin Person, «linkedin.com/in/», «»).
    Slug профілю після /in/ (без домену та префікса).
    """
    key = linkedin_person_match_key(value)
    if not key.startswith("in/"):
        return ""
    slug = key[3:].split("/")[0].strip()
    return slug


def format_linkedin_url_for_tool(url: str) -> str:
    """Формат для тулзи: https://www.linkedin.com/in/slug/"""
    if not url:
        return ""
    return url if url.endswith("/") else f"{url}/"


def linkedin_person_urls(
    df: pd.DataFrame,
    *,
    column: str = LINKEDIN_PERSON_COLUMN,
    trailing_slash: bool = False,
) -> list[str]:
    if column not in df.columns:
        return []
    urls: list[str] = []
    for raw in df[column]:
        url = normalize_linkedin_person_url(raw)
        if url:
            if trailing_slash:
                url = format_linkedin_url_for_tool(url)
            urls.append(url)
    return urls


def _icp_name_source_column(df: pd.DataFrame) -> str | None:
    if ICP_NAME_SOURCE_COLUMN in df.columns:
        return ICP_NAME_SOURCE_COLUMN
    if ACTIVITY_SUBJECT_COLUMN in df.columns:
        return ACTIVITY_SUBJECT_COLUMN
    return None


def build_linkedin_export(df: pd.DataFrame, *, column: str = LINKEDIN_PERSON_COLUMN) -> pd.DataFrame:
    if column not in df.columns:
        return pd.DataFrame(columns=list(LINKEDIN_EXPORT_COLUMNS))

    icp_col = _icp_name_source_column(df)
    rows: list[dict[str, str]] = []
    for idx in range(len(df)):
        url = normalize_linkedin_person_url(df[column].iloc[idx])
        if not url:
            continue
        icp = ""
        if icp_col is not None:
            icp = str(df[icp_col].iloc[idx]).strip()
        rows.append(
            {
                ICP_NAME_EXPORT_COLUMN: icp,
                LINKEDIN_PERSON_COLUMN: format_linkedin_url_for_tool(url),
            }
        )
    return pd.DataFrame(rows, columns=list(LINKEDIN_EXPORT_COLUMNS))


def linkedin_export_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return build_linkedin_export(df).to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def save_linkedin_export(df: pd.DataFrame, path: Path | str) -> Path:
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        build_linkedin_export(df).to_csv(out, index=False, encoding="utf-8-sig")
    except OSError as exc:
        log_disk_failure(out, exc, action="write linkedin CSV")
        raise
    log_disk_write(out, action="write linkedin CSV")
    return out
