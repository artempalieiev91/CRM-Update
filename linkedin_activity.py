"""Результати тулзи LinkedIn activity → колонка Person - Linkedin Active у CRM."""

from __future__ import annotations

import io
import re
from pathlib import Path

import pandas as pd

from crm_export import CRM_COMP_COLUMNS
from linkedin_urls import (
    LINKEDIN_PERSON_COLUMN,
    linkedin_person_match_key,
    normalize_linkedin_person_url,
)

PERSON_LINKEDIN_ACTIVE_COLUMN = "Person - Linkedin Active"
ACTIVITY_URL_COLUMNS = ("url", "linkedin person", "linkedin person url", "linkedin url")
ACTIVITY_VALUE_COLUMNS = ("recent activity", "activity", "recent_activity")

_ACTIVITY_RE = re.compile(
    r"^\s*(\d+)\s*(mo|yr|y|wk|w|d|day|days|h|hr|hours?|min|minutes?|m)?\s*$",
    re.IGNORECASE,
)
_DAYS_PER_MONTH = 30


def parse_recent_activity_days(value: object) -> float | None:
    """Повертає тривалість у днях або None, якщо розпарсити не вдалося."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    if not text:
        return None
    m = _ACTIVITY_RE.match(text)
    if not m:
        return None
    amount = int(m.group(1))
    unit = (m.group(2) or "d").casefold()
    if unit == "mo":
        return float(amount * _DAYS_PER_MONTH)
    if unit in ("yr", "y"):
        return float(amount * 365)
    if unit in ("wk", "w"):
        return float(amount * 7)
    if unit in ("d", "day", "days"):
        return float(amount)
    if unit in ("h", "hr", "hour", "hours"):
        return amount / 24.0
    if unit in ("min", "minute", "minutes", "m"):
        return amount / (24.0 * 60.0)
    return float(amount)


def recent_activity_to_active_flag(value: object) -> str:
    """< 1 місяця → yes; >= 1 місяця → no; порожнє/невідоме → \"\"."""
    days = parse_recent_activity_days(value)
    if days is None:
        return ""
    if days < _DAYS_PER_MONTH:
        return "yes"
    return "no"


def _resolve_column(df: pd.DataFrame, candidates: tuple[str, ...]) -> str | None:
    lower_map = {str(c).strip().casefold(): c for c in df.columns}
    for name in candidates:
        if name in lower_map:
            return lower_map[name]
    return None


def normalize_activity_results(df: pd.DataFrame) -> pd.DataFrame:
    url_col = _resolve_column(df, ACTIVITY_URL_COLUMNS)
    activity_col = _resolve_column(df, ACTIVITY_VALUE_COLUMNS)
    if url_col is None:
        raise ValueError(
            "У файлі результатів немає колонки з URL (очікується «Url» або «Linkedin Person»)."
        )
    if activity_col is None:
        raise ValueError(
            "У файлі результатів немає колонки активності (очікується «Recent Activity»)."
        )
    out = pd.DataFrame(
        {
            "url": df[url_col].map(normalize_linkedin_person_url),
            "activity": df[activity_col],
        }
    )
    return out.loc[out["url"] != ""].copy()


def load_activity_results_csv(
    source: str | bytes | Path | io.BytesIO,
    *,
    encoding: str = "utf-8-sig",
) -> pd.DataFrame:
    if isinstance(source, Path):
        raw = pd.read_csv(source, dtype=str, keep_default_na=False, encoding=encoding)
    elif isinstance(source, bytes):
        raw = pd.read_csv(io.BytesIO(source), dtype=str, keep_default_na=False, encoding=encoding)
    else:
        raw = pd.read_csv(io.StringIO(str(source)), dtype=str, keep_default_na=False)
    return normalize_activity_results(raw)


def build_activity_lookup(activity_df: pd.DataFrame) -> dict[str, str]:
    """Ключ профілю (in/slug) → yes / no; URL з тулзи з http/www зводиться до того ж ключа, що CRM."""
    normalized = normalize_activity_results(activity_df)
    lookup: dict[str, str] = {}
    for url, activity in zip(normalized["url"], normalized["activity"]):
        key = linkedin_person_match_key(url)
        if not key:
            continue
        flag = recent_activity_to_active_flag(activity)
        lookup[key] = flag if flag else "no"
    return lookup


def _resolve_linkedin_column(
    df: pd.DataFrame,
    candidates: tuple[str, ...],
) -> str | None:
    lower_map = {str(c).strip().casefold(): c for c in df.columns}
    for name in candidates:
        key = name.casefold()
        if key in lower_map:
            return lower_map[key]
    return None


def merge_linkedin_active_into_df(
    df: pd.DataFrame,
    activity_source: pd.DataFrame | str | bytes | Path | io.BytesIO,
    *,
    linkedin_columns: tuple[str, ...] = (
        LINKEDIN_PERSON_COLUMN,
        "Person - LinkedIn",
    ),
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Person - Linkedin Active за URL профілю (Leads comp або Activities)."""
    if isinstance(activity_source, pd.DataFrame):
        activity_df = normalize_activity_results(activity_source)
    else:
        activity_df = load_activity_results_csv(activity_source)

    linkedin_col = _resolve_linkedin_column(df, linkedin_columns)
    if linkedin_col is None:
        raise ValueError(
            f"У таблиці немає колонки LinkedIn (очікується одна з: {', '.join(linkedin_columns)})."
        )

    lookup = build_activity_lookup(activity_df)
    out = df.copy()
    if PERSON_LINKEDIN_ACTIVE_COLUMN not in out.columns:
        out[PERSON_LINKEDIN_ACTIVE_COLUMN] = ""
    flags: list[str] = []
    yes_count = 0
    no_from_activity = 0
    not_in_activity_file = 0
    empty_linkedin = 0
    for i, raw_url in enumerate(out[linkedin_col]):
        existing = str(out.iloc[i][PERSON_LINKEDIN_ACTIVE_COLUMN]).strip()
        key = linkedin_person_match_key(raw_url)
        if not key:
            flags.append(existing)
            empty_linkedin += 1
            continue
        if key in lookup:
            flag = lookup[key]
            flags.append(flag)
            if flag == "yes":
                yes_count += 1
            elif flag == "no":
                no_from_activity += 1
        else:
            flags.append("")
            not_in_activity_file += 1

    out[PERSON_LINKEDIN_ACTIVE_COLUMN] = flags
    stats = {
        "activity_rows": len(activity_df),
        "lookup_keys": len(lookup),
        "comp_rows": len(out),
        "yes": yes_count,
        "no_from_activity": no_from_activity,
        "not_in_activity_file": not_in_activity_file,
        "no_missing_result": not_in_activity_file,
        "empty_linkedin": empty_linkedin,
    }
    return out, stats


def merge_linkedin_active_into_comp(
    comp_df: pd.DataFrame,
    activity_source: pd.DataFrame | str | bytes | Path | io.BytesIO,
    *,
    linkedin_column: str = LINKEDIN_PERSON_COLUMN,
) -> tuple[pd.DataFrame, dict[str, int]]:
    out, stats = merge_linkedin_active_into_df(
        comp_df,
        activity_source,
        linkedin_columns=(linkedin_column, LINKEDIN_PERSON_COLUMN, "Person - LinkedIn"),
    )
    ordered = [c for c in CRM_COMP_COLUMNS if c in out.columns]
    extra = [c for c in out.columns if c not in ordered]
    return out[ordered + extra], stats
