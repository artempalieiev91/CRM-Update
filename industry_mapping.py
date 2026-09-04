"""Відповідність назв індустрій з джерела → значення для CRM (колонка Industry)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

INDUSTRY_COLUMN = "Industry"
DEFAULT_MAPPING_PATH = Path(__file__).resolve().parent / "industry_mapping.csv"


def load_industry_mapping(path: Path | str = DEFAULT_MAPPING_PATH) -> dict[str, str]:
    """Завантажує CSV: стовпець 1 — джерело, стовпець 2 — CRM. Пізніші рядки перекривають ранні."""
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(f"Файл мапінгу індустрій не знайдено: {p}")

    df = pd.read_csv(
        p,
        header=None,
        names=["source", "crm"],
        dtype=str,
        keep_default_na=False,
        encoding="utf-8-sig",
    )
    mapping: dict[str, str] = {}
    for source, crm in zip(df["source"], df["crm"]):
        key = str(source).strip()
        if not key:
            continue
        mapping[key] = str(crm).strip()
    return mapping


def _build_lookups(mapping: dict[str, str]) -> tuple[dict[str, str], dict[str, str]]:
    exact = dict(mapping)
    casefold = {k.casefold(): v for k, v in mapping.items()}
    return exact, casefold


_CACHED: tuple[dict[str, str], dict[str, str]] | None = None


def _get_lookups(
    mapping: dict[str, str] | None = None,
    *,
    mapping_path: Path | str | None = None,
) -> tuple[dict[str, str], dict[str, str]]:
    global _CACHED
    if mapping is not None:
        return _build_lookups(mapping)
    if mapping_path is not None:
        return _build_lookups(load_industry_mapping(mapping_path))
    if _CACHED is None:
        _CACHED = _build_lookups(load_industry_mapping())
    return _CACHED


def normalize_industry_value(
    value: object,
    *,
    mapping: dict[str, str] | None = None,
    mapping_path: Path | str | None = None,
) -> str:
    """Джерелова назва → CRM; без збігу — без змін."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    exact, casefold = _get_lookups(mapping, mapping_path=mapping_path)
    if text in exact:
        return exact[text]
    mapped = casefold.get(text.casefold())
    return mapped if mapped is not None else text


def normalize_industry_column(
    df: pd.DataFrame,
    column: str = INDUSTRY_COLUMN,
    *,
    mapping: dict[str, str] | None = None,
    mapping_path: Path | str | None = None,
) -> pd.DataFrame:
    out = df.copy()
    if column not in out.columns:
        return out
    fn = lambda v: normalize_industry_value(
        v, mapping=mapping, mapping_path=mapping_path
    )
    out[column] = out[column].map(fn)
    return out
