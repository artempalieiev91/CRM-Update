"""Запис ID у Excel як числа (коректне сортування, без апострофа при імпорті в CRM)."""

from __future__ import annotations

import pandas as pd

EXCEL_NUMERIC_ID_COLUMNS = frozenset(
    {
        "Activity - ID",
        "Person - ID",
        "Organization - ID",
        "Website ID",
        "Linkedin ID",
        "Person - ID by Email",
        "ID by Linkedin Person",
    }
)


def excel_cell_value(col_name: str, value: object):
    """ID-колонки — int/float у Excel; решта — текст."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip().replace("<NA>", "")
    if not text or text.lower() in {"nan", "none"}:
        return None
    if text.startswith("'"):
        text = text[1:].strip()
    if not text:
        return None
    if col_name in EXCEL_NUMERIC_ID_COLUMNS:
        try:
            number = float(text)
            if number.is_integer():
                return int(number)
            return number
        except ValueError:
            return text
    return text
