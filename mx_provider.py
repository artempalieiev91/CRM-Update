"""Адаптер MX checker → колонка CRM Organization - Email provider."""

from __future__ import annotations

from typing import Callable

import pandas as pd

from mx_checker import (
    CRM_PROVIDER_COLUMN,
    EMAIL_COLUMN,
    auto_detect_column,
    fill_email_providers,
)

__all__ = [
    "EMAIL_COLUMN",
    "CRM_PROVIDER_COLUMN",
    "PROVIDER_COLUMN",
    "PROVIDER_OUTLOOK",
    "PROVIDER_OTHER",
    "auto_detect_column",
    "fill_email_providers",
]

PROVIDER_COLUMN = CRM_PROVIDER_COLUMN
PROVIDER_OUTLOOK = "outlook"
PROVIDER_OTHER = "Other (gmail, etc)"


def apply_mx_providers(
    df: pd.DataFrame,
    *,
    overwrite: bool = False,
    workers: int = 30,
    timeout: float = 5.0,
    on_progress: Callable[[int, int], None] | None = None,
) -> pd.DataFrame:
    return fill_email_providers(
        df,
        provider_column=CRM_PROVIDER_COLUMN,
        overwrite=overwrite,
        workers=workers,
        timeout=timeout,
        on_progress=on_progress,
    )
