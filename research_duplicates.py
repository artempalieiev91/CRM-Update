"""Дублікати Email та Linkedin Person у завантаженому дослідженні."""

from __future__ import annotations

import pandas as pd

from linkedin_urls import LINKEDIN_PERSON_COLUMN, linkedin_person_match_key
from mx_checker import EMAIL_COLUMN
from person_lookup import normalize_email_key

RESEARCH_DUPLICATE_EMAIL_COLUMN = "Duplicate Email"
RESEARCH_DUPLICATE_LINKEDIN_COLUMN = "Duplicate Linkedin Person"


def research_linkedin_person_key(value: object) -> str:
    """Ключ лише для профілів /in/ (не company URL)."""
    key = linkedin_person_match_key(value)
    if key.startswith("in/"):
        return key
    return ""


def duplicate_research_email_keys(df: pd.DataFrame) -> frozenset[str]:
    """Email, що зустрічаються у 2+ рядках (нормалізований ключ)."""
    if EMAIL_COLUMN not in df.columns:
        return frozenset()
    keys = df[EMAIL_COLUMN].map(normalize_email_key)
    keys = keys.loc[keys != ""]
    if keys.empty:
        return frozenset()
    counts = keys.value_counts()
    return frozenset(k for k, n in counts.items() if n > 1)


def duplicate_research_linkedin_keys(df: pd.DataFrame) -> frozenset[str]:
    """Linkedin Person (/in/), що зустрічаються у 2+ рядках."""
    if LINKEDIN_PERSON_COLUMN not in df.columns:
        return frozenset()
    keys = df[LINKEDIN_PERSON_COLUMN].map(research_linkedin_person_key)
    keys = keys.loc[keys != ""]
    if keys.empty:
        return frozenset()
    counts = keys.value_counts()
    return frozenset(k for k, n in counts.items() if n > 1)


def annotate_research_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Додає колонки Duplicate Email / Duplicate Linkedin Person (yes = дублікат)."""
    out = df.copy()
    email_dupes = duplicate_research_email_keys(out)
    li_dupes = duplicate_research_linkedin_keys(out)

    if EMAIL_COLUMN in out.columns:
        out[RESEARCH_DUPLICATE_EMAIL_COLUMN] = out[EMAIL_COLUMN].map(
            lambda v: "yes" if normalize_email_key(v) in email_dupes else ""
        )
    else:
        out[RESEARCH_DUPLICATE_EMAIL_COLUMN] = ""

    if LINKEDIN_PERSON_COLUMN in out.columns:
        out[RESEARCH_DUPLICATE_LINKEDIN_COLUMN] = out[LINKEDIN_PERSON_COLUMN].map(
            lambda v: "yes" if research_linkedin_person_key(v) in li_dupes else ""
        )
    else:
        out[RESEARCH_DUPLICATE_LINKEDIN_COLUMN] = ""

    return out


def research_duplicate_stats(df: pd.DataFrame) -> dict[str, int]:
    """Лічильники для UI після завантаження дослідження."""
    work = annotate_research_duplicates(df)
    total = len(work)
    email_keys = len(duplicate_research_email_keys(work))
    linkedin_keys = len(duplicate_research_linkedin_keys(work))
    email_rows = 0
    linkedin_rows = 0
    any_rows = 0
    if RESEARCH_DUPLICATE_EMAIL_COLUMN in work.columns:
        email_rows = int((work[RESEARCH_DUPLICATE_EMAIL_COLUMN] == "yes").sum())
    if RESEARCH_DUPLICATE_LINKEDIN_COLUMN in work.columns:
        linkedin_rows = int((work[RESEARCH_DUPLICATE_LINKEDIN_COLUMN] == "yes").sum())
    if email_rows or linkedin_rows:
        mask = pd.Series(False, index=work.index)
        if RESEARCH_DUPLICATE_EMAIL_COLUMN in work.columns:
            mask |= work[RESEARCH_DUPLICATE_EMAIL_COLUMN] == "yes"
        if RESEARCH_DUPLICATE_LINKEDIN_COLUMN in work.columns:
            mask |= work[RESEARCH_DUPLICATE_LINKEDIN_COLUMN] == "yes"
        any_rows = int(mask.sum())
    return {
        "total": total,
        "duplicate_email_keys": email_keys,
        "duplicate_linkedin_keys": linkedin_keys,
        "rows_with_duplicate_email": email_rows,
        "rows_with_duplicate_linkedin": linkedin_rows,
        "rows_with_any_duplicate": any_rows,
    }


def build_research_duplicates_report(df: pd.DataFrame) -> pd.DataFrame:
    """Зведення: значення → скільки рядків."""
    work = annotate_research_duplicates(df)
    rows: list[dict[str, object]] = []

    email_dupes = duplicate_research_email_keys(work)
    if email_dupes and EMAIL_COLUMN in work.columns:
        for key in sorted(email_dupes):
            group = work.loc[
                work[EMAIL_COLUMN].map(normalize_email_key) == key
            ]
            sample = str(group[EMAIL_COLUMN].iloc[0]).strip()
            rows.append(
                {
                    "Type": "Email",
                    "Value": sample,
                    "Count": len(group),
                    "Row numbers (1-based)": ", ".join(
                        str(i + 2) for i in group.index[:20]
                    )
                    + (" …" if len(group) > 20 else ""),
                }
            )

    li_dupes = duplicate_research_linkedin_keys(work)
    if li_dupes and LINKEDIN_PERSON_COLUMN in work.columns:
        for key in sorted(li_dupes):
            group = work.loc[
                work[LINKEDIN_PERSON_COLUMN].map(research_linkedin_person_key) == key
            ]
            sample = str(group[LINKEDIN_PERSON_COLUMN].iloc[0]).strip()
            rows.append(
                {
                    "Type": "Linkedin Person",
                    "Value": sample,
                    "Count": len(group),
                    "Row numbers (1-based)": ", ".join(
                        str(i + 2) for i in group.index[:20]
                    )
                    + (" …" if len(group) > 20 else ""),
                }
            )

    if not rows:
        return pd.DataFrame(
            columns=["Type", "Value", "Count", "Row numbers (1-based)"]
        )
    report = pd.DataFrame(rows)
    return report.sort_values(["Type", "Count"], ascending=[True, False]).reset_index(
        drop=True
    )


def research_duplicate_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Усі рядки дослідження, де Email або Linkedin Person — дублікат."""
    work = annotate_research_duplicates(df)
    mask = pd.Series(False, index=work.index)
    if RESEARCH_DUPLICATE_EMAIL_COLUMN in work.columns:
        mask |= work[RESEARCH_DUPLICATE_EMAIL_COLUMN] == "yes"
    if RESEARCH_DUPLICATE_LINKEDIN_COLUMN in work.columns:
        mask |= work[RESEARCH_DUPLICATE_LINKEDIN_COLUMN] == "yes"
    if not mask.any():
        return work.iloc[0:0].copy()
    return work.loc[mask].copy().reset_index(drop=True)
