"""Пошук Person - ID та CRM LinkedIn за експортом контактів з CRM."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from linkedin_urls import linkedin_person_match_key

PERSON_EMAIL_COLUMN = "Person - Email (in use)"
PERSON_LINKEDIN_COLUMN = "Person - LinkedIn (in use)"
PERSON_ID_COLUMN = "Person - ID"

PERSON_ID_BY_EMAIL_COLUMN = "Person - ID by Email"
ID_BY_LINKEDIN_PERSON_COLUMN = "ID by Linkedin Person"
CRM_BY_EMAIL_ID_COLUMN = "CRM by Email ID"
CRM_LINKEDIN_ID_COLUMN = "CRM Linkedin ID"
PERSON_LOOKUP_ISSUE_COLUMN = "Person lookup issue"
PERSON_HIGHLIGHT_COLUMN = "Підсвітити (CRM Email / LinkedIn ID)"

_EMAIL_COL_ALIASES = (
    PERSON_EMAIL_COLUMN,
    "Person - Email",
    "Person - Email - Work",
    "Person - Email - Home",
    "Person - Email - Other",
    "Email (in use)",
    "Email",
)
_LINKEDIN_COL_ALIASES = (
    PERSON_LINKEDIN_COLUMN,
    "Person - LinkedIn",
    "Person - Linkedin",
    "Person - Linkedin (in use)",
    "Linkedin Person",
    "LinkedIn URL",
)
_ID_COL_ALIASES = (
    PERSON_ID_COLUMN,
    "Person ID",
    "Person - Id",
    "pipedrive_contact_id",
    "Pipedrive contact id",
)


def _clean_header(name: object) -> str:
    return str(name).strip().lstrip("\ufeff").strip('"')


def _resolve_column(
    df: pd.DataFrame,
    aliases: tuple[str, ...],
    *,
    fuzzy_contains: tuple[str, ...] | None = None,
) -> str | None:
    cols = {_clean_header(c): c for c in df.columns}
    for name in aliases:
        if name in cols:
            return cols[name]
    lower = {_clean_header(c).casefold(): c for c in df.columns}
    for name in aliases:
        key = name.casefold()
        if key in lower:
            return lower[key]
    if fuzzy_contains:
        need = tuple(s.casefold() for s in fuzzy_contains)
        for key, original in lower.items():
            if all(part in key for part in need):
                return original
    return None


def person_email_columns(df: pd.DataFrame) -> list[str]:
    """Усі email-колонки persons (Work / Home / Other / in use)."""
    found: list[str] = []
    for col in df.columns:
        key = _clean_header(col).casefold()
        if ("person" in key and "email" in key) or key == "email":
            found.append(col)

    def _sort_key(name: str) -> tuple[int, str]:
        low = name.casefold()
        if "work" in low:
            return (0, low)
        if "home" in low:
            return (1, low)
        if "other" in low:
            return (2, low)
        if "in use" in low:
            return (3, low)
        return (4, low)

    return sorted(found, key=_sort_key)


def detect_person_csv_columns(df: pd.DataFrame) -> dict[str, str | None]:
    emails = person_email_columns(df)
    return {
        "email": emails[0] if emails else _resolve_column(
            df,
            _EMAIL_COL_ALIASES,
            fuzzy_contains=("person", "email"),
        ),
        "emails": emails,
        "linkedin": _resolve_column(
            df,
            _LINKEDIN_COL_ALIASES,
            fuzzy_contains=("person", "linkedin"),
        ),
        "id": _resolve_column(
            df,
            _ID_COL_ALIASES,
            fuzzy_contains=("person", "id"),
        ),
    }


def normalize_person_id(value: object) -> str:
    """Ключ Person - ID (18737.0 → 18737)."""
    s = str(value).strip()
    if not s or s.lower() in {"nan", "<na>"}:
        return ""
    try:
        num = float(s)
        if num == int(num):
            return str(int(num))
    except ValueError:
        pass
    return s


def split_multi_value(value: object) -> list[str]:
    """Розбиває «a@x.com, b@x.com» або «a;b» на окремі значення."""
    raw = str(value or "").strip()
    if not raw or raw.lower() in {"nan", "<na>"}:
        return []
    sep = "," if "," in raw else ";"
    return [v.strip() for v in raw.split(sep) if v.strip()]


def _split_cell(value: str) -> list[str]:
    return split_multi_value(value)


def crm_cell_contains_email(crm_cell: object, email: object) -> bool:
    """Чи email з дослідження входить у CRM-комірку (через кому / кракрапку з комою)."""
    key = normalize_email_key(email)
    if not key:
        return False
    return any(normalize_email_key(part) == key for part in split_multi_value(crm_cell))


def crm_cell_contains_linkedin_person(crm_cell: object, linkedin: object) -> bool:
    """Чи Linkedin Person входить у CRM-комірку (через кому / кракрапку з комою)."""
    key = linkedin_person_match_key(linkedin)
    if not key:
        return False
    return any(
        linkedin_person_match_key(part) == key for part in split_multi_value(crm_cell)
    )


def emails_display_from_row(row: pd.Series, email_cols: list[str]) -> str:
    """Усі email з рядка people, як VLOOKUP people!B (через кому)."""
    parts: list[str] = []
    seen: set[str] = set()
    for col in email_cols:
        raw = str(row.get(col, "")).strip()
        if not raw:
            continue
        for item in _split_cell(raw):
            key = normalize_email_key(item)
            if key and key not in seen:
                seen.add(key)
                parts.append(item)
    return ", ".join(parts)


def normalize_email_key(value: object) -> str:
    """Ключ для метчу email (як у формулах Comp)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    raw = str(value).strip().lower()
    if not raw:
        return ""
    if "<" in raw and ">" in raw:
        m = re.search(r"<([^>]+)>", raw)
        if m:
            raw = m.group(1).strip()
    raw = re.sub(r"^mailto:", "", raw)
    return raw.strip()


def format_person_csv_help() -> str:
    return (
        "**Довідник Person CRM** — експорт **people** з CRM (не файл лідів для імпорту), "
        "наприклад `people-….csv`:\n\n"
        "| Колонка | Роль |\n"
        "|---------|------|\n"
        "| **Person - Email - Work** / Home / Other | ключі пошуку → **Person - ID by Email** |\n"
        "| **Person - LinkedIn** | ключ → **ID by Linkedin Person** |\n"
        "| **Person - ID** | **Person - ID** + **CRM by Email ID** / **CRM Linkedin ID** (VLOOKUP за ID) |\n"
        "| **Person - Name** | не для метчу |\n\n"
        "У лідах з дослідження колонка **Email** зіставляється з будь-яким із цих email.\n\n"
        "**Метч:** email → ID; якщо немає — **Linkedin Person** → ID."
    )


def is_persons_crm_export(df: pd.DataFrame) -> bool:
    """Експорт people (~Person - ID + LinkedIn), не шаблон comp для імпорту."""
    detected = detect_person_csv_columns(df)
    if not detected.get("id"):
        return False
    comp_markers = (
        "Activity Subject",
        "ICP name",
        "Organization - Research ICP",
        "Person - Linkedin Active",
    )
    if any(m in df.columns for m in comp_markers):
        return False
    return bool(detected.get("linkedin") or detected.get("emails"))


def format_leads_person_columns_help() -> str:
    from crm_export import CRM_PPL_LEADS_COLUMNS

    cols = "\n".join(f"{i + 1}. `{c}`" for i, c in enumerate(CRM_PPL_LEADS_COLUMNS))
    return (
        "**PPL for CRM** — [шаблон Google](https://docs.google.com/spreadsheets/d/"
        "1VmDcr8_qYJYDVa-6qdN_XFk1E6_II2aveIyLKjGqedg/edit?gid=0#gid=0), "
        "вкладка **PPL** (24 колонки):\n\n"
        f"{cols}\n\n"
        "ID Person / Organization — з **«2 — Ліди CRM»** (people) та **«1 — Організації CRM»**."
    )


@dataclass(frozen=True)
class PersonLookup:
    row_count: int
    by_email: dict[str, str]
    by_linkedin: dict[str, str]
    by_id_email: dict[str, str]
    by_id_linkedin: dict[str, str]

    @classmethod
    def from_dataframe(cls, df: pd.DataFrame) -> PersonLookup:
        detected = detect_person_csv_columns(df)
        email_cols = detected.get("emails") or []
        if not email_cols and detected.get("email"):
            email_cols = [detected["email"]]
        linkedin_col = detected["linkedin"]
        id_col = detected["id"]
        if not id_col:
            found = ", ".join(_clean_header(c) for c in df.columns[:20])
            extra = "" if len(df.columns) <= 20 else f" … (+{len(df.columns) - 20})"
            raise ValueError(
                f"У файлі контактів немає колонки «{PERSON_ID_COLUMN}». "
                f"Знайдені заголовки: {found}{extra}."
            )

        by_email: dict[str, str] = {}
        by_linkedin: dict[str, str] = {}
        by_id_email: dict[str, str] = {}
        by_id_linkedin: dict[str, str] = {}

        for _, row in df.iterrows():
            person_id = normalize_person_id(row.get(id_col, ""))
            if not person_id:
                continue
            li = str(row.get(linkedin_col, "")).strip() if linkedin_col else ""
            emails_display = emails_display_from_row(row, email_cols)
            if emails_display:
                existing = by_id_email.get(person_id, "")
                if existing and existing != emails_display:
                    merged = {normalize_email_key(p) for p in existing.split(",")}
                    for part in emails_display.split(","):
                        part = part.strip()
                        if part and normalize_email_key(part) not in merged:
                            existing = f"{existing}, {part}"
                            merged.add(normalize_email_key(part))
                    by_id_email[person_id] = existing
                else:
                    by_id_email[person_id] = emails_display or existing
            if li:
                existing_li = by_id_linkedin.get(person_id, "")
                if existing_li:
                    seen_li = {linkedin_person_match_key(x) for x in _split_cell(existing_li)}
                    for li_item in _split_cell(li):
                        if linkedin_person_match_key(li_item) not in seen_li:
                            existing_li = f"{existing_li}, {li_item}"
                    by_id_linkedin[person_id] = existing_li
                else:
                    by_id_linkedin[person_id] = li

            for email_col in email_cols:
                for email_item in _split_cell(str(row.get(email_col, "")).strip()):
                    ekey = normalize_email_key(email_item)
                    if ekey and ekey not in by_email:
                        by_email[ekey] = person_id

            for li_item in _split_cell(li) if li else []:
                lkey = linkedin_person_match_key(li_item)
                if lkey and lkey not in by_linkedin:
                    by_linkedin[lkey] = person_id

        return cls(
            row_count=len(df),
            by_email=by_email,
            by_linkedin=by_linkedin,
            by_id_email=by_id_email,
            by_id_linkedin=by_id_linkedin,
        )

    @classmethod
    def from_csv(cls, path: Path | str) -> PersonLookup:
        p = Path(path)
        df = pd.read_csv(p, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        df.columns = [str(c).strip().lstrip("\ufeff") for c in df.columns]
        return cls.from_dataframe(df)

    def id_by_email(self, email: object) -> str:
        return self.by_email.get(normalize_email_key(email), "")

    def id_by_linkedin(self, linkedin: object) -> str:
        return self.by_linkedin.get(linkedin_person_match_key(linkedin), "")

    def resolve_person_id(self, email: object, linkedin: object) -> str:
        """Як Person - ID на Comp: email, інакше LinkedIn Person."""
        if not normalize_email_key(email) and not linkedin_person_match_key(linkedin):
            return ""
        found = self.id_by_email(email)
        if found:
            return found
        return self.id_by_linkedin(linkedin)

    def crm_email_for_id(self, person_id: object) -> str:
        """PPL M: VLOOKUP(Person - ID, people!A:C, 2) — email(и) з CRM."""
        key = normalize_person_id(person_id)
        if not key:
            return ""
        return self.by_id_email.get(key, "")

    def crm_linkedin_for_id(self, person_id: object) -> str:
        """PPL N: VLOOKUP(Person - ID, people!A:C, 3) — LinkedIn з CRM."""
        key = normalize_person_id(person_id)
        if not key:
            return ""
        return self.by_id_linkedin.get(key, "")


def apply_person_lookup(
    df: pd.DataFrame,
    lookup: PersonLookup | None,
    *,
    email_column: str = "Email",
    linkedin_column: str = "Linkedin Person",
    keep_existing_person_id: bool = False,
) -> pd.DataFrame:
    """Додає/оновлює Person-колонки як на вкладці Comp."""
    out = df.copy()
    for col in (
        PERSON_ID_COLUMN,
        PERSON_ID_BY_EMAIL_COLUMN,
        ID_BY_LINKEDIN_PERSON_COLUMN,
        CRM_BY_EMAIL_ID_COLUMN,
        CRM_LINKEDIN_ID_COLUMN,
    ):
        if col not in out.columns:
            out[col] = ""

    if lookup is None:
        return out

    emails = out[email_column] if email_column in out.columns else ""
    linkedins = out[linkedin_column] if linkedin_column in out.columns else ""

    person_ids: list[str] = []
    email_ids: list[str] = []
    linkedin_ids: list[str] = []
    crm_by_email: list[str] = []
    crm_by_linkedin: list[str] = []

    existing = (
        out[PERSON_ID_COLUMN].astype(str).str.strip()
        if keep_existing_person_id and PERSON_ID_COLUMN in out.columns
        else pd.Series([""] * len(out))
    )

    for i, (email, li) in enumerate(zip(emails, linkedins)):
        eid = lookup.id_by_email(email)
        lid = lookup.id_by_linkedin(li)
        manual = existing.iloc[i] if keep_existing_person_id else ""
        pid = normalize_person_id(
            manual if manual else lookup.resolve_person_id(email, li)
        )
        person_ids.append(pid)
        email_ids.append(normalize_person_id(eid))
        linkedin_ids.append(normalize_person_id(lid))
        crm_by_email.append(lookup.crm_email_for_id(pid) if pid else "")
        crm_by_linkedin.append(lookup.crm_linkedin_for_id(pid) if pid else "")

    out[PERSON_ID_COLUMN] = person_ids
    out[PERSON_ID_BY_EMAIL_COLUMN] = email_ids
    out[ID_BY_LINKEDIN_PERSON_COLUMN] = linkedin_ids
    out[CRM_BY_EMAIL_ID_COLUMN] = crm_by_email
    out[CRM_LINKEDIN_ID_COLUMN] = crm_by_linkedin
    return out
