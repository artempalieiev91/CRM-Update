"""Пошук Organization - ID та CRM URL за експортом організацій з CRM."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from person_lookup import split_multi_value

ORG_WEBSITE_COLUMN = "Organization - Website (in use)"
ORG_LINKEDIN_COLUMN = "Organization - LinkedIn (in use)"
ORG_ID_COLUMN = "Organization - ID"

WEBSITE_ID_COLUMN = "Website ID"
LINKEDIN_ID_COLUMN = "Linkedin ID"
CRM_WEBSITE_ID_COLUMN = "CRM Website ID"
CRM_LINKEDIN_ID_COLUMN = "CRM Linkedin ID"

_WEBSITE_COL_ALIASES = (
    ORG_WEBSITE_COLUMN,
    "Organization - Website",
    "Website",
)
_LINKEDIN_COL_ALIASES = (
    ORG_LINKEDIN_COLUMN,
    "Organization - LinkedIn",
    "Organization - Linkedin",
    "Linkedin Company",
)
_ID_COL_ALIASES = (
    ORG_ID_COLUMN,
    "Organization - ID",
    "Organization ID",
    "Organization - Id",
    "ID",
    "Org ID",
    "Pipedrive Organization ID",
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


def detect_organization_csv_columns(df: pd.DataFrame) -> dict[str, str | None]:
    """Які колонки знайдено в завантаженому CRM-експорті організацій."""
    return {
        "website": _resolve_column(
            df,
            _WEBSITE_COL_ALIASES,
            fuzzy_contains=("organization", "website"),
        ),
        "linkedin": _resolve_column(
            df,
            _LINKEDIN_COL_ALIASES,
            fuzzy_contains=("organization", "linkedin"),
        ),
        "id": _resolve_column(
            df,
            _ID_COL_ALIASES,
            fuzzy_contains=("organization", "id"),
        ),
    }


def format_organizations_crm_workflow_help() -> str:
    return (
        "### Цикл для **компаній** (як у вашому процесі)\n\n"
        "**Етап 1 — до імпорту в CRM**\n"
        "1. Вигрузка з CRM → завантажити в скрипт **«1 — Організації CRM»** (перший файл).\n"
        "2. Зібрати **companies** з дослідження, підставити ID де вже є в базі.\n"
        "3. **Скачати companies CSV** → імпорт у CRM.\n\n"
        "**Етап 2 — після імпорту в CRM**\n"
        "4. У CRM з’явились нові організації з ID → **знову вигрузити organizations** з CRM.\n"
        "5. **Замінити** файл у **«1 — Організації CRM»** (той самий пункт у sidebar).\n"
        "6. На **CRM companies** натиснути **«Перерахувати ID з бази організацій»** "
        "(і за потреби те саме для лідів PPL).\n\n"
        "Після етапу 2 у базі будуть домени/ID щойно створених компаній — метч стане повнішим. "
        "На етапі 1 частина ID порожня — **нормально**, бо їх ще немає в CRM."
    )


def format_organization_csv_help() -> str:
    return (
        format_organizations_crm_workflow_help()
        + "\n\n---\n\n"
        "**Колонки файлу організацій** (як **Copy of organizations**):\n\n"
        "| Колонка | Роль |\n"
        "|---------|------|\n"
        "| **Organization - Website** | ключ → **Website ID** |\n"
        "| **Organization - LinkedIn** | запасний ключ → **Linkedin ID** |\n"
        "| **Organization - ID** | числовий ID + CRM URL за ID |\n\n"
        "Суфікс **(in use)** у повному експорті CRM — ті самі поля. "
        "**Метч:** **Website** → ID; інакше **LinkedIn Company** → ID."
    )


def format_company_export_help() -> str:
    from company_export import COMPANY_CRM_COLUMNS

    cols = "\n".join(f"{i + 1}. `{c}`" for i, c in enumerate(COMPANY_CRM_COLUMNS))
    return (
        "**Файл компаній для CRM** (вкладка **Comp** у шаблоні) — один рядок на компанію:\n\n"
        f"{cols}\n\n"
        "Колонки **Organization - ID** … **CRM Linkedin ID** — з **«1 — Організації CRM»**. "
        "Після імпорту companies у CRM **оновіть** цей файл (етап 2) і перерахуйте ID."
    )


def normalize_website_key(value: object) -> str:
    """Як Website на Comp: без протоколу/www, лише домен до першого /."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    raw = str(value).strip().lower()
    if not raw or raw in {"nan", "none", "<na>", "n/a"}:
        return ""
    raw = re.sub(r"^mailto:", "", raw)
    if "@" in raw:
        raw = raw.split("@", 1)[1]
    raw = re.sub(r"^https?://", "", raw)
    raw = re.sub(r"^www\.", "", raw)
    raw = raw.split("/")[0].split("?")[0].split("#")[0]
    raw = raw.split(":")[0].rstrip(".")
    return raw


def crm_cell_contains_website(crm_cell: object, website: object) -> bool:
    """Чи Website входить у CRM-комірку (через кому / кракрапку з комою)."""
    key = normalize_website_key(website)
    if not key:
        return False
    return any(
        normalize_website_key(part) == key for part in split_multi_value(crm_cell)
    )


def crm_cell_contains_linkedin(crm_cell: object, linkedin: object) -> bool:
    """Чи LinkedIn компанії входить у CRM-комірку (через кому / кракрапку з комою)."""
    key = normalize_linkedin_key(linkedin)
    if not key:
        return False
    return any(
        normalize_linkedin_key(part) == key for part in split_multi_value(crm_cell)
    )


def format_comp_crm_url(value: object) -> str:
    """
    Як вихід формул CRM Website ID / CRM Linkedin ID на Comp (кол. L, M):
    REGEXREPLACE: прибрати https?://, www., хвости / (двічі).
    """
    raw = str(value or "").strip()
    if not raw:
        return ""
    raw = re.sub(r"^https?://", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"^www\.", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"/+$", "", raw)
    raw = re.sub(r"/+$", "", raw)
    return raw


def normalize_linkedin_key(value: object) -> str:
    """Нормалізація LinkedIn URL для VLOOKUP."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    raw = str(value).strip().lower()
    if not raw:
        return ""
    raw = re.sub(r"^https?://", "", raw)
    raw = re.sub(r"^www\.", "", raw)
    raw = re.sub(r"/+$", "", raw)
    return raw


@dataclass(frozen=True)
class OrganizationLookup:
    """Індекси з CRM-експорту organizations (*.csv)."""

    row_count: int
    by_website: dict[str, str]
    by_linkedin: dict[str, str]
    by_id: dict[str, tuple[str, str]]

    @classmethod
    def from_dataframe(cls, df: pd.DataFrame) -> OrganizationLookup:
        detected = detect_organization_csv_columns(df)
        website_col = detected["website"]
        linkedin_col = detected["linkedin"]
        id_col = detected["id"]
        if not id_col:
            found = ", ".join(_clean_header(c) for c in df.columns[:20])
            extra = "" if len(df.columns) <= 20 else f" … (+{len(df.columns) - 20})"
            raise ValueError(
                f"У файлі організацій немає колонки «{ORG_ID_COLUMN}» (або еквіваленту). "
                f"Знайдені заголовки: {found}{extra}. "
                f"Потрібні мінімум: ID + Website (+ LinkedIn за бажанням)."
            )

        by_website: dict[str, str] = {}
        by_linkedin: dict[str, str] = {}
        by_id: dict[str, tuple[str, str]] = {}

        for _, row in df.iterrows():
            org_id = str(row.get(id_col, "")).strip()
            if not org_id:
                continue
            web = str(row.get(website_col, "")).strip() if website_col else ""
            li = str(row.get(linkedin_col, "")).strip() if linkedin_col else ""
            by_id[org_id] = (web, li)

            for w in split_multi_value(web):
                wkey = normalize_website_key(w)
                if wkey and wkey not in by_website:
                    by_website[wkey] = org_id

            for l in split_multi_value(li):
                lkey = normalize_linkedin_key(l)
                if lkey and lkey not in by_linkedin:
                    by_linkedin[lkey] = org_id

        return cls(
            row_count=len(df),
            by_website=by_website,
            by_linkedin=by_linkedin,
            by_id=by_id,
        )

    @classmethod
    def from_csv(cls, path: Path | str) -> OrganizationLookup:
        p = Path(path)
        df = pd.read_csv(p, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        df.columns = [str(c).strip().lstrip("\ufeff") for c in df.columns]
        return cls.from_dataframe(df)

    def id_by_website(self, website: object) -> str:
        return self.by_website.get(normalize_website_key(website), "")

    def id_by_linkedin(self, linkedin: object) -> str:
        return self.by_linkedin.get(normalize_linkedin_key(linkedin), "")

    def resolve_organization_id(self, website: object, linkedin: object) -> str:
        """Як Organization - ID на Comp: website, інакше LinkedIn."""
        if not str(website or "").strip() and not str(linkedin or "").strip():
            return ""
        found = self.id_by_website(website)
        if found:
            return found
        return self.id_by_linkedin(linkedin)

    def crm_urls_for_id(self, org_id: object) -> tuple[str, str]:
        if not org_id:
            return "", ""
        entry = self.by_id.get(str(org_id).strip())
        if not entry:
            return "", ""
        web, li = entry
        return format_comp_crm_url(web), format_comp_crm_url(li)


def leads_have_org_lookup_columns(df: pd.DataFrame) -> bool:
    """Ліди з дослідження (Comp) або PPL — є Website / Linkedin Company."""
    return "Website" in df.columns or "Linkedin Company" in df.columns


def leads_use_org_crm_url_columns(df: pd.DataFrame) -> bool:
    """
    Чи заповнювати CRM Website ID / CRM Linkedin ID з бази організацій.
    На PPL (є Linkedin Person) ці колонки — Person з people, не Linkedin Company.
    """
    return "Linkedin Person" not in df.columns


def apply_leads_organization_lookup(
    df: pd.DataFrame,
    lookup: OrganizationLookup | None,
    *,
    keep_existing_org_id: bool = True,
    fill_org_crm_url_columns: bool | None = None,
) -> pd.DataFrame:
    """Organization - ID на лідах: Website, інакше Linkedin Company."""
    if lookup is None or not leads_have_org_lookup_columns(df):
        return df
    if fill_org_crm_url_columns is None:
        fill_org_crm_url_columns = leads_use_org_crm_url_columns(df)
    return apply_organization_lookup(
        df,
        lookup,
        website_column="Website",
        linkedin_column="Linkedin Company",
        keep_existing_org_id=keep_existing_org_id,
        fill_crm_url_columns=fill_org_crm_url_columns,
    )


def apply_organization_lookup(
    df: pd.DataFrame,
    lookup: OrganizationLookup | None,
    *,
    website_column: str = "Website",
    linkedin_column: str = "Linkedin Company",
    keep_existing_org_id: bool = False,
    fill_crm_url_columns: bool = True,
) -> pd.DataFrame:
    """Додає/оновлює ID-колонки як на вкладці Comp."""
    out = df.copy()
    for col in (
        "Organization - ID",
        WEBSITE_ID_COLUMN,
        LINKEDIN_ID_COLUMN,
        CRM_WEBSITE_ID_COLUMN,
        CRM_LINKEDIN_ID_COLUMN,
    ):
        if col not in out.columns:
            out[col] = ""

    if lookup is None:
        return out

    websites = out[website_column] if website_column in out.columns else ""
    linkedins = out[linkedin_column] if linkedin_column in out.columns else ""

    org_ids: list[str] = []
    website_ids: list[str] = []
    linkedin_ids: list[str] = []

    existing_org = (
        out["Organization - ID"].astype(str).str.strip()
        if keep_existing_org_id and "Organization - ID" in out.columns
        else pd.Series([""] * len(out))
    )

    crm_webs: list[str] | None = [] if fill_crm_url_columns else None
    crm_lis: list[str] | None = [] if fill_crm_url_columns else None

    for i, (web, li) in enumerate(zip(websites, linkedins)):
        wid = lookup.id_by_website(web)
        lid = lookup.id_by_linkedin(li)
        manual_oid = existing_org.iloc[i] if keep_existing_org_id else ""
        oid = manual_oid if manual_oid else lookup.resolve_organization_id(web, li)
        org_ids.append(oid)
        website_ids.append(wid)
        linkedin_ids.append(lid)
        if fill_crm_url_columns:
            cw, cl = lookup.crm_urls_for_id(oid) if oid else ("", "")
            assert crm_webs is not None and crm_lis is not None
            crm_webs.append(cw)
            crm_lis.append(cl)

    out["Organization - ID"] = org_ids
    out[WEBSITE_ID_COLUMN] = website_ids
    out[LINKEDIN_ID_COLUMN] = linkedin_ids
    if fill_crm_url_columns and crm_webs is not None and crm_lis is not None:
        out[CRM_WEBSITE_ID_COLUMN] = crm_webs
        out[CRM_LINKEDIN_ID_COLUMN] = crm_lis
    return out
