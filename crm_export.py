"""Формування CSV для імпорту в CRM (варіант comp)."""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

import pandas as pd

from excel_utils import excel_cell_value
from ids_runtime import log_disk_failure, log_disk_write
from industry_mapping import normalize_industry_column
from linkedin_urls import (
    LINKEDIN_HELPER_ID_COLUMN,
    LINKEDIN_PERSON_COLUMN,
    linkedin_helper_id_from_person,
    linkedin_person_match_key,
)
from mx_checker import EMAIL_COLUMN, fill_email_providers
from organization_lookup import (
    OrganizationLookup,
    apply_leads_organization_lookup,
    apply_organization_lookup,
    normalize_website_key,
)
from person_lookup import (
    CRM_BY_EMAIL_ID_COLUMN,
    CRM_LINKEDIN_ID_COLUMN,
    ID_BY_LINKEDIN_PERSON_COLUMN,
    PERSON_ID_BY_EMAIL_COLUMN,
    PERSON_ID_COLUMN,
    PERSON_HIGHLIGHT_COLUMN,
    PERSON_LOOKUP_ISSUE_COLUMN,
    PersonLookup,
    apply_person_lookup,
    crm_cell_contains_email,
    crm_cell_contains_linkedin_person,
    normalize_person_id,
)
from state_abbreviations import normalize_state_columns

ICP_NAME_COLUMN = "ICP name"
CONFERENCE_COLUMN = "Conference"
SOURCE_SHEET_COLUMN = "Source sheet"
RESEARCH_ID_COLUMN = "Research ID"
NOTION_CAMPAIGN_ID_COLUMN = "Notion Campaign ID"
ACTIVITY_TYPE_COLUMN = "Activity Type"
ACTIVITY_DUE_DATE_COLUMN = "Activity - Due date"

# За замовчуванням для PPL; змініть тут або передайте Activity Type у дослідженні/CSV — тоді підставиться з файлу.
DEFAULT_ACTIVITY_TYPE = "General Task"

# Notion Campaign ID: зараз вимкнено. True — для рядка з Activity Subject = ICP / Conference підставляється Research ID.
FILL_NOTION_CAMPAIGN_ID = False

# PPL for CRM Template — https://docs.google.com/spreadsheets/d/1VmDcr8_qYJYDVa-6qdN_XFk1E6_II2aveIyLKjGqedg (вкладка PPL)
CRM_PPL_LEADS_COLUMNS: tuple[str, ...] = (
    "Activity Subject",
    "First name",
    "Last name",
    "Title",
    "Email",
    "Linkedin Person",
    "Person - Apollo Contact id",
    PERSON_ID_COLUMN,
    PERSON_ID_BY_EMAIL_COLUMN,
    ID_BY_LINKEDIN_PERSON_COLUMN,
    "Website",
    "Linkedin Company",
    CRM_BY_EMAIL_ID_COLUMN,
    CRM_LINKEDIN_ID_COLUMN,
    "Organization - ID",
    "Person Country",
    "Person State",
    "Person City",
    "Organization - Email provider",
    NOTION_CAMPAIGN_ID_COLUMN,
    ACTIVITY_TYPE_COLUMN,
    "Person - Linkedin Active",
    LINKEDIN_HELPER_ID_COLUMN,
    ACTIVITY_DUE_DATE_COLUMN,
)

# Після колонок PPL у CSV перевірки лідів
PPL_LEADS_REVIEW_TAIL_COLUMNS: tuple[str, ...] = (
    PERSON_HIGHLIGHT_COLUMN,
    PERSON_LOOKUP_ISSUE_COLUMN,
)

CRM_COMP_COLUMNS: tuple[str, ...] = (
    "Activity Subject",
    "Organization - Research ICP",
    "Company",
    "Website",
    "Industry",
    "Country",
    "State",
    "City",
    "First name",
    "Last name",
    "Title",
    "Email",
    "Linkedin Person",
    PERSON_ID_BY_EMAIL_COLUMN,
    ID_BY_LINKEDIN_PERSON_COLUMN,
    PERSON_ID_COLUMN,
    CRM_BY_EMAIL_ID_COLUMN,
    CRM_LINKEDIN_ID_COLUMN,
    "Organization - ID",
    "Linkedin Company",
    "Number of employees",
    "Person Country",
    "Person State",
    "Person City",
    "Research ID",
    "Person - Linkedin Active",
    "Organization - Email provider",
    "Date",
    "Person - Apollo Contact id",
    "Organization - Apollo Account id",
    "Researcher Name",
)


def _fill_na_str(df: pd.DataFrame) -> pd.DataFrame:
    return df.fillna("").astype(str).replace({"nan": "", "<NA>": ""})


def _event_row_mask(src: pd.DataFrame) -> pd.Series:
    """Event-лист / рядок: Research ID або Source sheet містить EVENT, або ICP = Event(s)."""
    mask = pd.Series(False, index=src.index)
    if RESEARCH_ID_COLUMN in src.columns:
        mask |= (
            src[RESEARCH_ID_COLUMN]
            .astype(str)
            .str.strip()
            .str.casefold()
            .str.contains("event", na=False)
        )
    if SOURCE_SHEET_COLUMN in src.columns:
        mask |= (
            src[SOURCE_SHEET_COLUMN]
            .astype(str)
            .str.strip()
            .str.casefold()
            .str.contains("event", na=False)
        )
    if ICP_NAME_COLUMN in src.columns:
        icp = src[ICP_NAME_COLUMN].astype(str).str.strip().str.casefold()
        mask |= icp.isin({"event", "events"})
    return mask


def activity_due_date_today() -> str:
    """Дата активності на момент збірки файлу: MM/DD/YYYY (як 06/03/2026 у PPL)."""
    return datetime.now().strftime("%m/%d/%Y")


def fill_activity_type_column(
    out: pd.DataFrame,
    src: pd.DataFrame,
    *,
    default: str = DEFAULT_ACTIVITY_TYPE,
) -> pd.Series:
    """Activity Type: з джерела, якщо заповнено; інакше default (зараз General Task)."""
    n = len(out)
    if ACTIVITY_TYPE_COLUMN in src.columns:
        from_src = (
            src[ACTIVITY_TYPE_COLUMN]
            .astype(str)
            .str.strip()
            .replace({"nan": "", "<NA>": ""})
        )
        return from_src.where(from_src != "", default)
    return pd.Series([default] * n, index=out.index)


def fill_notion_campaign_id_column(
    out: pd.DataFrame,
    src: pd.DataFrame,
    *,
    enabled: bool = FILL_NOTION_CAMPAIGN_ID,
) -> pd.Series:
    """
    Notion Campaign ID: зараз порожньо.
    Якщо enabled або в джерелі вже є значення — підставка з файлу.
    Інакше (enabled): метч Activity Subject з ICP name / Conference → Research ID того ж рядка.
    """
    n = len(out)
    empty = pd.Series([""] * n, index=out.index)

    if NOTION_CAMPAIGN_ID_COLUMN in src.columns:
        from_src = (
            src[NOTION_CAMPAIGN_ID_COLUMN]
            .astype(str)
            .str.strip()
            .replace({"nan": "", "<NA>": ""})
        )
        if from_src.ne("").any():
            return from_src

    if not enabled or RESEARCH_ID_COLUMN not in src.columns:
        return empty

    research = (
        src[RESEARCH_ID_COLUMN]
        .astype(str)
        .str.strip()
        .replace({"nan": "", "<NA>": ""})
    )
    if "Activity Subject" not in out.columns:
        return research

    subject = (
        out["Activity Subject"]
        .astype(str)
        .str.strip()
        .replace({"nan": "", "<NA>": ""})
    )
    icp = (
        src[ICP_NAME_COLUMN]
        .astype(str)
        .str.strip()
        .replace({"nan": "", "<NA>": ""})
        if ICP_NAME_COLUMN in src.columns
        else pd.Series([""] * n, index=src.index)
    )
    conference = (
        src[CONFERENCE_COLUMN]
        .astype(str)
        .str.strip()
        .replace({"nan": "", "<NA>": ""})
        if CONFERENCE_COLUMN in src.columns
        else pd.Series([""] * n, index=src.index)
    )

    matched = (subject != "") & (research != "") & (
        (subject == icp) | (subject == conference)
    )
    return research.where(matched, "")


ACTIVITY_SUBJECT_COLUMN = "Activity Subject"
ACTIVITY_SUBJECT_SOURCE_ALIASES: tuple[str, ...] = (
    ACTIVITY_SUBJECT_COLUMN,
    "Activity - Subject",
)


def _normalize_activity_subject_series(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().replace({"nan": "", "<NA>": ""})


def resolve_activity_subject_source_column(df: pd.DataFrame) -> str | None:
    """Колонка subject у дослідженні: Activity Subject або Activity - Subject."""
    for name in ACTIVITY_SUBJECT_SOURCE_ALIASES:
        if name in df.columns:
            return name
    for col in df.columns:
        key = str(col).strip().casefold().replace("-", " ").replace("  ", " ")
        if key == "activity subject":
            return col
    return None


def build_activity_subject(src: pd.DataFrame) -> pd.Series:
    """Event → Conference; інакше Activity Subject з джерела (якщо є колонка), інакше ICP name."""
    n = len(src)
    icp = (
        _normalize_activity_subject_series(src[ICP_NAME_COLUMN])
        if ICP_NAME_COLUMN in src.columns
        else pd.Series([""] * n, index=src.index)
    )

    subject_col = resolve_activity_subject_source_column(src)
    has_icp_col = ICP_NAME_COLUMN in src.columns
    if subject_col is not None and has_icp_col:
        from_src = _normalize_activity_subject_series(src[subject_col])
        subject = from_src.where(from_src != "", icp)
    elif subject_col is not None:
        subject = _normalize_activity_subject_series(src[subject_col])
    else:
        subject = icp.copy()

    event_mask = _event_row_mask(src)
    if not event_mask.any() or CONFERENCE_COLUMN not in src.columns:
        return subject

    conference = src[CONFERENCE_COLUMN].astype(str).str.strip().replace({"nan": "", "<NA>": ""})
    subject = subject.copy()
    subject.loc[event_mask] = conference.loc[event_mask]
    missing = event_mask & (conference == "")
    subject.loc[missing] = icp.loc[missing]
    return subject


def is_ppl_leads_template(df: pd.DataFrame) -> bool:
    """Шаблон «PPL for CRM», не Comp з ICP/Company/Industry."""
    return any(
        m in df.columns
        for m in (
            NOTION_CAMPAIGN_ID_COLUMN,
            ACTIVITY_TYPE_COLUMN,
            LINKEDIN_HELPER_ID_COLUMN,
            ACTIVITY_DUE_DATE_COLUMN,
        )
    )


def ppl_leads_crm_file_columns(
    df: pd.DataFrame,
    *,
    include_notion_campaign_id: bool = True,
) -> list[str]:
    """Колонки PPL для імпорту в CRM (без службових колонок перевірки)."""
    cols = [c for c in CRM_PPL_LEADS_COLUMNS if c in df.columns]
    if not include_notion_campaign_id:
        cols = [c for c in cols if c != NOTION_CAMPAIGN_ID_COLUMN]
    return cols


def ppl_leads_for_crm_file(
    df: pd.DataFrame,
    *,
    include_notion_campaign_id: bool = True,
) -> pd.DataFrame:
    """Таблиця лідів лише з колонками PPL."""
    cols = ppl_leads_crm_file_columns(
        df, include_notion_campaign_id=include_notion_campaign_id
    )
    return df.loc[:, cols].reset_index(drop=True)


def align_ppl_leads_export(
    df: pd.DataFrame,
    *,
    issue_column: str | None = PERSON_LOOKUP_ISSUE_COLUMN,
    include_review_columns: bool = False,
) -> pd.DataFrame:
    """Колонки в порядку Google PPL + опційно колонки перевірки в кінці."""
    out = df.copy()
    for col in CRM_PPL_LEADS_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    ordered = list(CRM_PPL_LEADS_COLUMNS)
    if include_review_columns:
        for col in PPL_LEADS_REVIEW_TAIL_COLUMNS:
            if col not in out.columns:
                out[col] = ""
        ordered.extend(PPL_LEADS_REVIEW_TAIL_COLUMNS)
    elif issue_column:
        if issue_column not in out.columns:
            out[issue_column] = ""
        ordered.append(issue_column)
    return out[ordered]


def build_ppl_leads_export(
    df: pd.DataFrame,
    *,
    person_lookup: PersonLookup | None = None,
    org_lookup: OrganizationLookup | None = None,
    keep_existing_org_id: bool = True,
    fill_notion_campaign_id: bool | None = None,
    default_activity_type: str | None = None,
) -> pd.DataFrame:
    """Збирає ліди в порядку вкладки PPL (дослідження або вже PPL-файл)."""
    src = _fill_na_str(df)
    n = len(src)
    out = pd.DataFrame()

    if ICP_NAME_COLUMN in src.columns:
        out["Activity Subject"] = build_activity_subject(src)
    elif "Activity Subject" in src.columns:
        out["Activity Subject"] = src["Activity Subject"]
    else:
        out["Activity Subject"] = [""] * n

    for col in CRM_PPL_LEADS_COLUMNS:
        if col == "Activity Subject":
            continue
        if col in src.columns:
            out[col] = src[col]
        else:
            out[col] = [""] * n

    if n:
        out[ACTIVITY_DUE_DATE_COLUMN] = activity_due_date_today()
        activity_default = (
            DEFAULT_ACTIVITY_TYPE
            if default_activity_type is None
            else (str(default_activity_type).strip() or DEFAULT_ACTIVITY_TYPE)
        )
        out[ACTIVITY_TYPE_COLUMN] = fill_activity_type_column(
            out, src, default=activity_default
        )
        notion_enabled = (
            FILL_NOTION_CAMPAIGN_ID
            if fill_notion_campaign_id is None
            else fill_notion_campaign_id
        )
        out[NOTION_CAMPAIGN_ID_COLUMN] = fill_notion_campaign_id_column(
            out, src, enabled=notion_enabled
        )

    if LINKEDIN_PERSON_COLUMN in out.columns:
        out[LINKEDIN_HELPER_ID_COLUMN] = out[LINKEDIN_PERSON_COLUMN].map(
            linkedin_helper_id_from_person
        )

    if "Website" in out.columns:
        out["Website"] = out["Website"].apply(normalize_website_key)

    out = normalize_state_columns(out, columns=("Person State",))
    if EMAIL_COLUMN in out.columns:
        out = fill_email_providers(out, overwrite=True)
    out = apply_person_lookup(out, person_lookup)
    out = apply_leads_organization_lookup(
        out,
        org_lookup,
        keep_existing_org_id=keep_existing_org_id,
        fill_org_crm_url_columns=False,
    )
    return align_ppl_leads_export(out, issue_column=None)


def prepare_ppl_leads_export(
    df: pd.DataFrame,
    *,
    person_lookup: PersonLookup | None = None,
    org_lookup: OrganizationLookup | None = None,
    keep_existing_org_id: bool = True,
    fill_notion_campaign_id: bool | None = None,
    default_activity_type: str | None = None,
) -> pd.DataFrame:
    """Alias для build_ppl_leads_export (зворотна сумісність)."""
    return build_ppl_leads_export(
        df,
        person_lookup=person_lookup,
        org_lookup=org_lookup,
        keep_existing_org_id=keep_existing_org_id,
        fill_notion_campaign_id=fill_notion_campaign_id,
        default_activity_type=default_activity_type,
    )


def align_comp_to_crm_schema(df: pd.DataFrame) -> pd.DataFrame:
    """Додає відсутні колонки Comp-шаблону (напр. після старого workspace)."""
    out = df.copy()
    for col in CRM_COMP_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    tail = [c for c in out.columns if c not in CRM_COMP_COLUMNS]
    return out[list(CRM_COMP_COLUMNS) + tail]


def _has_contact_value(value: object) -> bool:
    return str(value).strip() != ""


def _crm_person_id(value: object) -> bool:
    return bool(normalize_person_id(value))


def _has_crm_cell_value(value: object) -> bool:
    return str(value).strip() != ""


def _trimmed_cells_differ(a: object, b: object) -> bool:
    """Як TRIM(M)<>TRIM(E) у Google PPL."""
    return str(a).strip() != str(b).strip()


def _crm_by_email_highlight_mismatch(email: object, crm_by_email: object) -> bool:
    """
    PPL кол. M: CRM не порожній і Email не входить у список CRM (через кому).
    """
    if not _has_crm_cell_value(crm_by_email):
        return False
    if not _has_contact_value(email):
        return _trimmed_cells_differ(crm_by_email, email)
    return not crm_cell_contains_email(crm_by_email, email)


def _crm_linkedin_id_highlight_mismatch(
    linkedin_person: object, crm_linkedin_id: object
) -> bool:
    """
    PPL кол. N: CRM не порожній і Linkedin Person не входить у список CRM (через кому).
    """
    if not _has_crm_cell_value(crm_linkedin_id):
        return False
    if not _has_contact_value(linkedin_person):
        return _trimmed_cells_differ(crm_linkedin_id, linkedin_person)
    return not crm_cell_contains_linkedin_person(crm_linkedin_id, linkedin_person)


def _person_id_by_email_linkedin_highlight_mismatch(eid: object, lid: object) -> bool:
    """Person - ID by Email і ID by Linkedin Person обидва заповнені, але не збігаються."""
    e = str(eid).strip()
    l = str(lid).strip()
    return bool(_crm_person_id(e) and _crm_person_id(l) and e != l)


def person_highlight_issue_parts(row: pd.Series) -> list[str]:
    """Умовне форматування PPL: CRM by Email ID (M) та CRM Linkedin ID (N)."""
    issues: list[str] = []
    email = row.get(EMAIL_COLUMN, "")
    li = row.get(LINKEDIN_PERSON_COLUMN, "")
    crm_email = row.get(CRM_BY_EMAIL_ID_COLUMN, "")
    crm_li = row.get(CRM_LINKEDIN_ID_COLUMN, "")

    if _crm_by_email_highlight_mismatch(email, crm_email):
        issues.append(f"Підсвічено: {CRM_BY_EMAIL_ID_COLUMN}")
    if _crm_linkedin_id_highlight_mismatch(li, crm_li):
        issues.append(f"Підсвічено: {CRM_LINKEDIN_ID_COLUMN}")

    eid = str(row.get(PERSON_ID_BY_EMAIL_COLUMN, "")).strip()
    lid = str(row.get(ID_BY_LINKEDIN_PERSON_COLUMN, "")).strip()
    if _person_id_by_email_linkedin_highlight_mismatch(eid, lid):
        issues.append(f"Підсвічено: {PERSON_ID_BY_EMAIL_COLUMN}")
        issues.append(f"Підсвічено: {ID_BY_LINKEDIN_PERSON_COLUMN}")

    return issues


def person_extra_issue_parts(row: pd.Series) -> list[str]:
    """Перевірки lookup без кольорової підсвітки (кол. I, J не фарбуються в Google)."""
    issues: list[str] = []
    email = str(row.get(EMAIL_COLUMN, "")).strip()
    li = str(row.get(LINKEDIN_PERSON_COLUMN, "")).strip()
    eid = str(row.get(PERSON_ID_BY_EMAIL_COLUMN, "")).strip()
    lid = str(row.get(ID_BY_LINKEDIN_PERSON_COLUMN, "")).strip()

    if _has_contact_value(email) and not _crm_person_id(eid):
        issues.append("Email не знайдено в CRM (Person - ID by Email порожній)")
    if _has_contact_value(li) and not _crm_person_id(lid):
        issues.append("Linkedin Person не знайдено в CRM (ID by Linkedin Person порожній)")

    return issues


def person_row_is_highlighted(row: pd.Series) -> bool:
    return bool(person_highlight_issue_parts(row))


def person_lookup_issue_label(row: pd.Series) -> str:
    return "; ".join(
        person_highlight_issue_parts(row) + person_extra_issue_parts(row)
    )


def person_id_stats(df: pd.DataFrame) -> dict[str, int]:
    total = len(df)
    if total == 0:
        return {
            "total": 0,
            "with_person_id": 0,
            "with_email_id": 0,
            "with_linkedin_id": 0,
            "with_crm_by_email": 0,
            "with_crm_linkedin": 0,
            "issues": 0,
        }

    def _count(col: str) -> int:
        if col not in df.columns:
            return 0
        return int(df[col].astype(str).str.strip().ne("").sum())

    return {
        "total": total,
        "with_person_id": _count(PERSON_ID_COLUMN),
        "with_email_id": _count(PERSON_ID_BY_EMAIL_COLUMN),
        "with_linkedin_id": _count(ID_BY_LINKEDIN_PERSON_COLUMN),
        "with_crm_by_email": _count(CRM_BY_EMAIL_ID_COLUMN),
        "with_crm_linkedin": _count(CRM_LINKEDIN_ID_COLUMN),
        "issues": person_highlighted_row_count(df),
    }


def _annotate_person_leads_review(work: pd.DataFrame) -> pd.DataFrame:
    out = work.copy()

    def _annotate_row(row: pd.Series) -> pd.Series:
        highlight = person_highlight_issue_parts(row)
        issue = "; ".join(highlight + person_extra_issue_parts(row))
        return pd.Series(
            {
                PERSON_LOOKUP_ISSUE_COLUMN: issue,
                PERSON_HIGHLIGHT_COLUMN: "yes" if highlight else "",
            }
        )

    labels = out.apply(_annotate_row, axis=1)
    out[PERSON_LOOKUP_ISSUE_COLUMN] = labels[PERSON_LOOKUP_ISSUE_COLUMN]
    out[PERSON_HIGHLIGHT_COLUMN] = labels[PERSON_HIGHLIGHT_COLUMN]
    return out


def _fix_zurich_person_city_for_review(work: pd.DataFrame) -> pd.DataFrame:
    """Person State = Zurich і Person City = Zurich → Person City = Zuerich."""
    out = work.copy()
    if "Person State" not in out.columns or "Person City" not in out.columns:
        return out
    state = out["Person State"].astype(str).str.strip()
    city = out["Person City"].astype(str).str.strip()
    mask = state.str.casefold().eq("zurich") & city.str.casefold().eq("zurich")
    if mask.any():
        out.loc[mask, "Person City"] = "Zuerich"
    return out


def prepare_person_leads_review(
    df: pd.DataFrame,
    *,
    org_lookup: OrganizationLookup | None = None,
) -> pd.DataFrame:
    """Повна таблиця PPL + колонки перевірки (для кешу та Excel)."""
    work = apply_leads_organization_lookup(
        df.copy(),
        org_lookup,
        keep_existing_org_id=True,
        fill_org_crm_url_columns=False,
    )
    work = _fix_zurich_person_city_for_review(work)
    work = _annotate_person_leads_review(work)
    return align_ppl_leads_export(work, include_review_columns=True)


def person_highlighted_row_count(df: pd.DataFrame) -> int:
    if PERSON_HIGHLIGHT_COLUMN in df.columns:
        return int((df[PERSON_HIGHLIGHT_COLUMN].astype(str).str.strip() == "yes").sum())
    return sum(1 for _, row in df.iterrows() if person_row_is_highlighted(row))


def person_leads_review_export(
    df: pd.DataFrame,
    *,
    org_lookup: OrganizationLookup | None = None,
) -> pd.DataFrame:
    """Усі ліди в порядку PPL (24 колонки) для перегляду, CSV та Excel."""
    work = prepare_person_leads_review(df, org_lookup=org_lookup)
    return ppl_leads_for_crm_file(work, include_notion_campaign_id=False)


def person_leads_review_stats(review_df: pd.DataFrame) -> dict[str, int]:
    """Метрики Person ID з уже підготовленої review-таблиці (без повторного export)."""
    total = len(review_df)
    if total == 0:
        return {
            "total": 0,
            "with_person_id": 0,
            "with_email_id": 0,
            "with_linkedin_id": 0,
            "with_crm_by_email": 0,
            "with_crm_linkedin": 0,
            "issues": 0,
        }

    def _count(col: str) -> int:
        if col not in review_df.columns:
            return 0
        return int(review_df[col].astype(str).str.strip().ne("").sum())

    return {
        "total": total,
        "with_person_id": _count(PERSON_ID_COLUMN),
        "with_email_id": _count(PERSON_ID_BY_EMAIL_COLUMN),
        "with_linkedin_id": _count(ID_BY_LINKEDIN_PERSON_COLUMN),
        "with_crm_by_email": _count(CRM_BY_EMAIL_ID_COLUMN),
        "with_crm_linkedin": _count(CRM_LINKEDIN_ID_COLUMN),
        "issues": person_highlighted_row_count(review_df),
    }


def person_problem_rows(
    df: pd.DataFrame,
    *,
    org_lookup: OrganizationLookup | None = None,
) -> pd.DataFrame:
    """Лише рядки з підсвіткою CRM by Email ID / CRM Linkedin ID (як у Google PPL)."""
    full = person_leads_review_export(df, org_lookup=org_lookup)
    mask = full.apply(person_row_is_highlighted, axis=1)
    return full.loc[mask].copy().reset_index(drop=True)


def _person_cell_highlight_kind(issue: str, column: str) -> str | None:
    if not issue:
        return None
    if column == CRM_BY_EMAIL_ID_COLUMN and f"Підсвічено: {CRM_BY_EMAIL_ID_COLUMN}" in issue:
        return "crm"
    if column == CRM_LINKEDIN_ID_COLUMN and f"Підсвічено: {CRM_LINKEDIN_ID_COLUMN}" in issue:
        return "crm"
    if column == PERSON_ID_BY_EMAIL_COLUMN and f"Підсвічено: {PERSON_ID_BY_EMAIL_COLUMN}" in issue:
        return "id"
    if column == ID_BY_LINKEDIN_PERSON_COLUMN and f"Підсвічено: {ID_BY_LINKEDIN_PERSON_COLUMN}" in issue:
        return "id"
    return None


def _person_cell_highlight_fill(issue: str, column: str) -> bool:
    return _person_cell_highlight_kind(issue, column) is not None


def person_leads_review_to_xlsx_bytes(
    df: pd.DataFrame,
    *,
    include_notion_campaign_id: bool = False,
    issues: pd.Series | None = None,
) -> bytes:
    """Excel з підсвіткою CRM (червоний) та Person ID (рожевий)."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill

    crm_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    id_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads review"

    columns = ppl_leads_crm_file_columns(
        df, include_notion_campaign_id=include_notion_campaign_id
    )
    ws.append(columns)
    issue_col = PERSON_LOOKUP_ISSUE_COLUMN if issues is None else None
    for idx, row in df.iterrows():
        if issues is not None:
            issue = str(issues.get(idx, "")).strip()
        elif issue_col and issue_col in df.columns:
            issue = str(row.get(issue_col, "")).strip()
        else:
            issue = person_lookup_issue_label(row)
        ws.append([excel_cell_value(c, row.get(c, "")) for c in columns])
        if not issue:
            continue
        excel_row = ws.max_row
        for col_idx, col_name in enumerate(columns, start=1):
            kind = _person_cell_highlight_kind(issue, col_name)
            if kind == "crm":
                ws.cell(row=excel_row, column=col_idx).fill = crm_fill
            elif kind == "id":
                ws.cell(row=excel_row, column=col_idx).fill = id_fill

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_comp_export(
    df: pd.DataFrame,
    *,
    person_lookup: PersonLookup | None = None,
    org_lookup: OrganizationLookup | None = None,
) -> pd.DataFrame:
    src = _fill_na_str(df)
    if ICP_NAME_COLUMN not in src.columns:
        raise ValueError(f"У таблиці немає колонки «{ICP_NAME_COLUMN}».")

    icp = src[ICP_NAME_COLUMN]
    out = pd.DataFrame(
        {
            "Activity Subject": build_activity_subject(src),
            "Organization - Research ICP": icp,
        }
    )
    for col in CRM_COMP_COLUMNS[2:]:
        out[col] = src[col] if col in src.columns else ""

    if "Website" in out.columns:
        out["Website"] = out["Website"].apply(normalize_website_key)

    out = normalize_state_columns(out)
    out = normalize_industry_column(out)
    if EMAIL_COLUMN in out.columns:
        out = fill_email_providers(out, overwrite=True)
    out = apply_person_lookup(out, person_lookup)
    out = apply_leads_organization_lookup(out, org_lookup, keep_existing_org_id=True)
    return align_comp_to_crm_schema(out)


def save_comp_export(
    df: pd.DataFrame,
    path: Path | str,
    *,
    person_lookup: PersonLookup | None = None,
    org_lookup: OrganizationLookup | None = None,
) -> Path:
    return save_comp_dataframe(
        build_comp_export(df, person_lookup=person_lookup, org_lookup=org_lookup),
        path,
    )


def save_comp_dataframe(df: pd.DataFrame, path: Path | str) -> Path:
    out_path = Path(path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
    except OSError as exc:
        log_disk_failure(out_path, exc, action="write export CSV")
        raise
    log_disk_write(out_path, action="write export CSV")
    return out_path
