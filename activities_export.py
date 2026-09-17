"""Обробка файлу активностей CRM — матч з дослідженням по Apollo Contact id."""

from __future__ import annotations

import io
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
from dateutil.relativedelta import relativedelta

from excel_utils import excel_cell_value
from linkedin_urls import LINKEDIN_PERSON_COLUMN, linkedin_person_match_key

# Ключ матчу
ACTIVITY_APOLLO_ID_COLUMN = "Person - Apollo Contact id"
RESEARCH_APOLLO_ID_COLUMN = "Person - Apollo Contact id"

# Колонки з файлу активностей
ACTIVITY_EMAIL_COLUMN = "Person - Email - Work"
ACTIVITY_LINKEDIN_COLUMN = "Person - LinkedIn"
ACTIVITY_CONTACT_STARTER_COLUMN = "Person - Contact starter"
ACTIVITY_RESEARCH_ID_COLUMN = "Person - Research ID"
ACTIVITY_RESEARCH_ICP_COLUMN = "Organization - Research ICP"

# Колонки з research CSV
RESEARCH_CONTACT_STARTER_COLUMN = "Researcher Name"
RESEARCH_RESEARCH_ID_COLUMN = "Research ID"
RESEARCH_ICP_COLUMN = "ICP name"
RESEARCH_ICP_FALLBACK_COLUMNS = (
    "ICP name",
    "Organization - Research ICP",
    "Activity Subject",
)
RESEARCH_NOTION_CAMPAIGN_ID_COLUMN = "Notion Campaign ID"
RESEARCH_EMAIL_COLUMN = "Email"
LEADS_EMAIL_COLUMN = "Email"
LEADS_APOLLO_ID_COLUMN = "Person - Apollo Contact id"
PERSON_LINKEDIN_ACTIVE_COLUMN = "Person - Linkedin Active"

# Порядок колонок у вихідному файлі
OUTPUT_COLUMNS = [
    "Person - Email - Work",
    "Person - Contact starter",
    "Person - Contact starter (research)",
    "Person - Research ID",
    "Person - Research ID (research)",
    "Organization - Research ICP",
    "Organization - Research ICP (research)",
    "Activity - ID",
    "Activity - Subject",
    "Person - ID",
    "Person - Labels",
    "Person - Contact Source",
    "Person - Owner",
    "Organization - ID",
    "Organization - Labels",
    "Organization - Source",
    "Organization - Owner",
    "Person - Replied at",
    "Person - Initial answer type",
    "Person - Opted Out",
    "Person - Last activity date",
    "Person - Last email sent",
    "Person - Last email received",
    "Organization - Last activity date",
    "Person - Next activity date",
    "Person - Linkedin Active",
    "Person - Email - Work (2)",
    "Person - LinkedIn",
    "Activity - Due date",
    "Person - SDR Responsible",
]


def _cell_text(value: object) -> str:
    """Текст з клітинки CSV/Excel; pd.NA не можна використовувати в `or`."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    if text.casefold() in {"nan", "<na>", "none"}:
        return ""
    return text


def _normalize_key(value: object) -> str:
    return _cell_text(value).lower()


def build_research_apollo_index(research_df: pd.DataFrame) -> dict[str, pd.Series]:
    """Person - Apollo Contact id → рядок research (перший збіг)."""
    idx: dict[str, pd.Series] = {}
    col = RESEARCH_APOLLO_ID_COLUMN
    if col not in research_df.columns:
        return idx
    for _, row in research_df.iterrows():
        key = _normalize_key(row.get(col, ""))
        if key and key not in idx:
            idx[key] = row
    return idx


def build_research_email_index(research_df: pd.DataFrame) -> dict[str, pd.Series]:
    """Email → рядок research (fallback, якщо в activities немає Apollo id)."""
    idx: dict[str, pd.Series] = {}
    if RESEARCH_EMAIL_COLUMN not in research_df.columns:
        return idx
    for _, row in research_df.iterrows():
        key = _normalize_key(row.get(RESEARCH_EMAIL_COLUMN, ""))
        if key and key not in idx:
            idx[key] = row
    return idx


def build_research_linkedin_index(research_df: pd.DataFrame) -> dict[str, pd.Series]:
    """LinkedIn Person → рядок research (ліди без email)."""
    idx: dict[str, pd.Series] = {}
    col = None
    for name in (LINKEDIN_PERSON_COLUMN, ACTIVITY_LINKEDIN_COLUMN):
        if name in research_df.columns:
            col = name
            break
    if col is None:
        return idx
    for _, row in research_df.iterrows():
        key = linkedin_person_match_key(row.get(col, ""))
        if key and key not in idx:
            idx[key] = row
    return idx


def _research_row_value(row: pd.Series, *columns: str) -> str:
    for col in columns:
        if col not in row.index:
            continue
        val = row.get(col, "")
        if pd.isna(val):
            continue
        text = str(val).strip()
        if text:
            return text
    return ""


def _split_apollo_ids(value: object) -> list[str]:
    """CRM іноді дає кілька Apollo id через кому — розбиваємо на окремі ключі."""
    text = _cell_text(value)
    if not text:
        return []
    return [_normalize_key(part) for part in text.split(",") if str(part).strip()]


def _build_person_id_index(research_df: pd.DataFrame) -> dict[str, pd.Series]:
    """Person - ID → рядок research (перший збіг)."""
    idx: dict[str, pd.Series] = {}
    col = "Person - ID"
    if research_df is None or col not in research_df.columns:
        return idx
    for _, row in research_df.iterrows():
        key = _normalize_key(row.get(col, ""))
        if key and key not in idx:
            idx[key] = row
    return idx


def _resolve_research_row(
    apollo_raw: object,
    email: object,
    person_id: object,
    *,
    apollo_idx: dict[str, pd.Series],
    email_idx: dict[str, pd.Series],
    person_id_idx: dict[str, pd.Series] | None = None,
    linkedin: object = None,
    linkedin_idx: dict[str, pd.Series] | None = None,
) -> pd.Series | None:
    """Рядок research: Apollo id → Email → LinkedIn → Person - ID."""
    person_id_idx = person_id_idx or {}
    linkedin_idx = linkedin_idx or {}
    for part in _split_apollo_ids(apollo_raw):
        row = apollo_idx.get(part)
        if row is not None:
            return row
    row = email_idx.get(_normalize_key(email))
    if row is not None:
        return row
    li_key = linkedin_person_match_key(linkedin)
    if li_key:
        row = linkedin_idx.get(li_key)
        if row is not None:
            return row
    return person_id_idx.get(_normalize_key(person_id))


def _lookup_research_rows(
    activities_df: pd.DataFrame,
    research_df: pd.DataFrame | None,
) -> list[pd.Series | None]:
    """Для кожного рядка activities — рядок research (Apollo → Email → LinkedIn → Person ID)."""
    n = len(activities_df)
    if n == 0 or research_df is None or research_df.empty:
        return [None] * n
    apollo_idx = build_research_apollo_index(research_df)
    email_idx = build_research_email_index(research_df)
    person_id_idx = _build_person_id_index(research_df)
    linkedin_idx = build_research_linkedin_index(research_df)
    apollo_raws = (
        activities_df[ACTIVITY_APOLLO_ID_COLUMN].tolist()
        if ACTIVITY_APOLLO_ID_COLUMN in activities_df.columns
        else [""] * n
    )
    emails = (
        activities_df[ACTIVITY_EMAIL_COLUMN].tolist()
        if ACTIVITY_EMAIL_COLUMN in activities_df.columns
        else [""] * n
    )
    person_ids = (
        activities_df["Person - ID"].tolist()
        if "Person - ID" in activities_df.columns
        else [""] * n
    )
    linkedins = (
        activities_df[ACTIVITY_LINKEDIN_COLUMN].tolist()
        if ACTIVITY_LINKEDIN_COLUMN in activities_df.columns
        else [""] * n
    )
    return [
        _resolve_research_row(
            apollo,
            email,
            pid,
            apollo_idx=apollo_idx,
            email_idx=email_idx,
            person_id_idx=person_id_idx,
            linkedin=linkedin,
            linkedin_idx=linkedin_idx,
        )
        for apollo, email, pid, linkedin in zip(
            apollo_raws, emails, person_ids, linkedins
        )
    ]


def _resolve_activity_apollo_ids(
    activities_df: pd.DataFrame,
    email_idx: dict[str, pd.Series],
    apollo_idx: dict[str, pd.Series] | None = None,
) -> list[str]:
    """Apollo id з activities; кілька id через кому; fallback — Email."""
    apollo_idx = apollo_idx or {}
    apollo_col = ACTIVITY_APOLLO_ID_COLUMN
    email_col = ACTIVITY_EMAIL_COLUMN
    has_apollo_col = apollo_col in activities_df.columns
    emails = (
        activities_df[email_col].tolist()
        if email_col in activities_df.columns
        else [""] * len(activities_df)
    )
    apollo_raw = (
        activities_df[apollo_col].tolist()
        if has_apollo_col
        else [""] * len(activities_df)
    )

    resolved: list[str] = []
    for apollo, email in zip(apollo_raw, emails):
        matched = ""
        for part in _split_apollo_ids(apollo):
            if part in apollo_idx:
                matched = part
                break
        if matched:
            resolved.append(matched)
            continue
        email_key = _normalize_key(email)
        research_row = email_idx.get(email_key)
        if research_row is not None:
            resolved.append(_normalize_key(research_row.get(RESEARCH_APOLLO_ID_COLUMN, "")))
            continue
        parts = _split_apollo_ids(apollo)
        resolved.append(parts[0] if parts else "")
    return resolved


def build_leads_linkedin_active_indexes(
    leads_df: pd.DataFrame,
) -> tuple[dict[str, str], dict[str, str]]:
    """Email / Apollo id → Person - Linkedin Active з Leads for CRM (comp_df)."""
    by_email: dict[str, str] = {}
    by_apollo: dict[str, str] = {}
    if leads_df is None or leads_df.empty:
        return by_email, by_apollo
    if PERSON_LINKEDIN_ACTIVE_COLUMN not in leads_df.columns:
        return by_email, by_apollo
    for _, row in leads_df.iterrows():
        flag = str(row.get(PERSON_LINKEDIN_ACTIVE_COLUMN, "")).strip()
        email = _normalize_key(row.get(LEADS_EMAIL_COLUMN, ""))
        apollo = _normalize_key(row.get(LEADS_APOLLO_ID_COLUMN, ""))
        if email:
            by_email[email] = flag
        if apollo:
            by_apollo[apollo] = flag
    return by_email, by_apollo


def apply_leads_linkedin_active(
    activities_df: pd.DataFrame,
    leads_df: pd.DataFrame | None,
) -> tuple[pd.DataFrame, int]:
    """
    Підставити Person - Linkedin Active з Leads for CRM (оновлені yes/no/порожньо).
    Метч: Apollo Contact id → Email. Повертає (df, кількість оновлених рядків).
    """
    out = activities_df.copy()
    if PERSON_LINKEDIN_ACTIVE_COLUMN not in out.columns:
        out[PERSON_LINKEDIN_ACTIVE_COLUMN] = ""

    if leads_df is None or leads_df.empty:
        return out, 0

    by_email, by_apollo = build_leads_linkedin_active_indexes(leads_df)
    if not by_email and not by_apollo:
        return out, 0

    apollo_ids = _resolve_activity_apollo_ids(
        out,
        build_research_email_index(leads_df),
        build_research_apollo_index(leads_df),
    )
    emails = (
        out[ACTIVITY_EMAIL_COLUMN].tolist()
        if ACTIVITY_EMAIL_COLUMN in out.columns
        else [""] * len(out)
    )

    updated = 0
    flags: list[str] = []
    for i in range(len(out)):
        apollo_id = apollo_ids[i]
        email = emails[i]
        flag = None
        if apollo_id and apollo_id in by_apollo:
            flag = by_apollo[apollo_id]
        else:
            email_key = _normalize_key(email)
            if email_key in by_email:
                flag = by_email[email_key]
        if flag is not None:
            flags.append(flag)
            updated += 1
        else:
            flags.append(str(out.iloc[i][PERSON_LINKEDIN_ACTIVE_COLUMN]).strip())

    out[PERSON_LINKEDIN_ACTIVE_COLUMN] = flags
    return out, updated


def merge_activities_with_research(
    activities_df: pd.DataFrame,
    research_df: pd.DataFrame | None,
    *,
    leads_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """
    Зливає activities з research: Apollo Contact id → Email → LinkedIn → Person - ID.
    Повертає DataFrame у форматі вихідного листа.
    """
    out = activities_df.copy()

    apollo_idx: dict[str, pd.Series] = {}
    email_idx: dict[str, pd.Series] = {}
    person_id_idx: dict[str, pd.Series] = {}
    linkedin_idx: dict[str, pd.Series] = {}
    if research_df is not None and not research_df.empty:
        apollo_idx = build_research_apollo_index(research_df)
        email_idx = build_research_email_index(research_df)
        person_id_idx = _build_person_id_index(research_df)
        linkedin_idx = build_research_linkedin_index(research_df)

    apollo_col = ACTIVITY_APOLLO_ID_COLUMN
    email_col = ACTIVITY_EMAIL_COLUMN
    person_col = "Person - ID"
    apollo_raws = (
        out[apollo_col].tolist()
        if apollo_col in out.columns
        else [""] * len(out)
    )
    emails = (
        out[email_col].tolist()
        if email_col in out.columns
        else [""] * len(out)
    )
    person_ids = (
        out[person_col].tolist()
        if person_col in out.columns
        else [""] * len(out)
    )
    linkedins = (
        out[ACTIVITY_LINKEDIN_COLUMN].tolist()
        if ACTIVITY_LINKEDIN_COLUMN in out.columns
        else [""] * len(out)
    )

    def _research_cols(*columns: str) -> list[str]:
        values: list[str] = []
        for apollo_raw, email, person_id, linkedin in zip(
            apollo_raws, emails, person_ids, linkedins
        ):
            row = _resolve_research_row(
                apollo_raw,
                email,
                person_id,
                apollo_idx=apollo_idx,
                email_idx=email_idx,
                person_id_idx=person_id_idx,
                linkedin=linkedin,
                linkedin_idx=linkedin_idx,
            )
            values.append(
                _research_row_value(row, *columns) if row is not None else ""
            )
        return values

    out["Person - Contact starter (research)"] = _research_cols(
        RESEARCH_CONTACT_STARTER_COLUMN
    )
    out["Person - Research ID (research)"] = _research_cols(RESEARCH_RESEARCH_ID_COLUMN)
    out["Organization - Research ICP (research)"] = _research_cols(
        *RESEARCH_ICP_FALLBACK_COLUMNS
    )
    out["Notion Campaign ID (research)"] = _research_cols(
        RESEARCH_NOTION_CAMPAIGN_ID_COLUMN
    )

    # Друга копія Email - Work
    email_col = ACTIVITY_EMAIL_COLUMN
    out["Person - Email - Work (2)"] = out[email_col] if email_col in out.columns else ""

    # Вирівнюємо колонки
    result = pd.DataFrame()
    for col in OUTPUT_COLUMNS:
        if col in out.columns:
            result[col] = out[col].values
        else:
            result[col] = ""

    if leads_df is not None and not leads_df.empty:
        result, _ = apply_leads_linkedin_active(result, leads_df)

    return result


# Колонки з activities (виділяємо червоним)
RED_COLUMNS = {
    "Person - Contact starter",
    "Person - Research ID",
    "Organization - Research ICP",
}
RED_HEADER_FILL = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")


def merged_to_excel_bytes(merged_df: pd.DataFrame) -> bytes:
    """Повертає xlsx-байти з підсвіченими колонками з activities."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Activities"

    cols = list(merged_df.columns)

    # Заголовки
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        if col_name in RED_COLUMNS:
            cell.fill = RED_HEADER_FILL
        cell.font = openpyxl.styles.Font(bold=True)

    # Дані з підсвіткою проблемних клітинок (всі рядки)
    _, cell_flags = _build_issue_mask(merged_df)
    for r_idx, orig_idx in enumerate(merged_df.index, start=2):
        row = merged_df.loc[orig_idx]
        for c_idx, col_name in enumerate(cols, start=1):
            value = row[col_name]
            cell = ws.cell(row=r_idx, column=c_idx, value=(None if pd.isna(value) else value))
            flag_series = cell_flags.get(col_name)
            if flag_series is not None and flag_series.loc[orig_idx]:
                cell.fill = ISSUE_FILL

    # Ширина колонок
    for c_idx, col_name in enumerate(cols, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c_idx)].width = max(
            15, min(len(col_name) + 2, 40)
        )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


ISSUE_FILL = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")


def _is_filled(value: object) -> bool:
    """Повертає True, якщо клітинка непорожня."""
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except (TypeError, ValueError):
        pass
    return str(value).strip() != ""


def _build_issue_mask(merged_df: pd.DataFrame) -> tuple[pd.Series, dict[str, pd.Series]]:
    """
    Повертає:
      - row_mask: булева серія — True для рядків з хоча б однією проблемою
      - cell_flags: {col_name: булева серія} — які клітинки підсвічувати
    """
    cell_flags: dict[str, pd.Series] = {}

    # Правило 1: Person - Initial answer type заповнено
    col = "Person - Initial answer type"
    if col in merged_df.columns:
        cell_flags[col] = merged_df[col].apply(_is_filled)

    # Правило 2: Person - Opted Out заповнено
    col = "Person - Opted Out"
    if col in merged_df.columns:
        cell_flags[col] = merged_df[col].apply(_is_filled)

    # Правило 3: Person - Last activity date — дата ≤ (сьогодні − 3 місяці)
    col = "Person - Last activity date"
    if col in merged_df.columns:
        cutoff = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=3)
        parsed = pd.to_datetime(merged_df[col], dayfirst=False, errors="coerce")
        cell_flags[col] = parsed.notna() & (parsed >= cutoff)

    # Правила 3б-3г: ті самі умови що на Person - Last activity date
    for _date_col in (
        "Person - Last email sent",
        "Person - Last email received",
        "Organization - Last activity date",
    ):
        if _date_col in merged_df.columns:
            cutoff = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - relativedelta(months=3)
            parsed = pd.to_datetime(merged_df[_date_col], dayfirst=False, errors="coerce")
            cell_flags[_date_col] = parsed.notna() & (parsed >= cutoff)

    # Правило 4: Person - Labels — непорожньо і не Prospect
    col = "Person - Labels"
    if col in merged_df.columns:
        raw = merged_df[col]
        is_filled = raw.notna() & raw.astype(str).str.strip().str.lower().ne("nan") & raw.astype(str).str.strip().str.len().gt(0)
        is_allowed = raw.astype(str).str.contains("Prospect", case=False, na=False, regex=False)
        cell_flags[col] = is_filled & ~is_allowed

    # Правило 5: Person - Contact Source — не пусто і не Research/Leadgen
    col = "Person - Contact Source"
    if col in merged_df.columns:
        raw = merged_df[col]
        is_filled = raw.notna() & raw.astype(str).str.strip().str.lower().ne("nan") & raw.astype(str).str.strip().str.len().gt(0)
        is_allowed = raw.astype(str).str.contains("Research/Leadgen", case=False, na=False, regex=False)
        cell_flags[col] = is_filled & ~is_allowed

    # Правило 6: Person - Owner — не пусто і не Researcher / Daria
    col = "Person - Owner"
    if col in merged_df.columns:
        raw = merged_df[col]
        is_filled = raw.notna() & raw.astype(str).str.strip().str.lower().ne("nan") & raw.astype(str).str.strip().str.len().gt(0)
        is_allowed = (
            raw.astype(str).str.contains("Researcher", case=False, na=False, regex=False)
            | raw.astype(str).str.contains("Daria", case=False, na=False, regex=False)
        )
        cell_flags[col] = is_filled & ~is_allowed

    # Правило 7: Organization - Labels — непорожньо і не Prospect
    # Виключення: Disqualified + ICP містить "Re-enriched & Disqualified" → OK
    col = "Organization - Labels"
    icp_col = "Organization - Research ICP"
    if col in merged_df.columns:
        raw = merged_df[col]
        is_filled = raw.notna() & raw.astype(str).str.strip().str.lower().ne("nan") & raw.astype(str).str.strip().str.len().gt(0)
        is_prospect = raw.astype(str).str.contains("Prospect", case=False, na=False, regex=False)
        icp_series = merged_df[icp_col].astype(str) if icp_col in merged_df.columns else pd.Series([""] * len(merged_df))
        is_re_enriched_ok = (
            raw.astype(str).str.contains("Disqualified", case=False, na=False, regex=False)
            & icp_series.str.contains("Re-enriched & Disqualified", case=False, na=False, regex=False)
        )
        cell_flags[col] = is_filled & ~is_prospect & ~is_re_enriched_ok

    # Правило 8: Organization - Owner — не пусто і не Daria / Maryna Chut / Researcher
    col = "Organization - Owner"
    if col in merged_df.columns:
        raw = merged_df[col]
        is_filled = raw.notna() & raw.astype(str).str.strip().str.lower().ne("nan") & raw.astype(str).str.strip().str.len().gt(0)
        is_allowed = (
            raw.astype(str).str.contains("Daria", case=False, na=False, regex=False)
            | raw.astype(str).str.contains("Maryna Chut", case=False, na=False, regex=False)
            | raw.astype(str).str.contains("Researcher", case=False, na=False, regex=False)
        )
        cell_flags[col] = is_filled & ~is_allowed

    # Правило 9: Person - Replied at — є значення
    col = "Person - Replied at"
    if col in merged_df.columns:
        cell_flags[col] = merged_df[col].apply(_is_filled)

    # Правило 10: Person - Next activity date — відрізняється від найпоширенішої дати
    col = "Person - Next activity date"
    if col in merged_df.columns:
        parsed = pd.to_datetime(merged_df[col], dayfirst=False, errors="coerce")
        normalized = parsed.dt.normalize()
        mode_series = normalized.dropna().mode()
        if not mode_series.empty:
            dominant_date = mode_series.iloc[0]
            cell_flags[col] = parsed.notna() & (normalized != dominant_date)
        else:
            cell_flags[col] = pd.Series([False] * len(merged_df), index=merged_df.index)

    # Правило: Organization - Source — непорожньо і не Leadgen/Researcher
    col = "Organization - Source"
    if col in merged_df.columns:
        raw = merged_df[col]
        is_filled = raw.notna() & raw.astype(str).str.strip().str.lower().ne("nan") & raw.astype(str).str.strip().str.len().gt(0)
        is_allowed = raw.astype(str).str.contains("Leadgen/Researcher", case=False, na=False, regex=False)
        cell_flags[col] = is_filled & ~is_allowed

    # --- тут додаватимуться нові правила ---

    if not cell_flags:
        return pd.Series([False] * len(merged_df)), {}

    row_mask = pd.concat(cell_flags.values(), axis=1).any(axis=1)
    return row_mask, cell_flags


def issues_to_excel_bytes(merged_df: pd.DataFrame) -> bytes:
    """
    Excel-файл лише з проблемними рядками.
    Підсвічує червоним конкретні клітинки з проблемами.
    """
    row_mask, cell_flags = _build_issue_mask(merged_df)
    # Зберігаємо оригінальний індекс — потрібен для правильного зіставлення cell_flags
    issues_df = merged_df[row_mask]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Issues"

    cols = list(issues_df.columns)

    # Заголовки
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        if col_name in RED_COLUMNS:
            cell.fill = RED_HEADER_FILL
        cell.font = openpyxl.styles.Font(bold=True)

    # Дані з підсвіткою — кожен рядок унікальний, кілька підсвіток в одному рядку OK
    for r_idx, orig_idx in enumerate(issues_df.index, start=2):
        row = issues_df.loc[orig_idx]
        for c_idx, col_name in enumerate(cols, start=1):
            value = row[col_name]
            cell = ws.cell(row=r_idx, column=c_idx, value=(None if pd.isna(value) else value))
            flag_series = cell_flags.get(col_name)
            if flag_series is not None and flag_series.loc[orig_idx]:
                cell.fill = ISSUE_FILL

    # Ширина колонок
    for c_idx, col_name in enumerate(cols, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c_idx)].width = max(
            15, min(len(col_name) + 2, 40)
        )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


FINAL_COLUMNS = [
    "Activity - ID",
    "Activity - Subject",
    "Activity - Due date",
    "Person - ID",
    "Person - Contact starter",
    "Person - Research ID",
    "Person - SDR Responsible",
    "Person - Labels",
    "Person - Contact Source",
    "Person - Owner",
    "Organization - ID",
    "Organization - Research ICP",
    "Organization - Labels",
    "Organization - Source",
    "Organization - Owner",
    "Person - Replied at",
    "Person - Initial answer type",
]


def build_final_export(
    filtered_df: pd.DataFrame,
    research_df: pd.DataFrame | None = None,
    *,
    include_linkedin_active: bool = False,
) -> pd.DataFrame:
    """Формує фінальний DataFrame з 17 колонками для заливки в CRM.
    Research: колонки (research) у merged, інакше Apollo → Email → LinkedIn → Person ID.
    Опційно додає Person - Linkedin Active після Person - Research ID.
    """
    today = datetime.now().strftime("%Y-%m-%d")
    n = len(filtered_df)
    research_rows = _lookup_research_rows(filtered_df, research_df)

    def _from_research(index: int, *columns: str) -> str:
        row = research_rows[index] if index < len(research_rows) else None
        return _research_row_value(row, *columns) if row is not None else ""

    result = pd.DataFrame()
    for col in FINAL_COLUMNS:
        if col == "Activity - Due date":
            result[col] = [today] * n
        elif col == "Person - Contact starter":
            research_vals = (
                filtered_df["Person - Contact starter (research)"]
                .fillna("")
                .astype(str)
                .str.strip()
                .tolist()
                if "Person - Contact starter (research)" in filtered_df.columns
                else [""] * n
            )
            act_vals = (
                filtered_df["Person - Contact starter"]
                .fillna("")
                .astype(str)
                .str.strip()
                .tolist()
                if "Person - Contact starter" in filtered_df.columns
                else [""] * n
            )
            combined_starter: list[str] = []
            for i, (rv, av) in enumerate(zip(research_vals, act_vals)):
                combined_starter.append(
                    rv or _from_research(i, RESEARCH_CONTACT_STARTER_COLUMN) or av
                )
            result[col] = combined_starter
        elif col == "Person - Research ID":
            act_ids = filtered_df["Person - Research ID"].fillna("").astype(str).str.strip().tolist() if "Person - Research ID" in filtered_df.columns else [""] * n
            notion_ids = filtered_df["Person - Research ID (research)"].fillna("").astype(str).str.strip().tolist() if "Person - Research ID (research)" in filtered_df.columns else [""] * n
            combined = []
            for i, (act, notion) in enumerate(zip(act_ids, notion_ids)):
                if not notion:
                    notion = _from_research(i, RESEARCH_RESEARCH_ID_COLUMN)
                if notion and act:
                    # Нове значення (research) — першим, далі попередні з activities.
                    parts = [p.strip() for p in f"{notion},{act}".split(",") if p.strip()]
                    seen = []
                    for p in parts:
                        if p not in seen:
                            seen.append(p)
                    combined.append(", ".join(seen))
                elif notion:
                    combined.append(notion)
                else:
                    combined.append(act)
            result[col] = combined
        elif col == "Person - Labels":
            result[col] = ["Prospect"] * n
        elif col == "Person - Contact Source":
            result[col] = ["Research/Leadgen"] * n
        elif col == "Person - Owner":
            result[col] = ["Researcher"] * n
        elif col == "Organization - Labels":
            result[col] = ["Prospect"] * n
        elif col == "Organization - Source":
            result[col] = ["Leadgen/Researcher"] * n
        elif col == "Organization - Owner":
            result[col] = ["Researcher"] * n
        elif col in ("Person - Replied at", "Person - Initial answer type"):
            result[col] = [""] * n
        elif col == "Organization - Research ICP":
            icp_vals = (
                filtered_df["Organization - Research ICP"]
                .fillna("")
                .astype(str)
                .str.strip()
                .tolist()
                if "Organization - Research ICP" in filtered_df.columns
                else [""] * n
            )
            # Для івентів Activity - Subject = Conference; в ICP додаємо ICP name (Events), не Conference.
            research_icp_src = "Organization - Research ICP (research)"
            if research_icp_src in filtered_df.columns:
                icp_name_vals = (
                    filtered_df[research_icp_src].fillna("").astype(str).str.strip().tolist()
                )
            else:
                icp_name_vals = [""] * n
            subj_vals = (
                filtered_df["Activity - Subject"]
                .fillna("")
                .astype(str)
                .str.strip()
                .tolist()
                if "Activity - Subject" in filtered_df.columns
                else [""] * n
            )
            combined_icp = []
            for i, (icp, icp_name, subj) in enumerate(zip(icp_vals, icp_name_vals, subj_vals)):
                if not icp_name:
                    icp_name = _from_research(i, *RESEARCH_ICP_FALLBACK_COLUMNS)
                add = icp_name if icp_name else subj
                # Нове значення (ICP name) — першим, далі попередні з activities.
                parts = [p.strip() for p in f"{add},{icp}".split(",") if p.strip()]
                seen = []
                for p in parts:
                    if p not in seen and p.lower() != "re-enriched & disqualified":
                        seen.append(p)
                combined_icp.append(", ".join(seen))
            result[col] = combined_icp
        elif col == "Person - SDR Responsible":
            result[col] = ["Daria H."] * n
        else:
            result[col] = filtered_df[col].values if col in filtered_df.columns else [""] * n

    if include_linkedin_active:
        col = PERSON_LINKEDIN_ACTIVE_COLUMN
        vals = (
            filtered_df[col].fillna("").astype(str).str.strip().tolist()
            if col in filtered_df.columns
            else [""] * n
        )
        insert_at = int(result.columns.get_loc("Person - Research ID")) + 1
        result.insert(insert_at, col, vals)

    return result


def final_to_excel_bytes(
    filtered_df: pd.DataFrame,
    research_df: pd.DataFrame | None = None,
    *,
    include_linkedin_active: bool = False,
) -> bytes:
    """Excel-файл фінального експорту для CRM (17 колонок, без підсвітки)."""
    df = build_final_export(
        filtered_df,
        research_df,
        include_linkedin_active=include_linkedin_active,
    )
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Final"

    cols = list(df.columns)
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.font = openpyxl.styles.Font(bold=True)

    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, (col_name, value) in enumerate(zip(cols, row), start=1):
            ws.cell(row=r_idx, column=c_idx, value=excel_cell_value(col_name, value))

    for c_idx, col_name in enumerate(cols, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c_idx)].width = max(
            15, min(len(col_name) + 2, 40)
        )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


ACTIVITIES_DELETION_COLUMNS = [
    "Activity - Subject",
    "Activity - type",
    "Activity - ID",
    "Person - ID",
    "Person - Contact Source",
    "Person - Owner",
    "Organization - ID",
    "Organization - Source",
    "Organization - Owner",
]

DEFAULT_ACTIVITIES_DELETION_TYPE = "email campaign"
DEFAULT_PERSON_CONTACT_SOURCE = "Research/Leadgen"
DEFAULT_ORGANIZATION_SOURCE = "Leadgen/Researcher"

# Альтернативні назви Activity - type у вихідному activities CSV
ACTIVITY_TYPE_SOURCE_COLUMNS = (
    "Activity - type",
    "Activity Type",
    "Activity - Type",
)


RESEARCHER_NAME_COLUMN = "Researcher Name"


def is_activities_work_file(df: pd.DataFrame) -> bool:
    """Робочий del-файл: рядки activities + Researcher Name / ICP для видалення."""
    if df is None or df.empty:
        return False
    if "Activity - ID" not in df.columns or "Activity - Subject" not in df.columns:
        return False
    return RESEARCHER_NAME_COLUMN in df.columns or any(
        str(col).startswith("Organization - Research ICP") for col in df.columns
    )


def normalize_activities_work_file(df: pd.DataFrame) -> pd.DataFrame:
    """Приводить робочий CSV/Excel до очікуваних назв колонок."""
    out = df.copy()
    rename: dict[str, str] = {}
    icp_seen = 0
    new_columns: list[str] = []
    for col in out.columns:
        text = str(col).strip()
        if text.startswith("Organization - Research ICP"):
            if icp_seen == 0:
                new_columns.append("Organization - Research ICP")
            else:
                new_columns.append(f"Organization - Research ICP ({icp_seen + 1})")
            icp_seen += 1
            continue
        new_columns.append(text)
    out.columns = new_columns
    return out.fillna("").astype(str)


def _deletion_owner(value: object) -> str:
    """Person / Organization - Owner у del-файлі: порожньо або Daria → Researcher."""
    text = _cell_text(value)
    if not text:
        return "Researcher"
    first = text.casefold().split()[0]
    if first == "daria":
        return "Researcher"
    return text


def build_activities_deletion_from_work(work_df: pd.DataFrame) -> pd.DataFrame:
    """Del-файл з робочого файлу (як Google gid=1359526139)."""
    work = normalize_activities_work_file(work_df)
    if work.empty:
        return pd.DataFrame(columns=ACTIVITIES_DELETION_COLUMNS)

    activity_type_col = next(
        (col for col in ACTIVITY_TYPE_SOURCE_COLUMNS if col in work.columns),
        None,
    )
    activity_type_default = DEFAULT_ACTIVITIES_DELETION_TYPE

    rows: list[dict[str, str]] = []
    for _, row in work.iterrows():
        if activity_type_col:
            activity_type = str(row.get(activity_type_col, "")).strip() or activity_type_default
        else:
            activity_type = activity_type_default
        rows.append(
            {
                "Activity - type": activity_type,
                "Activity - Subject": str(row.get("Activity - Subject", "")).strip(),
                "Activity - ID": str(row.get("Activity - ID", "")).strip(),
                "Person - ID": str(row.get("Person - ID", "")).strip(),
                "Person - Contact Source": DEFAULT_PERSON_CONTACT_SOURCE,
                "Person - Owner": _deletion_owner(row.get("Person - Owner", "")),
                "Organization - ID": str(row.get("Organization - ID", "")).strip(),
                "Organization - Source": DEFAULT_ORGANIZATION_SOURCE,
                "Organization - Owner": _deletion_owner(row.get("Organization - Owner", "")),
            }
        )
    return pd.DataFrame(rows, columns=ACTIVITIES_DELETION_COLUMNS)


def _merged_row_to_deletion(row: pd.Series) -> dict[str, str]:
    """Fallback: merged activities → del (без build_final_export)."""
    activity_type = DEFAULT_ACTIVITIES_DELETION_TYPE
    for col in ACTIVITY_TYPE_SOURCE_COLUMNS:
        val = str(row.get(col, "")).strip()
        if val:
            activity_type = val
            break
    return {
        "Activity - type": activity_type,
        "Activity - Subject": str(row.get("Activity - Subject", "")).strip(),
        "Activity - ID": str(row.get("Activity - ID", "")).strip(),
        "Person - ID": str(row.get("Person - ID", "")).strip(),
        "Person - Contact Source": DEFAULT_PERSON_CONTACT_SOURCE,
        "Person - Owner": _deletion_owner(row.get("Person - Owner", "")),
        "Organization - ID": str(row.get("Organization - ID", "")).strip(),
        "Organization - Source": DEFAULT_ORGANIZATION_SOURCE,
        "Organization - Owner": _deletion_owner(row.get("Organization - Owner", "")),
    }


def build_activities_deletion_export(
    excluded_df: pd.DataFrame,
    research_df: pd.DataFrame | None = None,
    *,
    delete_source_df: pd.DataFrame | None = None,
    default_activity_type: str = DEFAULT_ACTIVITIES_DELETION_TYPE,
) -> pd.DataFrame:
    """Файл для видалення activities у CRM.

    Якщо завантажено робочий del-файл — беремо дані напряму з нього.
    Інакше — з рядків merged activities.
    """
    source = delete_source_df if delete_source_df is not None else excluded_df
    if is_activities_work_file(source):
        return build_activities_deletion_from_work(source)

    if excluded_df is None or excluded_df.empty:
        return pd.DataFrame(columns=ACTIVITIES_DELETION_COLUMNS)

    rows = [_merged_row_to_deletion(excluded_df.iloc[i]) for i in range(len(excluded_df))]
    return pd.DataFrame(rows, columns=ACTIVITIES_DELETION_COLUMNS)


def activities_deletion_to_excel_bytes(
    excluded_df: pd.DataFrame,
    research_df: pd.DataFrame | None = None,
    *,
    delete_source_df: pd.DataFrame | None = None,
    default_activity_type: str = DEFAULT_ACTIVITIES_DELETION_TYPE,
) -> bytes:
    """Excel для видалення activities."""
    df = build_activities_deletion_export(
        excluded_df,
        research_df,
        delete_source_df=delete_source_df,
        default_activity_type=default_activity_type,
    )
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Delete"

    cols = list(df.columns)
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.font = openpyxl.styles.Font(bold=True)

    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, (col_name, value) in enumerate(zip(cols, row), start=1):
            ws.cell(row=r_idx, column=c_idx, value=excel_cell_value(col_name, value))

    for c_idx, col_name in enumerate(cols, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c_idx)].width = max(
            15, min(len(col_name) + 2, 40)
        )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def exclude_activities(
    merged_df: pd.DataFrame,
    exclude_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, int, pd.DataFrame | None]:
    """
    Виключення з merged activities.
    Повертає (залишок, виключені з merged, кількість, робочий df для del-експорту).
    """
    if is_activities_work_file(exclude_df):
        work = normalize_activities_work_file(exclude_df)
        delete_source = work.copy()
        activity_ids = {
            _normalize_key(v)
            for v in work.get("Activity - ID", pd.Series(dtype=str)).tolist()
            if _normalize_key(v)
        }
        if "Activity - ID" in merged_df.columns and activity_ids:
            mask = (
                merged_df["Activity - ID"]
                .astype(str)
                .str.strip()
                .str.lower()
                .isin(activity_ids)
            )
        else:
            person_ids = {
                _normalize_key(v)
                for v in work.get("Person - ID", pd.Series(dtype=str)).tolist()
                if _normalize_key(v)
            }
            mask = (
                merged_df["Person - ID"]
                .astype(str)
                .str.strip()
                .str.lower()
                .isin(person_ids)
                if person_ids and "Person - ID" in merged_df.columns
                else pd.Series([False] * len(merged_df), index=merged_df.index)
            )
        excluded = merged_df[mask].copy()
        kept = merged_df[~mask]
        return kept, excluded, int(mask.sum()), delete_source

    kept, excluded, count = exclude_by_person_id(merged_df, exclude_df)
    return kept, excluded, count, None


def exclude_by_person_id(
    merged_df: pd.DataFrame,
    exclude_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, int]:
    """
    Розділяє merged_df на залишок і виключених по Person - ID.
    Повертає (залишок, виключені рядки, кількість виключених).
    """
    id_col = "Person - ID"
    if id_col not in exclude_df.columns:
        # Беремо першу колонку як Person - ID якщо заголовок не збігається
        exclude_df = exclude_df.copy()
        exclude_df.columns = [id_col] + list(exclude_df.columns[1:])

    exclude_ids = set(
        exclude_df[id_col].dropna().astype(str).str.strip().str.lower()
    )
    if not exclude_ids:
        return merged_df, merged_df.iloc[0:0].copy(), 0

    if id_col not in merged_df.columns:
        return merged_df, merged_df.iloc[0:0].copy(), 0

    mask_exclude = merged_df[id_col].astype(str).str.strip().str.lower().isin(exclude_ids)
    excluded = merged_df[mask_exclude].copy()
    filtered = merged_df[~mask_exclude]
    return filtered, excluded, int(mask_exclude.sum())


def matched_stats(merged_df: pd.DataFrame) -> dict[str, int]:
    total = len(merged_df)
    contact_starter_matched = int(
        (merged_df.get("Person - Contact starter (research)", pd.Series([""] * total))
         .astype(str).str.strip() != "").sum()
    )
    research_id_matched = int(
        (merged_df.get("Person - Research ID (research)", pd.Series([""] * total))
         .astype(str).str.strip() != "").sum()
    )
    return {
        "total": total,
        "contact_starter_matched": contact_starter_matched,
        "research_id_matched": research_id_matched,
    }
