"""
Streamlit: робота з CRM (без API) — підготовка даних з Google Таблиці / CSV та файли CRM.
"""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import pandas as pd
import streamlit as st

from company_export import (
    APOLLO_ACCOUNT_ID_COLUMN,
    COMPANY_COLUMN,
    DUPLICATE_NAME_COLUMN,
    DUPLICATES_MERGE_COLUMNS,
    PIPEDRIVE_ORG_ID_COLUMN,
    align_companies_to_crm_schema,
    annotate_duplicate_company_names,
    build_company_export,
    build_duplicate_name_report,
    duplicate_company_name_keys,
    duplicate_company_rows,
    format_merge_changes_summary,
    merge_companies_by_apollo_id,
    normalize_merge_changes_df,
    companies_for_crm_file,
    companies_review_export,
    companies_review_to_xlsx_bytes,
    companies_review_xlsx_columns,
    organization_review_stats,
    prepare_companies_review,
    organization_id_stats,
    ORG_HIGHLIGHT_COLUMN,
    ORG_LOOKUP_ISSUE_COLUMN,
    organization_problem_rows,
    save_company_export,
)
from crm_export import (
    align_comp_to_crm_schema,
    align_ppl_leads_export,
    build_comp_export,
    build_ppl_leads_export,
    is_ppl_leads_template,
    person_id_stats,
    person_leads_review_export,
    person_leads_review_to_xlsx_bytes,
    person_leads_review_stats,
    prepare_person_leads_review,
    person_lookup_issue_label,
    person_problem_rows,
    person_row_is_highlighted,
    ppl_leads_for_crm_file,
    ACTIVITY_TYPE_COLUMN,
    DEFAULT_ACTIVITY_TYPE,
    NOTION_CAMPAIGN_ID_COLUMN,
    prepare_ppl_leads_export,
    save_comp_dataframe,
)
from person_lookup import (
    CRM_BY_EMAIL_ID_COLUMN,
    CRM_LINKEDIN_ID_COLUMN as PERSON_CRM_LINKEDIN_COLUMN,
    ID_BY_LINKEDIN_PERSON_COLUMN,
    PERSON_ID_BY_EMAIL_COLUMN,
    PERSON_ID_COLUMN,
    PERSON_HIGHLIGHT_COLUMN,
    PERSON_LOOKUP_ISSUE_COLUMN,
    PersonLookup,
    apply_person_lookup,
    format_leads_person_columns_help,
    format_person_csv_help,
    is_persons_crm_export,
)
from activities_export import matched_stats, merge_activities_with_research, merged_to_excel_bytes, issues_to_excel_bytes, exclude_activities, build_activities_deletion_export, activities_deletion_to_excel_bytes, build_final_export, final_to_excel_bytes, apply_leads_linkedin_active
from mx_checker import EMAIL_COLUMN, fill_email_providers
from ids_runtime import ensure_ops_started
from linkedin_activity import (
    PERSON_LINKEDIN_ACTIVE_COLUMN,
    merge_linkedin_active_into_comp,
    merge_linkedin_active_into_df,
)
from linkedin_urls import (
    ICP_NAME_EXPORT_COLUMN,
    LINKEDIN_PERSON_COLUMN,
    build_linkedin_export,
    save_linkedin_export,
)
from organization_lookup import (
    CRM_LINKEDIN_ID_COLUMN,
    CRM_WEBSITE_ID_COLUMN,
    LINKEDIN_ID_COLUMN,
    OrganizationLookup,
    WEBSITE_ID_COLUMN,
    apply_leads_organization_lookup,
    apply_organization_lookup,
    format_company_export_help,
    format_organization_csv_help,
    format_organizations_crm_workflow_help,
    leads_have_org_lookup_columns,
)
from session_persistence import (
    apply_workspace,
    clear_organizations_crm,
    clear_persons_crm,
    clear_workspace,
    clear_export_artifacts,
    load_workspace,
    organizations_crm_exists,
    organizations_crm_path,
    persons_crm_exists,
    persons_crm_path,
    save_organizations_crm,
    save_persons_crm,
    save_workspace,
    save_workspace_frames,
    save_workspace_meta,
    workspace_exists,
)
from research_duplicates import (
    RESEARCH_DUPLICATE_EMAIL_COLUMN,
    RESEARCH_DUPLICATE_LINKEDIN_COLUMN,
    annotate_research_duplicates,
    build_research_duplicates_report,
    research_duplicate_rows,
    research_duplicate_stats,
)
from sheet_source import (
    DEFAULT_SPREADSHEET_URL,
    dataframe_to_csv_bytes,
    has_crm_status_column,
    load_all_sheets_from_urls,
    load_sheet_csv,
    load_source_csv,
    parse_spreadsheet_urls,
    CRM_STATUS_FILTER_ALL,
    CRM_STATUS_FILTER_PENDING,
    CRM_STATUS_FILTER_PENDING_AND_VERIFICATION,
    CRM_STATUS_FILTER_VERIFICATION,
    crm_status_filter_label,
    select_rows_for_export,
)

ROOT = Path(__file__).resolve().parent
ensure_ops_started(root=ROOT)
DEFAULT_FILE_STEM = datetime.now().strftime("%m_%d")
VALID_TABS = ("companies", "leads", "activities")
DEFAULT_ACTIVE_TAB = "companies"


def _export_paths(stem: str) -> tuple[Path, ...]:
    base = stem.strip() or DEFAULT_FILE_STEM
    return (
        ROOT / f"{base} - comp.csv",
        ROOT / f"{base} - linkedin.csv",
        ROOT / f"{base} - companies.csv",
        ROOT / f"{base} - companies-duplicates.csv",
        ROOT / f"{base} - companies-org-issues.csv",
        ROOT / f"{base} - companies-review.xlsx",
        ROOT / f"{base} - leads-person-issues.csv",
        ROOT / f"{base} - leads-person-issues.xlsx",
        ROOT / f"{base} - research-duplicates.csv",
    )


@st.cache_resource(show_spinner=False)
def _cached_org_lookup(csv_path: str, mtime_ns: int) -> OrganizationLookup:
    return OrganizationLookup.from_csv(csv_path)


def _get_org_lookup() -> OrganizationLookup | None:
    path = organizations_crm_path(ROOT)
    if not path.is_file():
        return None
    try:
        return _cached_org_lookup(str(path), path.stat().st_mtime_ns)
    except (OSError, ValueError) as exc:
        st.session_state.organizations_load_error = str(exc)
        return None


def _fill_notion_campaign_enabled() -> bool:
    return bool(st.session_state.get("fill_notion_campaign_id", False))


def _leads_default_activity_type() -> str:
    raw = st.session_state.get("leads_default_activity_type", DEFAULT_ACTIVITY_TYPE)
    value = str(raw).strip()
    return value or DEFAULT_ACTIVITY_TYPE


def _leads_export_kwargs() -> dict:
    return {
        "fill_notion_campaign_id": _fill_notion_campaign_enabled(),
        "default_activity_type": _leads_default_activity_type(),
    }


def _leads_export_source_df() -> pd.DataFrame | None:
    pending = st.session_state.pending_df
    if pending is not None and not pending.empty:
        return pending
    comp = st.session_state.comp_df
    if comp is not None and not comp.empty:
        return comp
    return None


def _rebuild_leads_export() -> None:
    src = _leads_export_source_df()
    if src is None:
        return
    st.session_state.comp_df = build_ppl_leads_export(
        src,
        person_lookup=_get_person_lookup(),
        org_lookup=_get_org_lookup(),
        **_leads_export_kwargs(),
    )
    st.session_state.linkedin_df = build_linkedin_export(src)
    _invalidate_leads_review_cache()


def _on_leads_export_options_change() -> None:
    _rebuild_leads_export()
    _persist_workspace()


def _rebuild_companies_with_org_lookup(*, persist: bool = True) -> None:
    """Перерахувати ID організацій після завантаження бази CRM."""
    lookup = _get_org_lookup()
    pending = st.session_state.pending_df
    if pending is not None and not pending.empty:
        st.session_state.companies_df = build_company_export(pending, org_lookup=lookup)
    else:
        companies = st.session_state.companies_df
        if companies is not None and not companies.empty:
            companies = align_companies_to_crm_schema(companies)
            st.session_state.companies_df = apply_organization_lookup(companies, lookup)
    _invalidate_companies_review_cache()
    _sync_comp_org_ids(force=True, persist=False)
    if persist:
        _persist_workspace_frames("comp_df", "companies_df")


def _sync_comp_org_ids(*, force: bool = False, persist: bool = True) -> bool:
    """Organization - ID на лідах (Comp або PPL): Website → Linkedin Company."""
    comp = st.session_state.comp_df
    if comp is None or comp.empty or not leads_have_org_lookup_columns(comp):
        return False
    lookup = _get_org_lookup()
    if not lookup:
        return False
    col = "Organization - ID"
    if not force and col in comp.columns:
        filled = int((comp[col].astype(str).str.strip() != "").sum())
        if filled > 0:
            return False
    before = int((comp[col].astype(str).str.strip() != "").sum()) if col in comp.columns else 0
    st.session_state.comp_df = prepare_ppl_leads_export(
        comp,
        person_lookup=_get_person_lookup(),
        org_lookup=lookup,
        keep_existing_org_id=not force,
        **_leads_export_kwargs(),
    )
    after = int(
        (st.session_state.comp_df[col].astype(str).str.strip() != "").sum()
    )
    if force or after > before:
        _invalidate_leads_review_cache()
        if persist:
            _persist_workspace()
        return True
    return False


def _sync_companies_org_ids(*, force: bool = False) -> bool:
    """Підставити ID-колонки, якщо база CRM є, а в таблиці їх ще немає."""
    companies = st.session_state.companies_df
    if companies is None or companies.empty:
        return False
    lookup = _get_org_lookup()
    if not lookup:
        return False
    companies = align_companies_to_crm_schema(companies)
    needs = force or WEBSITE_ID_COLUMN not in st.session_state.companies_df.columns
    if not needs and PIPEDRIVE_ORG_ID_COLUMN in companies.columns:
        needs = int(
            (companies[PIPEDRIVE_ORG_ID_COLUMN].astype(str).str.strip() != "").sum()
        ) == 0
    if needs:
        st.session_state.companies_df = apply_organization_lookup(
            companies, lookup, keep_existing_org_id=False
        )
        _invalidate_companies_review_cache()
        _persist_workspace()
        return True
    st.session_state.companies_df = companies
    return False


def _load_organizations_file(file_bytes: bytes, file_name: str) -> None:
    df = load_source_csv(file_bytes)
    lookup = OrganizationLookup.from_dataframe(df)
    save_organizations_crm(ROOT, df)
    _cached_org_lookup.clear()
    st.session_state.crm_organizations_name = file_name
    st.session_state.organizations_row_count = lookup.row_count
    st.session_state.organizations_load_error = None
    _rebuild_companies_with_org_lookup(persist=True)


def _clear_organizations() -> None:
    clear_organizations_crm(ROOT)
    _cached_org_lookup.clear()
    st.session_state.crm_organizations_name = None
    st.session_state.crm_organizations_key = None
    st.session_state.organizations_row_count = None
    st.session_state.organizations_load_error = None
    _persist_workspace_meta()


@st.cache_resource(show_spinner=False)
def _cached_person_lookup(csv_path: str, mtime_ns: int) -> PersonLookup:
    return PersonLookup.from_csv(csv_path)


def _get_person_lookup() -> PersonLookup | None:
    path = persons_crm_path(ROOT)
    if not path.is_file():
        return None
    try:
        return _cached_person_lookup(str(path), path.stat().st_mtime_ns)
    except (OSError, ValueError) as exc:
        st.session_state.persons_load_error = str(exc)
        return None


def _rebuild_comp_with_person_lookup() -> None:
    lookup = _get_person_lookup()
    pending = st.session_state.pending_df
    if pending is not None and not pending.empty:
        _rebuild_leads_export()
        return
    comp = st.session_state.comp_df
    if comp is not None and not comp.empty:
        st.session_state.comp_df = prepare_ppl_leads_export(
            comp,
            person_lookup=lookup,
            org_lookup=_get_org_lookup(),
            **_leads_export_kwargs(),
        )


def _sync_comp_person_ids(*, force: bool = False) -> bool:
    comp = st.session_state.comp_df
    if comp is None or comp.empty:
        return False
    lookup = _get_person_lookup()
    if not lookup:
        return False
    comp = align_ppl_leads_export(comp, issue_column=None)
    needs = force or PERSON_ID_COLUMN not in st.session_state.comp_df.columns
    if not needs and PERSON_ID_COLUMN in comp.columns:
        needs = int((comp[PERSON_ID_COLUMN].astype(str).str.strip() != "").sum()) == 0
    if needs:
        st.session_state.comp_df = prepare_ppl_leads_export(
            comp,
            person_lookup=lookup,
            org_lookup=_get_org_lookup(),
            **_leads_export_kwargs(),
        )
        _invalidate_leads_review_cache()
        _persist_workspace()
        return True
    return False


def _load_persons_file(file_bytes: bytes, file_name: str) -> None:
    df = load_source_csv(file_bytes)
    lookup = PersonLookup.from_dataframe(df)
    save_persons_crm(ROOT, df)
    _cached_person_lookup.clear()
    st.session_state.crm_persons_name = file_name
    st.session_state.persons_row_count = lookup.row_count
    st.session_state.persons_load_error = None
    _rebuild_comp_with_person_lookup()
    _persist_workspace()


def _clear_persons() -> None:
    clear_persons_crm(ROOT)
    _cached_person_lookup.clear()
    st.session_state.crm_persons_name = None
    st.session_state.crm_persons_key = None
    st.session_state.persons_row_count = None
    st.session_state.persons_load_error = None
    _persist_workspace()


def _align_leads_for_display(df: pd.DataFrame) -> pd.DataFrame:
    work = align_ppl_leads_export(df, issue_column=None)
    return ppl_leads_for_crm_file(
        work, include_notion_campaign_id=_fill_notion_campaign_enabled()
    )


def _build_exports(
    full: pd.DataFrame,
    *,
    crm_status_filter: str = CRM_STATUS_FILTER_PENDING,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    pending = select_rows_for_export(full, crm_status_filter=crm_status_filter)
    if pending.empty:
        raise ValueError(
            "Немає рядків для обробки (перевірте фільтр "
            f"«{crm_status_filter_label(crm_status_filter)}»)."
        )
    comp = build_ppl_leads_export(
        pending,
        person_lookup=_get_person_lookup(),
        org_lookup=_get_org_lookup(),
        **_leads_export_kwargs(),
    )
    linkedin_df = build_linkedin_export(pending)
    if is_ppl_leads_template(pending):
        companies_df = pd.DataFrame()
    else:
        companies_df = build_company_export(pending, org_lookup=_get_org_lookup())
    return pending, comp, linkedin_df, companies_df


def _store_research_data(
    full: pd.DataFrame,
    *,
    source_label: str,
    crm_status_filter: str,
) -> None:
    full = annotate_research_duplicates(full)
    with st.spinner(
        "Збірка файлів і MX Provider Checker (домен з Email)…"
    ):
        pending, comp, linkedin_df, companies_df = _build_exports(
            full, crm_status_filter=crm_status_filter
        )
    st.session_state.full_df = full
    st.session_state.pending_df = pending
    st.session_state.comp_df = comp
    st.session_state.linkedin_df = linkedin_df
    st.session_state.companies_df = companies_df
    _invalidate_leads_review_cache()
    _invalidate_companies_review_cache()
    _clear_companies_merge_preview()
    st.session_state.loaded_at = datetime.now(timezone.utc)
    st.session_state.data_source_label = source_label
    st.session_state.crm_status_filter = crm_status_filter
    st.session_state.pending_only = crm_status_filter == CRM_STATUS_FILTER_PENDING
    st.session_state.load_error = None
    _persist_workspace()




def _try_load_google(
    sheet_urls: str | list[str],
    *,
    crm_status_filter: str = CRM_STATUS_FILTER_PENDING,
) -> tuple[str, dict[str, int]]:
    if isinstance(sheet_urls, str):
        urls = [sheet_urls]
    else:
        urls = list(sheet_urls)
    active_urls = [str(url or "").strip() for url in urls if str(url or "").strip()]
    full, sheet_counts = load_all_sheets_from_urls(active_urls)
    loaded_sheets = sum(1 for n in sheet_counts.values() if n > 0)
    if is_ppl_leads_template(full):
        table_note = (
            f"Google PPL for CRM ({loaded_sheets} листів)"
            if len(active_urls) == 1
            else f"Google PPL for CRM ({len(active_urls)} таблиць, {loaded_sheets} листів)"
        )
        _store_research_data(
            full,
            source_label=table_note,
            crm_status_filter=CRM_STATUS_FILTER_ALL,
        )
        return "ppl", sheet_counts
    if not has_crm_status_column(full):
        raise ValueError("У таблиці немає колонки «CRM status».")
    source_label = (
        f"Google Таблиця ({loaded_sheets} листів)"
        if len(active_urls) == 1
        else f"Google ({len(active_urls)} таблиць, {loaded_sheets} листів)"
    )
    _store_research_data(
        full,
        source_label=source_label,
        crm_status_filter=crm_status_filter,
    )
    return "research", sheet_counts


def _try_load_research_csv(
    file_bytes: bytes,
    file_name: str,
    *,
    crm_status_filter: str,
) -> None:
    full = load_source_csv(file_bytes)
    if crm_status_filter != CRM_STATUS_FILTER_ALL and not has_crm_status_column(full):
        raise ValueError(
            f"У CSV немає колонки «CRM status», а обрано фільтр "
            f"«{crm_status_filter_label(crm_status_filter)}»."
        )
    _store_research_data(
        full,
        source_label=f"Дослідження CSV: {file_name}",
        crm_status_filter=crm_status_filter,
    )


def _load_crm_leads_file(file_bytes: bytes, file_name: str) -> None:
    comp = prepare_ppl_leads_export(
        load_source_csv(file_bytes),
        person_lookup=_get_person_lookup(),
        org_lookup=_get_org_lookup(),
        **_leads_export_kwargs(),
    )
    if LINKEDIN_PERSON_COLUMN in comp.columns:
        linkedin_df = build_linkedin_export(comp)
    else:
        linkedin_df = pd.DataFrame()
    st.session_state.comp_df = comp
    st.session_state.linkedin_df = linkedin_df
    st.session_state.crm_leads_name = file_name
    st.session_state.loaded_at = datetime.now(timezone.utc)
    _persist_workspace()


def _handle_sidebar_leads_upload(file_bytes: bytes, file_name: str) -> str:
    """
    Завантаження у «2 — Ліди CRM».
    people-….csv → довідник Person; comp-шаблон → таблиця лідів.
    Повертає тип: 'persons' | 'leads'.
    """
    preview = load_source_csv(file_bytes)
    if is_persons_crm_export(preview):
        _load_persons_file(file_bytes, file_name)
        return "persons"
    _load_crm_leads_file(file_bytes, file_name)
    return "leads"


def _load_crm_companies_file(file_bytes: bytes, file_name: str) -> None:
    companies = load_source_csv(file_bytes)
    if DUPLICATE_NAME_COLUMN not in companies.columns and "Company" in companies.columns:
        companies = annotate_duplicate_company_names(companies)
    companies = apply_organization_lookup(companies, _get_org_lookup())
    st.session_state.companies_df = companies
    _clear_companies_merge_preview()
    st.session_state.crm_companies_name = file_name
    st.session_state.loaded_at = datetime.now(timezone.utc)
    _persist_workspace()


def _persist_workspace() -> None:
    save_workspace(ROOT, st.session_state)


def _persist_workspace_meta() -> None:
    save_workspace_meta(ROOT, st.session_state)


def _persist_workspace_frames(*keys: str) -> None:
    save_workspace_frames(ROOT, st.session_state, keys)


def _sync_tab_query_params() -> None:
    tab = st.session_state.get("active_tab", DEFAULT_ACTIVE_TAB)
    if tab not in VALID_TABS:
        tab = DEFAULT_ACTIVE_TAB
        st.session_state.active_tab = tab
    if st.query_params.get("tab") != tab:
        st.query_params["tab"] = tab


def _restore_workspace_once() -> None:
    if st.session_state.get("_workspace_restored"):
        return
    saved = load_workspace(ROOT)
    if saved:
        apply_workspace(saved, st.session_state)
        if not st.session_state.get("crm_status_filter"):
            st.session_state.crm_status_filter = (
                CRM_STATUS_FILTER_PENDING
                if st.session_state.get("pending_only", True)
                else CRM_STATUS_FILTER_ALL
            )
        preview = normalize_merge_changes_df(
            st.session_state.get("companies_merge_preview")
        )
        st.session_state.companies_merge_preview = preview
    st.session_state._workspace_restored = True
    qp_tab = st.query_params.get("tab")
    if qp_tab in VALID_TABS:
        st.session_state.active_tab = qp_tab
    elif st.session_state.get("active_tab") not in VALID_TABS:
        st.session_state.active_tab = DEFAULT_ACTIVE_TAB
    _sync_tab_query_params()


def _reset_all_data() -> None:
    """Повне очищення — дослідження, CRM, збереження, превʼю, файли на диску."""
    clear_workspace(ROOT)
    clear_export_artifacts(ROOT)
    _cached_org_lookup.clear()
    _cached_person_lookup.clear()
    st.session_state.full_df = None
    st.session_state.pending_df = None
    st.session_state.comp_df = None
    st.session_state.linkedin_df = None
    st.session_state.companies_df = None
    st.session_state.companies_merge_preview = None
    st.session_state.loaded_at = None
    st.session_state.load_error = None
    st.session_state.data_source_label = None
    st.session_state.pending_only = True
    st.session_state.crm_status_filter = CRM_STATUS_FILTER_PENDING
    st.session_state.crm_leads_name = None
    st.session_state.crm_companies_name = None
    st.session_state.crm_leads_key = None
    st.session_state.crm_companies_key = None
    st.session_state.research_csv_key = None
    st.session_state.crm_organizations_name = None
    st.session_state.crm_organizations_key = None
    st.session_state.organizations_row_count = None
    st.session_state.organizations_load_error = None
    st.session_state.crm_persons_name = None
    st.session_state.crm_persons_key = None
    st.session_state.persons_row_count = None
    st.session_state.persons_load_error = None
    st.session_state.file_stem = DEFAULT_FILE_STEM
    st.session_state.active_tab = DEFAULT_ACTIVE_TAB
    st.query_params["tab"] = DEFAULT_ACTIVE_TAB
    st.session_state._workspace_restored = True
    st.session_state.uploader_generation = (
        int(st.session_state.get("uploader_generation", 0)) + 1
    )
    _clear_activities_cache()
    _invalidate_leads_review_cache()
    _invalidate_companies_review_cache()


def _clear_activities_cache() -> None:
    st.session_state.activities_cache_key = None
    st.session_state.activities_file_name = None
    st.session_state.activities_linkedin_activity_key = None
    st.session_state.merged_activities_df = None
    st.session_state.activities_exclude_key = None
    st.session_state.activities_final_df = None
    st.session_state.activities_excluded_count = None
    st.session_state.activities_excluded_df = None
    st.session_state.activities_delete_source_df = None


def _dataframe_cache_token(df: pd.DataFrame | None) -> str:
    if df is None or df.empty:
        return "empty"
    digest = pd.util.hash_pandas_object(df, index=True).sum()
    return f"{len(df)}:{int(digest)}"


def _invalidate_leads_review_cache() -> None:
    st.session_state.leads_review_cache_token = None
    st.session_state.person_review_annotated = None
    st.session_state.person_review_df = None
    st.session_state.person_review_xlsx_bytes = None


def _invalidate_companies_review_cache() -> None:
    st.session_state.companies_review_cache_token = None
    st.session_state.companies_review_annotated = None
    st.session_state.companies_review_xlsx_bytes = None


def _lookup_cache_suffix() -> str:
    parts: list[str] = []
    for path_fn in (persons_crm_path, organizations_crm_path):
        path = path_fn(ROOT)
        if path.is_file():
            try:
                parts.append(str(path.stat().st_mtime_ns))
            except OSError:
                parts.append("err")
        else:
            parts.append("none")
    return "|".join(parts)


def _get_cached_person_leads_review(
    comp_df: pd.DataFrame,
    org_lookup,
) -> tuple[pd.DataFrame, pd.DataFrame, bytes]:
    token = f"{_dataframe_cache_token(comp_df)}|{_lookup_cache_suffix()}"
    cached_review = st.session_state.get("person_review_df")
    cached_annotated = st.session_state.get("person_review_annotated")
    cache_ok = (
        st.session_state.get("leads_review_cache_token") == token
        and cached_annotated is not None
        and cached_review is not None
        and st.session_state.get("person_review_xlsx_bytes") is not None
        and len(cached_review) == len(cached_annotated)
        and cached_review.index.equals(cached_annotated.index)
    )
    if cache_ok:
        return (
            cached_review,
            cached_annotated,
            st.session_state.person_review_xlsx_bytes,
        )

    annotated = prepare_person_leads_review(comp_df, org_lookup=org_lookup).reset_index(
        drop=True
    )
    review = ppl_leads_for_crm_file(annotated, include_notion_campaign_id=False)
    issues = annotated[PERSON_LOOKUP_ISSUE_COLUMN]
    xlsx_bytes = person_leads_review_to_xlsx_bytes(review, issues=issues)
    st.session_state.leads_review_cache_token = token
    st.session_state.person_review_annotated = annotated
    st.session_state.person_review_df = review
    st.session_state.person_review_xlsx_bytes = xlsx_bytes
    return review, annotated, xlsx_bytes


def _get_cached_companies_review(companies_df: pd.DataFrame) -> tuple[pd.DataFrame, bytes]:
    token = f"{_dataframe_cache_token(companies_df)}|{_lookup_cache_suffix()}"
    if (
        st.session_state.get("companies_review_cache_token") == token
        and st.session_state.get("companies_review_annotated") is not None
        and st.session_state.get("companies_review_xlsx_bytes") is not None
    ):
        return (
            st.session_state.companies_review_annotated,
            st.session_state.companies_review_xlsx_bytes,
        )

    annotated = prepare_companies_review(companies_df)
    xlsx_bytes = companies_review_to_xlsx_bytes(annotated)
    st.session_state.companies_review_cache_token = token
    st.session_state.companies_review_annotated = annotated
    st.session_state.companies_review_xlsx_bytes = xlsx_bytes
    return annotated, xlsx_bytes


def _activities_research_df() -> pd.DataFrame | None:
    """Дослідження для activities: pending/full (Researcher Name, Research ID, ICP)."""
    for key in ("pending_df", "full_df", "comp_df"):
        df = st.session_state.get(key)
        if df is not None and not (hasattr(df, "empty") and df.empty):
            return df
    return None


def _activities_research_cache_key() -> str:
    parts: list[str] = []
    for key in ("pending_df", "full_df", "comp_df"):
        df = st.session_state.get(key)
        parts.append(str(len(df)) if df is not None else "0")
    loaded = st.session_state.get("loaded_at")
    parts.append(str(loaded) if loaded else "")
    return "|".join(parts)


def _activities_leads_cache_key() -> str:
    """Зміни Person - Linkedin Active у comp_df → перезлиття activities."""
    comp = st.session_state.get("comp_df")
    if comp is None or getattr(comp, "empty", True):
        return "0"
    col = PERSON_LINKEDIN_ACTIVE_COLUMN
    if col not in comp.columns:
        return "0"
    work = comp[[col] + ([c for c in ("Email", "Person - Apollo Contact id") if c in comp.columns])].fillna("")
    digest = int(pd.util.hash_pandas_object(work, index=True).sum())
    filled = int(work[col].astype(str).str.strip().ne("").sum())
    linkedin_key = st.session_state.get("activities_linkedin_activity_key") or ""
    return f"{len(comp)}:{filled}:{digest}|{linkedin_key}"


def _sync_activities_linkedin_from_leads() -> int:
    """Оновити Person - Linkedin Active у злитому activities з comp_df."""
    merged = st.session_state.get("merged_activities_df")
    leads = st.session_state.get("comp_df")
    if merged is None or getattr(merged, "empty", True):
        return 0
    updated, count = apply_leads_linkedin_active(merged, leads)
    st.session_state.merged_activities_df = updated
    return count


def _apply_linkedin_activity_file(file_bytes: bytes, file_name: str) -> dict[str, object]:
    """LinkedIn activity: comp_df + поточний merged activities."""
    result: dict[str, object] = {}
    comp = st.session_state.get("comp_df")
    if comp is not None and not comp.empty:
        merged_comp, comp_stats = merge_linkedin_active_into_comp(comp, file_bytes)
        st.session_state.comp_df = align_comp_to_crm_schema(merged_comp)
        result["comp"] = comp_stats
        _invalidate_leads_review_cache()
    else:
        result["comp"] = None

    merged = st.session_state.get("merged_activities_df")
    if merged is not None and not merged.empty:
        if result.get("comp") is not None:
            merged, leads_n = apply_leads_linkedin_active(merged, st.session_state.comp_df)
            result["activities_from_leads"] = leads_n
        else:
            merged, act_stats = merge_linkedin_active_into_df(
                merged,
                file_bytes,
                linkedin_columns=("Person - LinkedIn", "Linkedin Person"),
            )
            result["activities_direct"] = act_stats
        st.session_state.merged_activities_df = merged
    else:
        result["activities"] = None

    st.session_state.activities_linkedin_activity_key = f"{file_name}:{len(file_bytes)}"
    st.session_state.activities_cache_key = None
    st.session_state.activities_exclude_key = None
    st.session_state.activities_final_df = None
    st.session_state.activities_excluded_count = None
    st.session_state.activities_excluded_df = None
    st.session_state.activities_delete_source_df = None
    _persist_workspace()
    return result


def _invalidate_activities_merge_cache() -> None:
    st.session_state.activities_cache_key = None
    st.session_state.merged_activities_df = None
    st.session_state.activities_exclude_key = None
    st.session_state.activities_final_df = None
    st.session_state.activities_excluded_count = None
    st.session_state.activities_excluded_df = None
    st.session_state.activities_delete_source_df = None


def _sync_activities_merge(activities_file) -> pd.DataFrame | None:
    """Злиття activities + research; повторно лише при новому файлі або зміні дослідження."""
    if activities_file is None:
        return st.session_state.get("merged_activities_df")

    act_key = f"{activities_file.name}:{activities_file.size}"
    cache_key = f"{act_key}|{_activities_research_cache_key()}|{_activities_leads_cache_key()}"
    cached = st.session_state.get("merged_activities_df")
    if st.session_state.get("activities_cache_key") == cache_key and cached is not None:
        return cached

    research_df = _activities_research_df()
    leads_df = st.session_state.get("comp_df")
    activities_df = load_source_csv(activities_file.getvalue())
    merged_df = merge_activities_with_research(
        activities_df, research_df, leads_df=leads_df
    )
    st.session_state.activities_cache_key = cache_key
    st.session_state.activities_file_name = activities_file.name
    st.session_state.merged_activities_df = merged_df
    st.session_state.activities_exclude_key = None
    st.session_state.activities_final_df = None
    st.session_state.activities_excluded_count = None
    st.session_state.activities_excluded_df = None
    st.session_state.activities_delete_source_df = None
    return merged_df


def _sync_activities_exclude(
    merged_df: pd.DataFrame, exclude_file
) -> tuple[pd.DataFrame, pd.DataFrame, int, pd.DataFrame | None] | None:
    """Виключення; робочий del-файл або список Person - ID."""
    if exclude_file is None:
        final_df = st.session_state.get("activities_final_df")
        excluded_df = st.session_state.get("activities_excluded_df")
        delete_source = st.session_state.get("activities_delete_source_df")
        count = st.session_state.get("activities_excluded_count")
        if final_df is not None and count is not None:
            if excluded_df is None:
                excluded_df = merged_df.iloc[0:0].copy()
            return final_df, excluded_df, int(count), delete_source
        return None

    ex_key = f"{exclude_file.name}:{exclude_file.size}"
    if (
        st.session_state.get("activities_exclude_key") == ex_key
        and st.session_state.get("activities_final_df") is not None
        and st.session_state.get("activities_excluded_count") is not None
    ):
        excluded_df = st.session_state.get("activities_excluded_df")
        if excluded_df is None:
            excluded_df = merged_df.iloc[0:0].copy()
        return (
            st.session_state.activities_final_df,
            excluded_df,
            int(st.session_state.activities_excluded_count),
            st.session_state.get("activities_delete_source_df"),
        )

    if exclude_file.name.endswith(".xlsx"):
        exclude_df = pd.read_excel(exclude_file, dtype=str)
    else:
        exclude_df = pd.read_csv(exclude_file, dtype=str)
    final_df, excluded_df, excluded_count, delete_source = exclude_activities(
        merged_df, exclude_df
    )
    st.session_state.activities_exclude_key = ex_key
    st.session_state.activities_final_df = final_df
    st.session_state.activities_excluded_df = excluded_df
    st.session_state.activities_delete_source_df = delete_source
    st.session_state.activities_excluded_count = excluded_count
    return final_df, excluded_df, excluded_count, delete_source


def _clear_crm_leads() -> None:
    st.session_state.crm_leads_name = None
    st.session_state.crm_leads_key = None
    pending = st.session_state.pending_df
    if st.session_state.full_df is not None and pending is not None and not pending.empty:
        _rebuild_leads_export()
    else:
        st.session_state.comp_df = None
        st.session_state.linkedin_df = None
    _persist_workspace()


def _clear_crm_companies() -> None:
    st.session_state.crm_companies_name = None
    st.session_state.crm_companies_key = None
    pending = st.session_state.pending_df
    if st.session_state.full_df is not None and pending is not None and not pending.empty:
        st.session_state.companies_df = build_company_export(
            pending, org_lookup=_get_org_lookup()
        )
    else:
        st.session_state.companies_df = None
    st.session_state.companies_merge_preview = None
    _persist_workspace()


def _has_workspace_data() -> bool:
    return (
        st.session_state.comp_df is not None
        or st.session_state.companies_df is not None
        or st.session_state.full_df is not None
    )


def _workspace_data_status() -> tuple[bool, str]:
    """Короткий підпис: чи є завантажені дані (для заголовка біля «Почати спочатку»)."""
    parts: list[str] = []
    if organizations_crm_exists(ROOT) or st.session_state.get("crm_organizations_name"):
        n = st.session_state.get("organizations_row_count")
        label = st.session_state.get("crm_organizations_name") or "база організацій"
        parts.append(
            f"організації ({label}) — {n if n else '?'} ряд."
        )
    if persons_crm_exists(ROOT) or st.session_state.get("crm_persons_name"):
        pn = st.session_state.get("persons_row_count")
        plabel = st.session_state.get("crm_persons_name") or "база persons"
        parts.append(f"persons ({plabel}) — {pn if pn else '?'} ряд.")
    if st.session_state.full_df is not None:
        parts.append(f"дослідження — {len(st.session_state.full_df)} ряд.")
    if st.session_state.comp_df is not None:
        label = "ліди"
        if st.session_state.crm_leads_name:
            label = f"ліди ({st.session_state.crm_leads_name})"
        parts.append(f"{label} — {len(st.session_state.comp_df)} ряд.")
    if st.session_state.companies_df is not None:
        label = "компанії"
        if st.session_state.crm_companies_name:
            label = f"компанії ({st.session_state.crm_companies_name})"
        parts.append(f"{label} — {len(st.session_state.companies_df)} ряд.")
    linkedin = st.session_state.linkedin_df
    if linkedin is not None and not linkedin.empty:
        parts.append(f"LinkedIn — {len(linkedin)} URL")
    preview = st.session_state.get("companies_merge_preview")
    if preview is not None and hasattr(preview, "empty") and not preview.empty:
        parts.append("превʼю злиття дублікатів")
    has_data = bool(parts)
    if has_data:
        disk = ", збережено локально" if workspace_exists(ROOT) else ""
        return True, "; ".join(parts) + disk
    if workspace_exists(ROOT):
        return False, "У сесії порожньо; на диску є збережений стан (оновіть сторінку)"
    return False, "Немає завантажених даних"


st.set_page_config(page_title="CRM Update", layout="wide", initial_sidebar_state="expanded")

if "full_df" not in st.session_state:
    st.session_state.full_df = None
    st.session_state.pending_df = None
    st.session_state.comp_df = None
    st.session_state.linkedin_df = None
    st.session_state.companies_df = None
    st.session_state.loaded_at = None
    st.session_state.load_error = None
    st.session_state.data_source_label = None
    st.session_state.pending_only = True
    st.session_state.crm_status_filter = CRM_STATUS_FILTER_PENDING
    st.session_state.research_csv_key = None
    st.session_state.crm_leads_key = None
    st.session_state.crm_companies_key = None
    st.session_state.crm_leads_name = None
    st.session_state.crm_companies_name = None
    st.session_state.companies_merge_preview = None
    st.session_state.file_stem = DEFAULT_FILE_STEM
    st.session_state.fill_notion_campaign_id = False
    st.session_state.leads_default_activity_type = DEFAULT_ACTIVITY_TYPE
    st.session_state.active_tab = DEFAULT_ACTIVE_TAB
    st.session_state._workspace_restored = False
    st.session_state.uploader_generation = 0
    st.session_state.crm_organizations_name = None
    st.session_state.crm_organizations_key = None
    st.session_state.organizations_row_count = None
    st.session_state.organizations_load_error = None
    st.session_state.crm_persons_name = None
    st.session_state.crm_persons_key = None
    st.session_state.persons_row_count = None
    st.session_state.persons_load_error = None
    st.session_state.activities_cache_key = None
    st.session_state.activities_file_name = None
    st.session_state.activities_linkedin_activity_key = None
    st.session_state.merged_activities_df = None
    st.session_state.activities_exclude_key = None
    st.session_state.activities_final_df = None
    st.session_state.activities_excluded_count = None
    st.session_state.activities_excluded_df = None
    st.session_state.activities_delete_source_df = None
    st.session_state.leads_review_cache_token = None
    st.session_state.person_review_annotated = None
    st.session_state.person_review_df = None
    st.session_state.person_review_xlsx_bytes = None
    st.session_state.companies_review_cache_token = None
    st.session_state.companies_review_annotated = None
    st.session_state.companies_review_xlsx_bytes = None

_restore_workspace_once()
if organizations_crm_exists(ROOT) and st.session_state.get("organizations_row_count") is None:
    lookup = _get_org_lookup()
    if lookup:
        st.session_state.organizations_row_count = lookup.row_count
if persons_crm_exists(ROOT) and st.session_state.get("persons_row_count") is None:
    plookup = _get_person_lookup()
    if plookup:
        st.session_state.persons_row_count = plookup.row_count

_header_left, _header_right = st.columns([4, 2])
with _header_left:
    st.title("CRM Update")
    st.caption("Дослідження → файли CRM. У бічній панелі — готові файли лідів і компаній.")
    if workspace_exists(ROOT):
        st.caption("Робочий стан збережено локально — після оновлення сторінки дані відновляться.")
with _header_right:
    _has_data, _data_label = _workspace_data_status()
    if _has_data:
        st.success("Є дані")
        st.caption(_data_label)
    else:
        st.caption(f"**Немає даних** — {_data_label}")
    if st.button(
        "Почати спочатку",
        type="primary",
        help="Очистити всі дані: дослідження, CRM, превʼю, локальне збереження та CSV/XLSX у папці проєкту",
        key="reset_all_data",
    ):
        _reset_all_data()
        st.rerun()


def _clear_companies_merge_preview() -> None:
    st.session_state.companies_merge_preview = None
    _persist_workspace()


def _render_companies_merge_preview() -> None:
    raw = st.session_state.get("companies_merge_preview")
    if raw is None:
        return
    changes = normalize_merge_changes_df(raw)
    if changes is None:
        if raw is not None and hasattr(raw, "empty") and not raw.empty:
            st.warning(
                "Превʼю змін пошкоджене або застаріле — застосуйте CSV дублікатів ще раз."
            )
            _clear_companies_merge_preview()
        return
    st.session_state.companies_merge_preview = changes

    st.divider()
    st.subheader("Превʼю застосованих змін")
    if changes.empty:
        st.info(
            "Файл оброблено, але значення в колонках не змінились "
            "(ті самі дані, що були в таблиці)."
        )
    else:
        org_count = changes[APOLLO_ACCOUNT_ID_COLUMN].nunique()
        st.caption(
            f"**{len(changes)}** змін у полях · "
            f"**{org_count}** організацій ({APOLLO_ACCOUNT_ID_COLUMN})."
        )
        st.dataframe(
            format_merge_changes_summary(changes),
            use_container_width=True,
            hide_index=True,
        )
        with st.expander("Детально: було → стало", expanded=False):
            st.dataframe(changes, use_container_width=True, hide_index=True)
    if st.button("Приховати превʼю", key="clear_companies_merge_preview"):
        _clear_companies_merge_preview()
        st.rerun()


def _on_active_tab_change() -> None:
    _sync_tab_query_params()
    _persist_workspace_meta()


_uploader_gen = int(st.session_state.get("uploader_generation", 0))

with st.sidebar:
    st.header("2 файли з CRM")
    st.caption(
        "**1** — організації CRM: **до** імпорту companies (перша вигрузка), "
        "**після** імпорту — **нова** вигрузка з CRM (усі ID).  \n"
        "**2** — people / готові ліди. Компанії — з **дослідження**."
    )

    org_upload = st.file_uploader(
        "1 — Організації CRM",
        type=["csv"],
        help="Експорт organizations з CRM. Після імпорту companies — завантажте оновлений файл знову.",
        key=f"sidebar_crm_organizations_{_uploader_gen}",
    )
    if org_upload is not None:
        org_key = f"{org_upload.name}:{org_upload.size}"
        if st.session_state.crm_organizations_key != org_key:
            try:
                with st.spinner(
                    "База організацій + перерахунок Organization ID "
                    "у компаніях і лідах…"
                ):
                    _load_organizations_file(org_upload.getvalue(), org_upload.name)
                st.session_state.crm_organizations_key = org_key
                n = st.session_state.organizations_row_count
                matched_co = matched_leads = 0
                cdf = st.session_state.companies_df
                if cdf is not None and "Organization - ID" in cdf.columns:
                    matched_co = int(
                        (cdf["Organization - ID"].astype(str).str.strip() != "").sum()
                    )
                ldf = st.session_state.comp_df
                if ldf is not None and "Organization - ID" in ldf.columns:
                    matched_leads = int(
                        (ldf["Organization - ID"].astype(str).str.strip() != "").sum()
                    )
                st.success(
                    f"Організації: **{n:,}** · ID у компаніях: **{matched_co}** · "
                    f"у лідах PPL: **{matched_leads}**"
                )
                if st.session_state.companies_df is not None:
                    st.caption(
                        "У **головній області** оберіть розділ **CRM companies** — "
                        "там колонки Organization / Website / Linkedin ID."
                    )
            except Exception as exc:
                st.error(str(exc))

    if st.session_state.crm_organizations_name:
        st.caption(f"Організації: `{st.session_state.crm_organizations_name}`")
        if st.button("Видалити базу організацій", key="delete_crm_organizations", type="secondary"):
            _clear_organizations()
            st.session_state.uploader_generation = _uploader_gen + 1
            st.rerun()
    elif not organizations_crm_exists(ROOT):
        st.info(
            "Завантажте **перший** експорт organizations з CRM. "
            "Після імпорту companies — **другий** (оновлений) файл."
        )
    else:
        st.caption("База організацій збережена. Після імпорту в CRM — замініть файл новою вигрузкою.")

    with st.expander("Цикл: organizations → companies → CRM → organizations", expanded=False):
        st.markdown(format_organizations_crm_workflow_help())

    with st.expander("Які колонки для метчу (організації + Comp)", expanded=False):
        st.markdown(format_organization_csv_help())
        st.markdown(format_company_export_help())

    if st.session_state.get("organizations_load_error"):
        st.warning(st.session_state.organizations_load_error)

    with st.expander("Довідник Person CRM (для лідів)", expanded=False):
        st.markdown(format_person_csv_help())
        st.markdown(format_leads_person_columns_help())
        persons_upload = st.file_uploader(
            "CSV контактів (persons)",
            type=["csv"],
            key=f"sidebar_crm_persons_{_uploader_gen}",
        )
        if persons_upload is not None:
            persons_key = f"{persons_upload.name}:{persons_upload.size}"
            if st.session_state.crm_persons_key != persons_key:
                try:
                    with st.spinner("Завантаження бази контактів…"):
                        _load_persons_file(persons_upload.getvalue(), persons_upload.name)
                    st.session_state.crm_persons_key = persons_key
                    n = st.session_state.persons_row_count
                    matched = 0
                    cdf = st.session_state.comp_df
                    if cdf is not None and PERSON_ID_COLUMN in cdf.columns:
                        matched = int(
                            (cdf[PERSON_ID_COLUMN].astype(str).str.strip() != "").sum()
                        )
                    st.success(f"Persons: **{n:,}** · з Person - ID у лідах: **{matched}**")
                    if st.session_state.comp_df is not None:
                        st.caption("Вкладка **CRM Leads** → крок **Person ID**.")
                except Exception as exc:
                    st.error(str(exc))
        if st.session_state.crm_persons_name:
            st.caption(f"Persons: `{st.session_state.crm_persons_name}`")
            if st.button(
                "Видалити базу persons",
                key="delete_crm_persons",
                type="secondary",
            ):
                _clear_persons()
                st.session_state.uploader_generation = _uploader_gen + 1
                st.rerun()
        elif not persons_crm_exists(ROOT):
            st.caption("Завантажте експорт **усіх контактів** з CRM для Person - ID у лідах.")
        else:
            st.caption("База persons збережена локально.")
        if st.session_state.get("persons_load_error"):
            st.warning(st.session_state.persons_load_error)

    st.divider()

    leads_upload = st.file_uploader(
        "2 — Ліди CRM",
        type=["csv"],
        help="Comp-шаблон для імпорту або people-….csv (автоматично піде в довідник Person).",
        key=f"sidebar_crm_leads_{_uploader_gen}",
    )
    if leads_upload is not None:
        upload_key = f"{leads_upload.name}:{leads_upload.size}"
        if (
            st.session_state.crm_leads_key != upload_key
            and st.session_state.crm_persons_key != upload_key
        ):
            try:
                with st.spinner(f"Завантаження {leads_upload.name}…"):
                    kind = _handle_sidebar_leads_upload(
                        leads_upload.getvalue(), leads_upload.name
                    )
                if kind == "persons":
                    st.session_state.crm_persons_key = upload_key
                    n = st.session_state.persons_row_count
                    matched = 0
                    cdf = st.session_state.comp_df
                    if cdf is not None and PERSON_ID_COLUMN in cdf.columns:
                        matched = int(
                            (cdf[PERSON_ID_COLUMN].astype(str).str.strip() != "").sum()
                        )
                    st.success(
                        f"**people** збережено як довідник Person: **{n:,}** контактів · "
                        f"з Person - ID у лідах: **{matched}**"
                    )
                    st.info(
                        "Це **база контактів** CRM (Person - ID, LinkedIn), не таблиця лідів для імпорту. "
                        "Рядки лідів — з **дослідження зверху**; Person ID — вкладка **CRM Leads**, крок 3."
                    )
                else:
                    st.session_state.crm_leads_key = upload_key
                    n_li = len(st.session_state.linkedin_df or [])
                    st.success(
                        f"Ліди (comp): **{len(st.session_state.comp_df)}** · "
                        f"LinkedIn URL: **{n_li}**"
                    )
            except Exception as exc:
                st.error(str(exc))

    if st.session_state.crm_leads_name:
        st.caption(f"Ліди: `{st.session_state.crm_leads_name}`")
        if st.button("Видалити файл лідів", key="delete_crm_leads", type="secondary"):
            _clear_crm_leads()
            st.session_state.uploader_generation = _uploader_gen + 1
            st.rerun()

    with st.expander("Додатково: замінити компанії готовим CSV", expanded=False):
        st.caption(
            "Зазвичай не потрібно: компанії для CRM будуються з дослідження. "
            "Лише якщо хочете підставити свій готовий файл замість згенерованого."
        )
        companies_upload = st.file_uploader(
            "CSV компаній (опційно)",
            type=["csv"],
            key=f"sidebar_crm_companies_{_uploader_gen}",
        )
        if companies_upload is not None:
            companies_key = f"{companies_upload.name}:{companies_upload.size}"
            if st.session_state.crm_companies_key != companies_key:
                try:
                    _load_crm_companies_file(companies_upload.getvalue(), companies_upload.name)
                    st.session_state.crm_companies_key = companies_key
                    st.success(f"Компанії: **{len(st.session_state.companies_df)}**")
                except Exception as exc:
                    st.error(str(exc))

        if st.session_state.crm_companies_name:
            st.caption(f"Файл: `{st.session_state.crm_companies_name}`")
            if st.button(
                "Видалити завантажені компанії",
                key="delete_crm_companies",
                type="secondary",
            ):
                _clear_crm_companies()
                st.session_state.uploader_generation = _uploader_gen + 1
                st.rerun()

    st.divider()
    st.text_input(
        "Базове імʼя збережених файлів",
        help="06_03 → «06_03 - comp.csv» тощо",
        key="file_stem",
        on_change=_persist_workspace_meta,
    )

    if workspace_exists(ROOT) or _has_workspace_data():
        st.caption("Дані збережені між оновленнями сторінки.")

(
    comp_path,
    linkedin_path,
    companies_path,
    companies_dupes_path,
    companies_org_path,
    companies_review_xlsx_path,
    leads_person_path,
    leads_person_xlsx_path,
    research_duplicates_path,
) = _export_paths(st.session_state.file_stem)

st.header("Дослідження — зібрати файли з Google / CSV")

source_mode = st.radio(
    "Джерело дослідження",
    options=["google", "csv"],
    format_func=lambda x: "Google Таблиця" if x == "google" else "CSV дослідження",
    horizontal=True,
)

_default_crm_filter = st.session_state.get(
    "crm_status_filter", CRM_STATUS_FILTER_PENDING
)
crm_status_filter = _default_crm_filter
sheet_urls_text = DEFAULT_SPREADSHEET_URL
research_csv = None

if source_mode == "google":
    sheet_urls_text = st.text_area(
        "Посилання на Google Таблицю",
        value=DEFAULT_SPREADSHEET_URL,
        height=88,
        key="google_sheet_urls",
        placeholder="Одне або кілька посилань — кожне з нового рядка",
    )
    _crm_filter_options = (
        CRM_STATUS_FILTER_PENDING,
        CRM_STATUS_FILTER_VERIFICATION,
        CRM_STATUS_FILTER_PENDING_AND_VERIFICATION,
    )

    def _crm_filter_label(value: str) -> str:
        if value == CRM_STATUS_FILTER_PENDING:
            return "Порожній (pending)"
        if value == CRM_STATUS_FILTER_VERIFICATION:
            return "verification"
        return "Порожній + verification"

    _crm_filter_index = (
        _crm_filter_options.index(_default_crm_filter)
        if _default_crm_filter in _crm_filter_options
        else 0
    )
    crm_status_filter = st.radio(
        "CRM status для лідів",
        options=_crm_filter_options,
        format_func=_crm_filter_label,
        index=_crm_filter_index,
        horizontal=True,
        key="google_crm_status_filter",
    )
    st.caption(
        "Одне або кілька посилань (кожне з нового рядка). "
        "З кожної таблиці завантажуються **усі листи** (колонка **Source sheet**). "
        "Якщо таблиць кілька — у Source sheet буде **Таблиця 1** / **Таблиця 2** тощо. "
        "**pending** — нові ліди; **verification** — на перевірці; "
        "**Порожній + verification** — обидва разом."
    )
else:
    research_csv = st.file_uploader(
        "CSV з дослідження (готовий файл — усі рядки з файлу)",
        type=["csv"],
        key="research_csv_upload",
    )
    if research_csv is not None:
        rkey = f"{research_csv.name}:{research_csv.size}"
        if st.session_state.research_csv_key != rkey:
            with st.spinner(f"Обробка {research_csv.name}…"):
                try:
                    _try_load_research_csv(
                        research_csv.getvalue(),
                        research_csv.name,
                        crm_status_filter=CRM_STATUS_FILTER_ALL,
                    )
                    st.session_state.research_csv_key = rkey
                    st.success(
                        f"Згенеровано з **{research_csv.name}**: ліди **{len(st.session_state.comp_df)}**, "
                        f"компанії **{len(st.session_state.companies_df)}**."
                    )
                except Exception as exc:
                    st.error(str(exc))

col_gen, col_save = st.columns([1, 1])

with col_gen:
    if source_mode == "google":
        gen_clicked = st.button("Завантажити з Google і зібрати файли", type="primary")
    else:
        gen_clicked = False

with col_save:
    save_clicked = st.button(
        "Зберегти файли на диск",
        help="Потрібно спочатку завантажити дослідження з Google або CSV.",
    )

if st.session_state.pending_df is not None and not st.session_state.pending_df.empty:
    _dl_stem = st.session_state.file_stem or DEFAULT_FILE_STEM
    _dl_filter = st.session_state.get("crm_status_filter", CRM_STATUS_FILTER_PENDING)
    _dl_label = crm_status_filter_label(_dl_filter).replace(" ", "-").replace("(", "").replace(")", "")
    _dl_name = f"{_dl_stem} - leads-{_dl_label}.csv"
    st.download_button(
        "⬇ Завантажити CSV лідів (для редагування)",
        data=dataframe_to_csv_bytes(st.session_state.pending_df),
        file_name=_dl_name,
        mime="text/csv",
        help=(
            "Тільки ліди за обраним фільтром CRM status. "
            "Відредагуйте в Excel / Google Sheets, "
            "потім оберіть «CSV дослідження» і завантажте змінений файл — "
            "всі файли перезберуться автоматично."
        ),
    )

if gen_clicked and source_mode == "google":
    google_urls = parse_spreadsheet_urls(sheet_urls_text)
    if not google_urls:
        st.error("Вставте хоча б одне посилання на Google Таблицю.")
    else:
        with st.spinner(
            "Завантаження всіх листів…"
            if len(google_urls) == 1
            else f"Завантаження {len(google_urls)} таблиць…"
        ):
            try:
                _kind, sheet_counts = _try_load_google(
                    google_urls, crm_status_filter=crm_status_filter
                )
                loaded = [f"**{n}** ({c})" for n, c in sheet_counts.items() if c > 0]
                sheets_note = f" · листи: {', '.join(loaded[:8])}"
                if len(loaded) > 8:
                    sheets_note += f" … (+{len(loaded) - 8})"
                st.success(
                    f"Готово ({crm_status_filter_label(crm_status_filter)}): "
                    f"ліди **{len(st.session_state.comp_df)}**, "
                    f"компанії **{len(st.session_state.companies_df)}**, "
                    f"рядків дослідження **{len(st.session_state.full_df)}**"
                    f"{sheets_note}."
                )
            except Exception as exc:
                st.error(str(exc))

if save_clicked:
    if not _has_workspace_data():
        st.warning("Спочатку завантажте дослідження з Google або CSV.")
    else:
        saved: list[str] = []
        if st.session_state.comp_df is not None:
            p = save_comp_dataframe(st.session_state.comp_df, comp_path)
            saved.append(f"`{p}` ({len(st.session_state.comp_df)} лідів)")
        if st.session_state.companies_df is not None:
            p = save_comp_dataframe(st.session_state.companies_df, companies_path)
            saved.append(f"`{p}` ({len(st.session_state.companies_df)} компаній)")
        if st.session_state.linkedin_df is not None and len(st.session_state.linkedin_df):
            if st.session_state.pending_df is not None:
                p = save_linkedin_export(st.session_state.pending_df, linkedin_path)
            else:
                tmp = Path(linkedin_path)
                tmp.parent.mkdir(parents=True, exist_ok=True)
                st.session_state.linkedin_df.to_csv(tmp, index=False, encoding="utf-8-sig")
                p = tmp
            saved.append(f"`{p}` ({len(st.session_state.linkedin_df)} URL)")
        st.success("Збережено:\n" + "\n".join(f"- {s}" for s in saved))

st.divider()

if not _has_workspace_data():
    st.info(
        "**Бічна панель:** 1) організації CRM, 2) ліди CRM.  \n"
        "**Зверху:** дослідження → з нього збираються **ліди + компанії** для імпорту."
    )
else:
    st.header("Робота з файлами CRM")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ліди", len(st.session_state.comp_df) if st.session_state.comp_df is not None else 0)
    c2.metric(
        "Компанії",
        len(st.session_state.companies_df) if st.session_state.companies_df is not None else 0,
    )
    c3.metric(
        "LinkedIn URL",
        len(st.session_state.linkedin_df) if st.session_state.linkedin_df is not None else 0,
    )
    if st.session_state.full_df is not None:
        c4.metric("Рядків дослідження", len(st.session_state.full_df))
    elif st.session_state.loaded_at:
        c4.metric("Оновлено", "✓")

    labels: list[str] = []
    if st.session_state.crm_leads_name:
        labels.append(f"ліди: {st.session_state.crm_leads_name}")
    if st.session_state.crm_companies_name:
        labels.append(f"компанії: {st.session_state.crm_companies_name}")
    if st.session_state.data_source_label:
        labels.append(st.session_state.data_source_label)
    if labels and st.session_state.loaded_at:
        st.caption(
            f"{' · '.join(labels)} · "
            f"{st.session_state.loaded_at.astimezone().strftime('%Y-%m-%d %H:%M:%S')}"
        )

    if st.session_state.full_df is not None:
        full_research = st.session_state.full_df
        dupe_stats = research_duplicate_stats(full_research)
        if (
            dupe_stats["rows_with_any_duplicate"] > 0
            or dupe_stats["duplicate_email_keys"]
            or dupe_stats["duplicate_linkedin_keys"]
        ):
            st.subheader("Дублікати в дослідженні (Email / Linkedin Person)")
            d1, d2, d3, d4 = st.columns(4)
            d1.metric("Дублікатів Email", dupe_stats["duplicate_email_keys"])
            d2.metric("Рядків з дублікатом Email", dupe_stats["rows_with_duplicate_email"])
            d3.metric("Дублікатів LinkedIn", dupe_stats["duplicate_linkedin_keys"])
            d4.metric(
                "Рядків з дублікатом LinkedIn",
                dupe_stats["rows_with_duplicate_linkedin"],
            )
            st.warning(
                f"**{dupe_stats['rows_with_any_duplicate']}** з **{dupe_stats['total']}** "
                f"рядків — повторюваний **Email** і/або **Linkedin Person** "
                f"(колонки **{RESEARCH_DUPLICATE_EMAIL_COLUMN}**, "
                f"**{RESEARCH_DUPLICATE_LINKEDIN_COLUMN}** = `yes`). "
                "Перевірте перед імпортом у CRM."
            )
            report_df = build_research_duplicates_report(full_research)
            if not report_df.empty:
                st.dataframe(report_df, use_container_width=True, height=220)
            dupe_rows_df = research_duplicate_rows(full_research)
            preview_cols = [
                c
                for c in (
                    "First name",
                    "Last name",
                    EMAIL_COLUMN,
                    LINKEDIN_PERSON_COLUMN,
                    "Company",
                    RESEARCH_DUPLICATE_EMAIL_COLUMN,
                    RESEARCH_DUPLICATE_LINKEDIN_COLUMN,
                )
                if c in dupe_rows_df.columns
            ]
            if preview_cols:
                st.dataframe(
                    dupe_rows_df[preview_cols].head(40),
                    use_container_width=True,
                    height=260,
                )

            def _highlight_research_duplicates(row: pd.Series):
                styles = [""] * len(row)
                for i, col in enumerate(row.index):
                    if col == RESEARCH_DUPLICATE_EMAIL_COLUMN and str(
                        row.get(col, "")
                    ).strip() == "yes":
                        styles[i] = "background-color: #fff3cd"
                    elif col == RESEARCH_DUPLICATE_LINKEDIN_COLUMN and str(
                        row.get(col, "")
                    ).strip() == "yes":
                        styles[i] = "background-color: #fff3cd"
                    elif col == EMAIL_COLUMN and str(
                        row.get(RESEARCH_DUPLICATE_EMAIL_COLUMN, "")
                    ).strip() == "yes":
                        styles[i] = "background-color: #fff3cd"
                    elif col == LINKEDIN_PERSON_COLUMN and str(
                        row.get(RESEARCH_DUPLICATE_LINKEDIN_COLUMN, "")
                    ).strip() == "yes":
                        styles[i] = "background-color: #fff3cd"
                return styles

            if len(dupe_rows_df) > 0 and preview_cols:
                with st.expander("Превʼю з підсвіткою (до 50 рядків)", expanded=False):
                    st.dataframe(
                        dupe_rows_df[preview_cols]
                        .head(50)
                        .style.apply(_highlight_research_duplicates, axis=1),
                        use_container_width=True,
                        height=280,
                    )
            st.download_button(
                label=f"CSV дублікатів дослідження ({len(dupe_rows_df)})",
                data=dataframe_to_csv_bytes(dupe_rows_df),
                file_name=research_duplicates_path.name,
                mime="text/csv",
                key="download_research_duplicates_csv",
            )
        else:
            st.success(
                "Дублікатів **Email** та **Linkedin Person** (/in/) у дослідженні не знайдено."
            )

    def _tab_label(t: str) -> str:
        if t == "companies":
            return "CRM companies"
        if t == "leads":
            return "CRM Leads"
        return "CRM Activities"

    st.radio(
        "Розділ",
        options=list(VALID_TABS),
        format_func=_tab_label,
        horizontal=True,
        key="active_tab",
        label_visibility="collapsed",
        on_change=_on_active_tab_change,
    )
    _sync_tab_query_params()

    if st.session_state.active_tab == "leads":
        if st.session_state.comp_df is None:
            st.warning("Завантажте **файл лідів** у бічній панелі або згенеруйте з дослідження.")
        else:
            if _sync_comp_person_ids():
                st.info(
                    "Підставлено **Person - ID**, **Person - ID by Email**, "
                    "**ID by Linkedin Person**, **CRM by Email ID**, **CRM Linkedin ID**."
                )
            if _sync_comp_org_ids():
                st.info(
                    "Підставлено **Organization - ID** з бази організацій "
                    "(спочатку **Website**, інакше **Linkedin Company** — як у Google PPL)."
                )
            comp_df = _align_leads_for_display(st.session_state.comp_df)

            st.subheader("Крок 1 — LinkedIn URL для тулзи")
            st.caption(
                f"`{linkedin_path.name}` · **{ICP_NAME_EXPORT_COLUMN}** + **{LINKEDIN_PERSON_COLUMN}** · "
                f"`https://www.linkedin.com/in/.../` (слеш в кінці)"
            )
            linkedin_df = st.session_state.linkedin_df
            if linkedin_df is not None and ICP_NAME_EXPORT_COLUMN not in linkedin_df.columns:
                src = _leads_export_source_df()
                if src is not None:
                    st.session_state.linkedin_df = build_linkedin_export(src)
                    linkedin_df = st.session_state.linkedin_df
            if linkedin_df is not None and len(linkedin_df):
                st.dataframe(linkedin_df, use_container_width=True, height=240)
                st.download_button(
                    label="Скачати linkedin CSV",
                    data=dataframe_to_csv_bytes(linkedin_df),
                    file_name=linkedin_path.name,
                    mime="text/csv",
                )
            elif LINKEDIN_PERSON_COLUMN not in st.session_state.comp_df.columns:
                st.warning(f"У файлі лідів немає колонки «{LINKEDIN_PERSON_COLUMN}».")

            st.divider()
            st.subheader("Крок 2 — активність LinkedIn → leads")
            st.caption(
                f"Завантажте **частковий** CSV з тулзи (не обовʼязково всі URL з кроку 1). "
                f"Колонка **{PERSON_LINKEDIN_ACTIVE_COLUMN}** заповнюється так:\n\n"
                "- **yes** — профіль **є у файлі**, Recent Activity **< 1 місяця**\n"
                "- **no** — профіль **є у файлі**, Recent Activity **≥ 1 місяця** "
                "(або значення не розпарсилось)\n"
                "- **порожньо** — профілю **немає у файлі** (не перевіряли) "
                "або в ліді немає Linkedin Person"
            )
            activity_file = st.file_uploader(
                "Результати тулзи (CSV)",
                type=["csv"],
                key="linkedin_activity_upload",
            )
            if st.button(
                "Додати Person - Linkedin Active",
                disabled=activity_file is None,
                key="merge_linkedin_activity",
            ):
                try:
                    result = _apply_linkedin_activity_file(
                        activity_file.getvalue(), activity_file.name
                    )
                    stats = result.get("comp") or {}
                    st.success(
                        f"**yes:** {stats.get('yes', 0)} · **no** (у файлі): {stats.get('no_from_activity', 0)} · "
                        f"**порожньо** (не у файлі): {stats.get('not_in_activity_file', 0)} · "
                        f"без LinkedIn: {stats.get('empty_linkedin', 0)}"
                    )
                    if result.get("activities_from_leads") is not None:
                        st.info(
                            f"Activities оновлено: **{result['activities_from_leads']}** рядків "
                            f"({PERSON_LINKEDIN_ACTIVE_COLUMN})."
                        )
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

            st.divider()
            st.subheader("Крок 3 — Person ID та Organization ID з CRM")
            st.caption(
                "**Person** — метч по **Email** / **Linkedin Person** (довідник people). "
                "**Organization - ID** — метч по **Website**, якщо немає — по **Linkedin Company** "
                "(база **1 — Організації CRM**, як формула на PPL)."
            )
            if leads_have_org_lookup_columns(comp_df) and not _get_org_lookup():
                st.warning(
                    "Завантажте **1 — Організації CRM** у бічній панелі — "
                    "інакше **Organization - ID** не заповниться."
                )
            person_lookup = _get_person_lookup()
            if not person_lookup and not persons_crm_exists(ROOT):
                st.warning(
                    "Завантажте **Довідник Person CRM** у бічній панелі "
                    "(експорт усіх контактів: Email, LinkedIn, Person - ID)."
                )
            elif not person_lookup:
                st.warning("Перезавантажте CSV persons у бічній панелі.")
            else:
                org_lookup = _get_org_lookup()
                person_review_df, person_review_annotated, person_review_xlsx = (
                    _get_cached_person_leads_review(
                        st.session_state.comp_df, org_lookup
                    )
                )
                pstats = person_leads_review_stats(person_review_annotated)
                pc1, pc2, pc3, pc4 = st.columns(4)
                pc1.metric("Person - ID", pstats["with_person_id"])
                pc2.metric("Person - ID by Email", pstats["with_email_id"])
                pc3.metric("ID by Linkedin Person", pstats["with_linkedin_id"])
                pc4.metric("Проблемних", pstats["issues"])
                if st.button(
                    "Перерахувати Person + Organization ID",
                    key="recalc_person_ids_leads",
                    type="primary",
                ):
                    _sync_comp_person_ids(force=True)
                    _sync_comp_org_ids(force=True)
                    st.rerun()

                org_filled = 0
                if "Organization - ID" in comp_df.columns:
                    org_filled = int(
                        (comp_df["Organization - ID"].astype(str).str.strip() != "").sum()
                    )
                st.caption(f"**Organization - ID** заповнено: **{org_filled}** / {len(comp_df)}")

                highlighted_n = pstats["issues"]
                total_n = len(person_review_df)
                if highlighted_n == 0:
                    st.success("Проблемних рядків для Person ID немає.")
                else:
                    st.warning(
                        f"**{highlighted_n}** з **{total_n}** рядків потребують уваги "
                        f"(**CRM by Email ID** / **CRM Linkedin ID** — як у Google PPL)."
                    )

                    def _highlight_person_contact_cols(row: pd.Series):
                        issue = str(
                            person_review_annotated.loc[row.name, PERSON_LOOKUP_ISSUE_COLUMN]
                        ).strip()
                        if str(
                            person_review_annotated.loc[row.name, PERSON_HIGHLIGHT_COLUMN]
                        ).strip() != "yes":
                            return [""] * len(row)
                        styles = [""] * len(row)
                        for i, col in enumerate(row.index):
                            if col == CRM_BY_EMAIL_ID_COLUMN and (
                                f"Підсвічено: {CRM_BY_EMAIL_ID_COLUMN}" in issue
                            ):
                                styles[i] = "background-color: #ff6b6b"
                            elif col == PERSON_CRM_LINKEDIN_COLUMN and (
                                f"Підсвічено: {PERSON_CRM_LINKEDIN_COLUMN}" in issue
                            ):
                                styles[i] = "background-color: #ff6b6b"
                            elif col == PERSON_ID_BY_EMAIL_COLUMN and (
                                f"Підсвічено: {PERSON_ID_BY_EMAIL_COLUMN}" in issue
                            ):
                                styles[i] = "background-color: #f4cccc"
                            elif col == ID_BY_LINKEDIN_PERSON_COLUMN and (
                                f"Підсвічено: {ID_BY_LINKEDIN_PERSON_COLUMN}" in issue
                            ):
                                styles[i] = "background-color: #f4cccc"
                        return styles

                    highlight_mask = (
                        person_review_annotated[PERSON_HIGHLIGHT_COLUMN].astype(str).str.strip()
                        == "yes"
                    )
                    preview_review = person_review_df.loc[highlight_mask].head(50)
                    if preview_review.empty:
                        preview_review = person_review_df.head(50)
                    st.dataframe(
                        preview_review.style.apply(
                            _highlight_person_contact_cols, axis=1
                        ),
                        use_container_width=True,
                        height=220,
                    )
                    st.caption(
                        "**Excel (.xlsx)** — червона заливка **CRM by Email ID** / **CRM Linkedin ID**, "
                        f"рожева **{PERSON_ID_BY_EMAIL_COLUMN}** / **{ID_BY_LINKEDIN_PERSON_COLUMN}**. "
                        "**CSV** — ті самі 24 колонки PPL без службових полів."
                    )
                    dl1, dl2 = st.columns(2)
                    with dl1:
                        st.download_button(
                            label=f"Excel з червоною підсвіткою ({total_n})",
                            data=person_review_xlsx,
                            file_name=leads_person_xlsx_path.name,
                            mime=(
                                "application/vnd.openxmlformats-officedocument"
                                ".spreadsheetml.sheet"
                            ),
                            key="download_leads_person_issues_xlsx",
                            type="primary",
                        )
                    with dl2:
                        st.download_button(
                            label=f"CSV усіх лідів ({total_n})",
                            data=dataframe_to_csv_bytes(person_review_df),
                            file_name=leads_person_path.name,
                            mime="text/csv",
                            key="download_leads_person_issues_csv",
                        )

            st.divider()
            st.subheader("Крок 4 — ліди для CRM")
            opt_notion, opt_person_id, opt_activity = st.columns([2, 2, 1])
            with opt_person_id:
                st.checkbox(
                    f"Додати **{PERSON_ID_COLUMN}**",
                    key="include_person_id_in_export",
                    help="Колонка з'являється лише з галочкою. Завжди є у внутрішніх даних для матчу з activities.",
                )
            with opt_notion:
                st.checkbox(
                    f"Додати **{NOTION_CAMPAIGN_ID_COLUMN}**",
                    key="fill_notion_campaign_id",
                    help=(
                        "Колонка з’являється лише з галочкою. **Research ID** для рядка, "
                        "де **Activity Subject** = колонка з дослідження (якщо є), інакше **ICP name**; "
                        "для event → **Conference**."
                    ),
                    on_change=_on_leads_export_options_change,
                )
            with opt_activity:
                st.text_input(
                    ACTIVITY_TYPE_COLUMN,
                    key="leads_default_activity_type",
                    help=(
                        "Якщо в дослідженні **Activity Type** порожній — підставляється це "
                        "значення для всіх рядків."
                    ),
                    on_change=_on_leads_export_options_change,
                )
            comp_df = _align_leads_for_display(st.session_state.comp_df)
            if st.session_state.fill_notion_campaign_id:
                if NOTION_CAMPAIGN_ID_COLUMN in comp_df.columns:
                    notion_n = int(
                        comp_df[NOTION_CAMPAIGN_ID_COLUMN].astype(str).str.strip().ne("").sum()
                    )
                    st.caption(
                        f"**{NOTION_CAMPAIGN_ID_COLUMN}**: заповнено **{notion_n}** "
                        f"з **{len(comp_df)}**."
                    )
            # Person - ID: приховуємо з файлу якщо галочка не поставлена
            comp_df_export = comp_df.copy()
            if not st.session_state.get("include_person_id_in_export", False):
                comp_df_export = comp_df_export.drop(columns=[PERSON_ID_COLUMN], errors="ignore")

            st.dataframe(comp_df_export, use_container_width=True, height=420)
            st.download_button(
                label="Скачати CRM Leads CSV",
                data=dataframe_to_csv_bytes(comp_df_export),
                file_name=comp_path.name,
                mime="text/csv",
            )

    elif st.session_state.active_tab == "companies":
        if st.session_state.companies_df is None:
            st.warning(
                "Спочатку згенеруйте компанії **з дослідження зверху** (Google / CSV)."
            )
        else:
            if _sync_companies_org_ids():
                st.info(
                    "Підставлено **Organization - ID**, **Website ID**, **Linkedin ID**, "
                    "**CRM Website ID**, **CRM Linkedin ID** з бази організацій."
                )
            companies_df = align_companies_to_crm_schema(st.session_state.companies_df)
            dupe_keys = duplicate_company_name_keys(companies_df)
            dupe_rows = 0
            if DUPLICATE_NAME_COLUMN in companies_df.columns:
                dupe_rows = int((companies_df[DUPLICATE_NAME_COLUMN] == "yes").sum())

            st.subheader("Крок 1 — дублікати назв Company")
            st.caption(
                f"У **companies для CRM** уже **1 рядок на website** (дублікати домену з дослідження "
                f"зливаються). Файл дублікатів — лише коли **одна й та сама назва Company** "
                f"залишилась у **2+ рядках** (часто різні website), "
                f"**{DUPLICATE_NAME_COLUMN}** = yes. Злиття після правок — **{APOLLO_ACCOUNT_ID_COLUMN}**."
            )
            if dupe_keys:
                st.warning(
                    f"**{len(dupe_keys)}** назв повторюються · **{dupe_rows}** рядків для перевірки."
                )
                st.dataframe(
                    build_duplicate_name_report(companies_df),
                    use_container_width=True,
                    hide_index=True,
                )
                dupes_df = duplicate_company_rows(companies_df)
                st.download_button(
                    label="Скачати CSV дублікатів (yes)",
                    data=dataframe_to_csv_bytes(dupes_df),
                    file_name=companies_dupes_path.name,
                    mime="text/csv",
                )
                st.caption(
                    f"Після правок у CSV застосовується лише колонка **{COMPANY_COLUMN}** — "
                    "ID та інші поля з файлу ігноруються."
                )
                dupes_upload = st.file_uploader(
                    "Відредагований CSV дублікатів",
                    type=["csv"],
                    key="companies_duplicates_upload",
                )
                if st.button(
                    "Застосувати зміни до companies",
                    disabled=dupes_upload is None,
                    key="merge_companies_duplicates",
                ):
                    try:
                        edited = load_source_csv(dupes_upload.getvalue())
                        merged, stats, changes = merge_companies_by_apollo_id(
                            st.session_state.companies_df,
                            edited,
                            merge_columns=DUPLICATES_MERGE_COLUMNS,
                        )
                        st.session_state.companies_df = merged
                        st.session_state.companies_merge_preview = normalize_merge_changes_df(
                            changes
                        )
                        _invalidate_companies_review_cache()
                        _persist_workspace()
                        st.success(
                            f"Злито **{stats['updated']}** рядків за Apollo Account id · "
                            f"змін у полях: **{stats['changed_fields']}** "
                            f"(**{stats['changed_rows']}** організацій). "
                            f"Не знайдено: **{stats['not_found']}** · "
                            f"без id: **{stats['skipped_no_id']}**."
                        )
                        if stats["not_found"]:
                            st.warning(
                                "Частина Apollo id з файлу не збіглася з поточною таблицею companies."
                            )
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Помилка злиття: {exc}")
            else:
                st.success("Дублікатів назв немає — можна скачувати повний файл компаній.")

            _render_companies_merge_preview()

            st.divider()
            st.subheader("Крок 2 — ID з бази CRM (замість формул Comp)")
            st.info(
                "**Етап 1:** база **«1 — Організації CRM»** — перша вигрузка з CRM → метч частини ID → "
                "**Крок 3** — імпорт **companies** у CRM.  \n"
                "**Етап 2:** нова вигрузка **organizations** з CRM → **замінити** файл у sidebar → "
                "кнопка **«Перерахувати ID»** нижче — тоді ID з’являться для щойно створених компаній."
            )
            st.caption(
                "Колонка **Organization lookup issue** — де Website/LinkedIn не знайшлись або URL ≠ CRM."
            )
            with st.expander("Які колонки підставляються", expanded=False):
                st.markdown(format_organization_csv_help())
                st.markdown(format_company_export_help())

            lookup = _get_org_lookup()
            if not lookup and not organizations_crm_exists(ROOT):
                st.warning(
                    "Завантажте **1 — Організації CRM** у бічній панелі "
                    "(Organization - Website / LinkedIn / ID)."
                )
            elif not lookup:
                st.warning("Перезавантажте CSV організацій у бічній панелі.")
            else:
                companies_review_df, companies_review_xlsx = _get_cached_companies_review(
                    companies_df
                )
                stats = organization_review_stats(companies_review_df)
                c_org1, c_org2, c_org3, c_org4 = st.columns(4)
                c_org1.metric("Organization - ID", stats["with_org_id"])
                c_org2.metric("Website ID", stats["with_website_id"])
                c_org3.metric("Linkedin ID", stats["with_linkedin_id"])
                c_org4.metric("Підсвічено", stats["highlighted"])
                st.caption(
                    f"У таблиці **{stats['total']}** компаній · "
                    f"CRM Website: **{stats['with_crm_website']}** · "
                    f"CRM Linkedin: **{stats['with_crm_linkedin']}** · "
                    f"усього сигналів: **{stats['issues']}**"
                )
                if st.button(
                    "Перерахувати ID з бази організацій",
                    key="recalc_org_ids_companies",
                    type="primary",
                ):
                    _sync_companies_org_ids(force=True)
                    st.rerun()

                org_issues_df = companies_review_df.loc[
                    companies_review_df[ORG_LOOKUP_ISSUE_COLUMN].astype(str).str.strip() != ""
                ].copy()
                highlighted_n = stats["highlighted"]
                total_n = len(companies_review_df)
                if highlighted_n == 0 and stats["issues"] == 0:
                    st.success(
                        "Усі рядки з Website/LinkedIn мають узгоджені ID "
                        "(проблемних для вигрузки немає)."
                    )
                else:
                    if highlighted_n:
                        st.warning(
                            f"**{highlighted_n}** з **{total_n}** компаній — підсвітка "
                            f"**{CRM_WEBSITE_ID_COLUMN}** / **{CRM_LINKEDIN_ID_COLUMN}** "
                            f"(як у Google). "
                            f"У CSV: **{ORG_HIGHLIGHT_COLUMN}** = `yes`, "
                            f"деталі — **{ORG_LOOKUP_ISSUE_COLUMN}**."
                        )
                    if stats["issues"] > highlighted_n:
                        st.info(
                            f"Ще **{stats['issues'] - highlighted_n}** рядків з іншими "
                            f"сигналами (різні ID, немає Organization - ID тощо) — "
                            f"див. CSV проблемних нижче."
                        )

                    display_cols = companies_review_xlsx_columns(companies_review_df)

                    def _highlight_org_contact_cols(row: pd.Series):
                        full = companies_review_df.loc[row.name]
                        if str(full.get(ORG_HIGHLIGHT_COLUMN, "")).strip() != "yes":
                            return [""] * len(row)
                        issue = str(full.get(ORG_LOOKUP_ISSUE_COLUMN, ""))
                        styles = [""] * len(row)
                        for i, col in enumerate(row.index):
                            if (
                                col == CRM_WEBSITE_ID_COLUMN
                                and "Підсвічено: CRM Website ID" in issue
                            ):
                                styles[i] = "background-color: #ff6b6b"
                            elif (
                                col == CRM_LINKEDIN_ID_COLUMN
                                and "Підсвічено: CRM Linkedin ID" in issue
                            ):
                                styles[i] = "background-color: #ff6b6b"
                            elif (
                                col == WEBSITE_ID_COLUMN
                                and "Підсвічено: Website ID" in issue
                            ):
                                styles[i] = "background-color: #f4cccc"
                            elif (
                                col == LINKEDIN_ID_COLUMN
                                and "Підсвічено: Linkedin ID" in issue
                            ):
                                styles[i] = "background-color: #f4cccc"
                        return styles

                    preview_review = companies_review_df.loc[
                        companies_review_df[ORG_HIGHLIGHT_COLUMN] == "yes"
                    ].head(50)
                    if preview_review.empty:
                        preview_review = companies_review_df.head(50)
                    preview_display = preview_review[display_cols]
                    st.dataframe(
                        preview_display.style.apply(
                            _highlight_org_contact_cols, axis=1
                        ),
                        use_container_width=True,
                        height=220,
                    )
                    st.caption(
                        f"**Excel (.xlsx)** — {len(display_cols)} колонок Comp: червоні "
                        f"**{CRM_WEBSITE_ID_COLUMN}** / **{CRM_LINKEDIN_ID_COLUMN}**, "
                        f"рожеві **{WEBSITE_ID_COLUMN}** / **{LINKEDIN_ID_COLUMN}** "
                        f"(як у Google). Деталі проблем — у **CSV проблемних**."
                    )
                    dl_xlsx, dl_review_csv, dl_issues = st.columns(3)
                    with dl_xlsx:
                        st.download_button(
                            label=f"Excel з підсвіткою ({total_n})",
                            data=companies_review_xlsx,
                            file_name=companies_review_xlsx_path.name,
                            mime=(
                                "application/vnd.openxmlformats-officedocument"
                                ".spreadsheetml.sheet"
                            ),
                            key="download_companies_review_xlsx",
                            type="primary",
                        )
                    with dl_review_csv:
                        review_csv_df = companies_review_df[display_cols]
                        st.download_button(
                            label=f"CSV усіх компаній ({total_n})",
                            data=dataframe_to_csv_bytes(review_csv_df),
                            file_name=companies_review_xlsx_path.with_suffix(
                                ".csv"
                            ).name,
                            mime="text/csv",
                            key="download_companies_review_csv",
                        )
                    with dl_issues:
                        st.download_button(
                            label=f"CSV проблемних ({len(org_issues_df)})",
                            data=dataframe_to_csv_bytes(org_issues_df),
                            file_name=companies_org_path.name,
                            mime="text/csv",
                            key="download_companies_org_issues",
                        )
                    st.caption(
                        f"Ключ злиття — **{APOLLO_ACCOUNT_ID_COLUMN}**. "
                        "ID-колонки з CSV не перезаписуються; після правок оновлюються з CRM lookup."
                    )
                    org_fix_upload = st.file_uploader(
                        "Відредагований CSV (проблемні ID)",
                        type=["csv"],
                        key="companies_org_issues_upload",
                    )
                    if st.button(
                        "Застосувати правки + оновити ID",
                        disabled=org_fix_upload is None,
                        key="merge_companies_org_issues",
                    ):
                        try:
                            edited = load_source_csv(org_fix_upload.getvalue())
                            merged, merge_stats, changes = merge_companies_by_apollo_id(
                                st.session_state.companies_df, edited
                            )
                            merged = apply_organization_lookup(
                                merged, lookup, keep_existing_org_id=True
                            )
                            st.session_state.companies_df = merged
                            st.session_state.companies_merge_preview = (
                                normalize_merge_changes_df(changes)
                            )
                            _invalidate_companies_review_cache()
                            _persist_workspace()
                            remaining = len(organization_problem_rows(merged))
                            st.success(
                                f"Злито **{merge_stats['updated']}** рядків · "
                                f"залишилось проблемних: **{remaining}**."
                            )
                            if merge_stats["not_found"]:
                                st.warning(
                                    "Частина Apollo id з файлу не збіглася з таблицею companies."
                                )
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Помилка: {exc}")

            st.divider()
            st.subheader("Крок 3 — усі компанії для CRM")
            companies_crm_df = companies_for_crm_file(companies_df)
            if dupe_keys:

                def _highlight_duplicate_names(row):
                    name_key = str(row.get("Company", "")).strip().casefold()
                    if name_key in dupe_keys:
                        return ["background-color: #fff3cd"] * len(row)
                    return [""] * len(row)

                st.dataframe(
                    companies_crm_df.style.apply(_highlight_duplicate_names, axis=1),
                    use_container_width=True,
                    height=360,
                )
            else:
                st.dataframe(companies_crm_df, use_container_width=True, height=360)
            step3_csv, step3_xlsx = st.columns(2)
            with step3_csv:
                st.download_button(
                    label="Скачати companies CSV (імпорт CRM)",
                    data=dataframe_to_csv_bytes(companies_crm_df),
                    file_name=companies_path.name,
                    mime="text/csv",
                )
            with step3_xlsx:
                if lookup:
                    _, step3_review_xlsx = _get_cached_companies_review(companies_df)
                    st.download_button(
                        label=f"Excel з підсвіткою {CRM_WEBSITE_ID_COLUMN} / {CRM_LINKEDIN_ID_COLUMN}",
                        data=step3_review_xlsx,
                        file_name=companies_review_xlsx_path.name,
                        mime=(
                            "application/vnd.openxmlformats-officedocument"
                            ".spreadsheetml.sheet"
                        ),
                        key="download_companies_step3_xlsx",
                    )
                else:
                    st.caption("Excel з підсвіткою — після завантаження організацій CRM.")
            st.caption(
                "Після імпорту цього файлу в CRM: вигрузіть **новий** organizations CSV → "
                "sidebar **«1 — Організації CRM»** → **Крок 2 — Перерахувати ID**."
            )

if st.session_state.get("active_tab") == "activities":
    st.header("CRM Activities")
    activities_file = st.file_uploader(
        "Файл активностей з CRM (CSV)",
        type=["csv"],
        key="activities_upload",
    )
    _act_cache_key = None
    if activities_file is not None:
        _act_cache_key = (
            f"{activities_file.name}:{activities_file.size}|{_activities_research_cache_key()}|{_activities_leads_cache_key()}"
        )
    _needs_merge = (
        activities_file is not None
        and (
            st.session_state.get("activities_cache_key") != _act_cache_key
            or st.session_state.get("merged_activities_df") is None
        )
    )
    if _needs_merge:
        with st.spinner("Злиття з дослідженням…"):
            merged_df = _sync_activities_merge(activities_file)
    else:
        merged_df = _sync_activities_merge(activities_file)
    if merged_df is not None:
        research_df = _activities_research_df()
        if activities_file is None and st.session_state.get("activities_file_name"):
            st.info(
                f"Використовується збережений результат злиття: "
                f"**{st.session_state.activities_file_name}** "
                f"({len(merged_df)} рядків)."
            )
        elif activities_file is not None:
            st.success(
                f"Завантажено **{activities_file.name}**: "
                f"**{len(merged_df)}** рядків · **{len(merged_df.columns)}** колонок."
            )

        if research_df is None:
            st.warning(
                "Дослідження не завантажено — колонки з research CSV будуть порожніми. "
                "Завантаж дослідження зверху (Google Таблиця або CSV дослідження)."
            )
        elif st.session_state.get("comp_df") is not None:
            st.caption(
                f"**{PERSON_LINKEDIN_ACTIVE_COLUMN}** — з **Leads for CRM** (comp), "
                "якщо людина знайдена по Email / Apollo Contact id; "
                "інакше лишається значення з CRM activities. "
                "Оновити можна файлом нижче."
            )

        st.divider()
        st.subheader("LinkedIn activity → Person - Linkedin Active")
        st.caption(
            "Той самий CSV з тулзи, що на **Leads → Крок 2**. "
            "Оновлює **comp_df** і поточний злитий файл activities (по Email / Apollo id). "
            "**yes** — активність < 1 міс.; **no** — ≥ 1 міс.; **порожньо** — профілю не було у файлі."
        )
        act_linkedin_file = st.file_uploader(
            "Результати тулзи LinkedIn (CSV)",
            type=["csv"],
            key="activities_linkedin_activity_upload",
        )
        act_li_col1, act_li_col2 = st.columns(2)
        with act_li_col1:
            if st.button(
                "Додати Person - Linkedin Active",
                disabled=act_linkedin_file is None or merged_df is None,
                key="activities_apply_linkedin_activity",
            ):
                try:
                    result = _apply_linkedin_activity_file(
                        act_linkedin_file.getvalue(), act_linkedin_file.name
                    )
                    stats = result.get("comp") or result.get("activities_direct") or {}
                    msg = (
                        f"**yes:** {stats.get('yes', 0)} · **no** (у файлі): {stats.get('no_from_activity', 0)} · "
                        f"**порожньо** (не у файлі): {stats.get('not_in_activity_file', 0)}"
                    )
                    if result.get("activities_from_leads") is not None:
                        msg += f" · activities оновлено: **{result['activities_from_leads']}**"
                    st.success(msg)
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
        with act_li_col2:
            if st.button(
                "Синхронізувати з Leads (comp)",
                disabled=merged_df is None or st.session_state.get("comp_df") is None,
                key="activities_sync_linkedin_from_leads",
                help="Без нового файлу — підтягнути Person - Linkedin Active з уже оновленого comp_df.",
            ):
                count = _sync_activities_linkedin_from_leads()
                _persist_workspace()
                st.success(f"Оновлено **{count}** рядків у activities з Leads (comp).")
                st.rerun()

        st.divider()
        stats = matched_stats(merged_df)
        mc1, mc2, mc3 = st.columns(3)
        mc1.metric("Рядків", stats["total"])
        mc2.metric("Contact starter знайдено", stats["contact_starter_matched"])
        mc3.metric("Research ID знайдено", stats["research_id_matched"])

        st.dataframe(merged_df.head(25), use_container_width=True, height=300)

        _act_stem = st.session_state.file_stem or DEFAULT_FILE_STEM
        _xlsx_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        dl1, dl2 = st.columns(2)
        dl1.download_button(
            "⬇ Скачати Activities Excel (всі + підсвітка)",
            data=merged_to_excel_bytes(merged_df),
            file_name=f"{_act_stem} - activities all highlighted.xlsx",
            mime=_xlsx_mime,
        )
        dl2.download_button(
            "⬇ Скачати Issues Excel (лише проблемні)",
            data=issues_to_excel_bytes(merged_df),
            file_name=f"{_act_stem} - activities issues.xlsx",
            mime=_xlsx_mime,
        )

        st.divider()
        st.subheader("Виключити людей з фінального файлу")
        st.caption(
            "Завантаж **робочий del-файл** (як у Google work) або простий список **Person - ID**. "
            "Рядки будуть прибрані з фінального CRM. Окремо — **activities delete** (9 колонок) "
            "напряму з робочого файлу."
        )

        exclude_file = st.file_uploader(
            "Робочий del-файл або Person - ID для виключення",
            type=["csv", "xlsx"],
            key="exclude_person_ids_file",
        )

        exclude_result = _sync_activities_exclude(merged_df, exclude_file)
        crm_source_df = merged_df
        if exclude_result is not None:
            crm_source_df, excluded_df, excluded_count, delete_source_df = exclude_result
            st.success(
                f"Видалено **{excluded_count}** рядків. "
                f"Залишилось **{len(crm_source_df)}** з **{len(merged_df)}**."
            )
            if excluded_count > 0:
                activities_delete_df = build_activities_deletion_export(
                    excluded_df,
                    research_df,
                    delete_source_df=delete_source_df,
                )
                src_label = (
                    "робочий del-файл"
                    if delete_source_df is not None
                    else "merged activities"
                )
                st.caption(
                    f"**Файл для видалення:** **{len(activities_delete_df)}** рядків · "
                    f"**{len(activities_delete_df.columns)}** колонок · джерело: **{src_label}**."
                )
                st.dataframe(
                    activities_delete_df.head(50),
                    use_container_width=True,
                    height=320,
                )
                del_col1, del_col2 = st.columns(2)
                with del_col1:
                    st.download_button(
                        "⬇ Скачати файл для видалення (CRM)",
                        data=activities_deletion_to_excel_bytes(
                            excluded_df,
                            research_df,
                            delete_source_df=delete_source_df,
                        ),
                        file_name=f"{_act_stem} - activities delete.xlsx",
                        mime=_xlsx_mime,
                        type="primary",
                    )
                with del_col2:
                    st.download_button(
                        "⬇ Скачати Activities без виключених",
                        data=merged_to_excel_bytes(crm_source_df),
                        file_name=f"{_act_stem} - activities filtered.xlsx",
                        mime=_xlsx_mime,
                    )
            else:
                st.info(
                    "Жоден **Person - ID** з файлу не знайдено у поточних activities — "
                    "файл для видалення порожній."
                )

        st.divider()
        st.subheader("Фінальний файл для CRM")
        st.caption(
            "17 колонок для імпорту в CRM. Якщо вище застосовано виключення — "
            "у файлі буде менше рядків."
        )
        st.checkbox(
            f"Додати **{PERSON_LINKEDIN_ACTIVE_COLUMN}**",
            key="activities_include_linkedin_active",
            help=(
                "Опційна колонка після **Person - Research ID** "
                "(yes / no / порожньо з Leads або CRM activities)."
            ),
        )
        include_linkedin_active = st.session_state.get(
            "activities_include_linkedin_active", False
        )
        activities_crm_df = build_final_export(
            crm_source_df,
            research_df,
            include_linkedin_active=include_linkedin_active,
        )
        li_filled = 0
        if include_linkedin_active and PERSON_LINKEDIN_ACTIVE_COLUMN in activities_crm_df.columns:
            li_filled = int(
                activities_crm_df[PERSON_LINKEDIN_ACTIVE_COLUMN]
                .astype(str)
                .str.strip()
                .ne("")
                .sum()
            )
            st.caption(
                f"**{PERSON_LINKEDIN_ACTIVE_COLUMN}**: заповнено **{li_filled}** "
                f"з **{len(activities_crm_df)}**."
            )
        st.dataframe(activities_crm_df.head(50), use_container_width=True, height=360)
        st.caption(
            f"Превʼю: до **50** з **{len(activities_crm_df)}** рядків · "
            f"**{len(activities_crm_df.columns)}** колонок."
        )
        st.download_button(
            "⬇ Скачати фінальний файл для CRM",
            data=final_to_excel_bytes(
                crm_source_df,
                research_df,
                include_linkedin_active=include_linkedin_active,
            ),
            file_name=f"{_act_stem} - activities CRM.xlsx",
            mime=_xlsx_mime,
            type="primary",
        )

