"""
MX Provider Checker — логіка з CodeIT lib/mx_checker.py (mail_checker).

Домен з колонки Email → DNS MX → mail_provider → CRM (outlook / Other (gmail, etc)).
"""

from __future__ import annotations

import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable
from urllib.parse import urlparse

from ids_runtime import log_key_action

import dns.resolver
import pandas as pd

EMAIL_COLUMN = "Email"
CRM_PROVIDER_COLUMN = "Organization - Email provider"
MAIL_PROVIDER_COLUMN = "mail_provider"

PROVIDER_OUTLOOK = "outlook"
PROVIDER_OTHER = "Other (gmail, etc)"

DEFAULT_WORKERS = 30
DEFAULT_TIMEOUT_SEC = 5

# З CodeIT mx_checker.py
PROVIDER_PATTERNS: dict[str, list[str]] = {
    "Google Workspace": [".google.com", ".googlemail.com", ".smtp.goog"],
    "Microsoft 365": [
        ".outlook.com",
        ".protection.outlook.com",
        ".mail.protection.outlook.com",
        ".mail.protection.outlook.",
        ".mx.microsoft",
        ".ppe-hosted.com",
    ],
    "Yandex Mail": [".yandex.net", ".yandex.ru"],
    "Mail.ru": [".mail.ru", ".mxs.mail.ru"],
    "Zoho Mail": [".zoho.com", ".zoho.in", ".zoho.eu"],
    "GoDaddy": [".secureserver.net", ".email.secureserver.net"],
    "Amazon SES": [".amazonaws.com", ".awsapps.com"],
    "ProtonMail": [".protonmail.ch"],
    "FastMail": [".fastmail.com", ".messagingengine.com"],
    "Namecheap": [".privateemail.com"],
    "Cloudflare": [".mx.cloudflare.net"],
    "Apple iCloud": [".icloud.com"],
    "Rackspace": [".emailsrvr.com"],
    "SendGrid": [".sendgrid.net"],
    "Mailgun": [".mailgun.org"],
    "Hostinger": [".hostinger.com", ".titan.email"],
    "cPanel": [".websitewelcome.com"],
    "Plesk": [".antispamcloud.com"],
    "Mimecast": [".mimecast.com"],
    "Barracuda": [".barracudanetworks.com"],
    "SpamTitan": [".spamtitan.com"],
    "Cisco IronPort": [".iphmx.com"],
}

_DOMAIN_RE = re.compile(
    r"^(?:[a-zA-Z0-9](?:[a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?\.)*"
    r"[a-zA-Z0-9](?:[a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?$"
)

AUTO_DETECT_KEYWORDS = ["domain", "url", "website", "site", "email"]

_CODEIT_ERROR_PROVIDERS = frozenset(
    {
        "",
        "Invalid Domain",
        "No MX Records",
        "Domain Not Found",
        "DNS Query Timeout",
        "Max Retries Exceeded",
        "Not Checked",
    }
)

_BLANK_VALUES = frozenset({"", "nan", "none", "null", "<na>", "na", "n/a"})


def _is_blank(value: object) -> bool:
    return str(value or "").strip().casefold() in _BLANK_VALUES


def validate_domain(domain: str) -> bool:
    if not domain:
        return False
    return bool(_DOMAIN_RE.match(domain))


def normalize_domain(url_or_domain: object) -> str | None:
    """З CodeIT — email → домен після @."""
    if url_or_domain is None or (isinstance(url_or_domain, float) and pd.isna(url_or_domain)):
        return None

    raw = str(url_or_domain).strip().lower()
    if not raw:
        return None

    raw = re.sub(r"^mailto:", "", raw)

    if "@" in raw:
        parts = raw.split("@")
        if len(parts) == 2:
            raw = parts[1]

    if "." not in raw:
        return None

    if not raw.startswith(("http://", "https://", "ftp://")):
        if "/" in raw:
            raw = "http://" + raw

    if raw.startswith(("http://", "https://", "ftp://")):
        try:
            parsed = urlparse(raw)
            domain = parsed.netloc
        except Exception:
            domain = raw
    else:
        domain = raw

    domain = re.sub(r"^www\.", "", domain)
    domain = domain.split(":")[0]
    domain = domain.split("/")[0]

    return domain if validate_domain(domain) else None


def auto_detect_column(columns: list[str] | pd.Index) -> str | None:
    """З CodeIT — перша колонка з domain/url/website/email у назві."""
    for col in columns:
        col_lower = str(col).lower()
        if any(kw in col_lower for kw in AUTO_DETECT_KEYWORDS):
            return str(col)
    return None


def get_mail_provider(
    domain: str,
    resolver: dns.resolver.Resolver,
    *,
    retries: int = 3,
) -> tuple[str, str, list[str]]:
    """З CodeIT — MX lookup + PROVIDER_PATTERNS."""
    if not domain:
        return "", "Invalid Domain", []

    for attempt in range(retries):
        try:
            mx_records = resolver.resolve(domain, "MX")
            mx_sorted = sorted(
                (r.preference, str(r.exchange).lower().rstrip(".")) for r in mx_records
            )
            mx_hostnames = [mx[1] for mx in mx_sorted]

            for provider_name, patterns in PROVIDER_PATTERNS.items():
                for pattern in patterns:
                    if any(pattern in hostname for hostname in mx_hostnames):
                        return domain, provider_name, mx_hostnames

            if mx_hostnames:
                return domain, "Other", mx_hostnames
            return domain, "No MX Records", []

        except dns.resolver.NXDOMAIN:
            return domain, "Domain Not Found", []
        except dns.resolver.NoAnswer:
            return domain, "No MX Records", []
        except dns.resolver.Timeout:
            if attempt < retries - 1:
                time.sleep(1)
                continue
            return domain, "DNS Query Timeout", []
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(1)
                continue
            return domain, f"Error: {str(e)}", []

    return domain, "Max Retries Exceeded", []


def codeit_provider_to_crm(provider: str) -> str:
    """mail_provider (CodeIT) → Organization - Email provider (CRM)."""
    name = str(provider or "").strip()
    if not name or name in _CODEIT_ERROR_PROVIDERS or name.startswith("Error:"):
        return ""
    if name == "Microsoft 365":
        return PROVIDER_OUTLOOK
    return PROVIDER_OTHER


def _make_resolver(timeout: int) -> dns.resolver.Resolver:
    resolver = dns.resolver.Resolver()
    resolver.timeout = timeout
    resolver.lifetime = timeout
    resolver.nameservers = ["8.8.8.8", "8.8.4.4", "1.1.1.1", "1.0.0.1"]
    return resolver


def build_domain_provider_map(
    domains: list[str],
    *,
    workers: int = DEFAULT_WORKERS,
    timeout: int = DEFAULT_TIMEOUT_SEC,
    on_progress: Callable[[int, int], None] | None = None,
) -> dict[str, str]:
    """Унікальні домени → CRM provider (логіка CodeIT get_mail_provider)."""
    crm_map, _ = build_domain_provider_map_with_raw(
        domains, workers=workers, timeout=timeout, on_progress=on_progress
    )
    return crm_map


def build_domain_provider_map_with_raw(
    domains: list[str],
    *,
    workers: int = DEFAULT_WORKERS,
    timeout: int = DEFAULT_TIMEOUT_SEC,
    on_progress: Callable[[int, int], None] | None = None,
) -> tuple[dict[str, str], dict[str, str]]:
    """Повертає (crm_map, raw_map): crm_map — CRM значення, raw_map — сирий DNS результат."""
    unique = sorted({normalize_domain(d) for d in domains if normalize_domain(d)})
    if not unique:
        return {}, {}

    resolver = _make_resolver(timeout)
    workers = max(1, min(100, workers))
    crm_results: dict[str, str] = {}
    raw_results: dict[str, str] = {}
    lock = threading.Lock()
    total = len(unique)

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_domain = {
            executor.submit(get_mail_provider, domain, resolver): domain
            for domain in unique
        }
        done = 0
        for future in as_completed(future_to_domain):
            domain = future_to_domain[future]
            _, provider, _ = future.result()
            crm = codeit_provider_to_crm(provider)
            with lock:
                crm_results[domain] = crm
                raw_results[domain] = provider
            done += 1
            if on_progress:
                on_progress(done, total)

    return crm_results, raw_results


MX_RAW_COLUMN = "mx_provider_raw"


def fill_email_providers(
    df: pd.DataFrame,
    *,
    source_column: str | None = None,
    provider_column: str = CRM_PROVIDER_COLUMN,
    raw_column: str | None = MX_RAW_COLUMN,
    overwrite: bool = False,
    workers: int = DEFAULT_WORKERS,
    timeout: int = DEFAULT_TIMEOUT_SEC,
    on_progress: Callable[[int, int], None] | None = None,
) -> pd.DataFrame:
    """Заповнює CRM-колонку; домен лише з Email (як у вашому CodeIT-процесі).

    raw_column — якщо вказано, зберігає сирий DNS-результат (для діагностики).
    """
    out = df.copy()
    if provider_column not in out.columns:
        out[provider_column] = ""

    email_col = source_column or EMAIL_COLUMN
    if email_col not in out.columns:
        detected = auto_detect_column(out.columns)
        if detected and "email" in str(detected).lower():
            email_col = detected
        else:
            return out

    work = out.copy()
    work["_normalized_domain"] = work[email_col].apply(normalize_domain)

    check_mask = work["_normalized_domain"].notna()
    if not overwrite:
        check_mask &= work[provider_column].map(_is_blank)

    unique_domains = work.loc[check_mask, "_normalized_domain"].dropna().unique().tolist()
    log_key_action(
        "MX DNS lookup start",
        ok=True,
        detail=f"{len(unique_domains)} domains",
        update_marker=False,
    )
    try:
        domain_map, raw_map = build_domain_provider_map_with_raw(
            unique_domains,
            workers=workers,
            timeout=timeout,
            on_progress=on_progress,
        )
    except Exception as exc:
        log_key_action("MX DNS lookup", ok=False, detail=str(exc))
        raise
    log_key_action("MX DNS lookup", ok=True, detail=f"{len(domain_map)} resolved")

    def _crm_value(row: pd.Series) -> str:
        if not overwrite and not _is_blank(row.get(provider_column)):
            return str(row[provider_column]).strip()
        norm = row.get("_normalized_domain")
        if pd.isna(norm) or not norm:
            return ""
        crm = domain_map.get(str(norm), "")
        if not crm:
            # DNS помилка, але домен є → fallback до Other (gmail, etc)
            return PROVIDER_OTHER
        return crm

    def _raw_value(row: pd.Series) -> str:
        norm = row.get("_normalized_domain")
        if pd.isna(norm) or not norm:
            return ""
        return raw_map.get(str(norm), "")

    out[provider_column] = work.apply(_crm_value, axis=1)
    if raw_column:
        out[raw_column] = work.apply(_raw_value, axis=1)
    return out
