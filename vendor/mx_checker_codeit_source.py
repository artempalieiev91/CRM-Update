"""
MX Provider Checker — DNS MX record lookup and email provider detection.

Adapted from mx_provider_checker/mail_checker.py for integration into
the codeit-workspace FastAPI application.
"""

import dns.resolver
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import threading
import re
import uuid
from dataclasses import dataclass, field
from typing import Tuple, List, Optional, Dict, Any
from urllib.parse import urlparse

from lib.file_handler import get_dataframe, store_dataframe
from lib.result_buffer import ResultBuffer
from lib.settings import logger


# ---------------------------------------------------------------------------
# Provider detection
# ---------------------------------------------------------------------------

PROVIDER_PATTERNS: Dict[str, List[str]] = {
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
    r'^(?:[a-zA-Z0-9](?:[a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?\.)*'
    r'[a-zA-Z0-9](?:[a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?$'
)

AUTO_DETECT_KEYWORDS = ['domain', 'url', 'website', 'site', 'email']


# ---------------------------------------------------------------------------
# Domain helpers
# ---------------------------------------------------------------------------

def validate_domain(domain: str) -> bool:
    if not domain:
        return False
    return bool(_DOMAIN_RE.match(domain))


def normalize_domain(url_or_domain) -> Optional[str]:
    """Extract a clean domain from a URL, email address, or plain domain string."""
    if not url_or_domain or pd.isna(url_or_domain):
        return None

    raw = str(url_or_domain).strip().lower()
    if not raw:
        return None

    # Strip mailto:
    raw = re.sub(r'^mailto:', '', raw)

    # Email → take domain part
    if '@' in raw:
        parts = raw.split('@')
        if len(parts) == 2:
            raw = parts[1]

    if '.' not in raw:
        return None

    # Add scheme for urlparse if needed
    if not raw.startswith(('http://', 'https://', 'ftp://')):
        if '/' in raw:
            raw = 'http://' + raw

    # Parse URL
    if raw.startswith(('http://', 'https://', 'ftp://')):
        try:
            parsed = urlparse(raw)
            domain = parsed.netloc
        except Exception:
            domain = raw
    else:
        domain = raw

    # Clean up
    domain = re.sub(r'^www\.', '', domain)
    domain = domain.split(':')[0]   # port
    domain = domain.split('/')[0]   # path

    return domain if validate_domain(domain) else None


def get_mail_provider(domain: str, resolver: dns.resolver.Resolver,
                      retries: int = 3) -> Tuple[str, str, List[str]]:
    """Look up MX records and identify the email provider."""
    if not domain:
        return "", "Invalid Domain", []

    for attempt in range(retries):
        try:
            mx_records = resolver.resolve(domain, 'MX')
            mx_sorted = sorted(
                [(r.preference, str(r.exchange).lower().rstrip('.')) for r in mx_records]
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


# ---------------------------------------------------------------------------
# Task manager — background processing with progress + cancellation
# ---------------------------------------------------------------------------

STALE_TASK_TTL = 30 * 60  # 30 minutes

@dataclass
class MXTask:
    task_id: str
    file_id: str
    column: str
    workers: int
    timeout: int
    total: int = 0
    completed: int = 0
    status: str = "starting"  # starting | running | completed | cancelled | error
    error: Optional[str] = None
    result_file_id: Optional[str] = None
    provider_stats: Dict[str, int] = field(default_factory=dict)
    total_records: int = 0
    valid_domains: int = 0
    start_time: float = field(default_factory=time.time)
    cancel_flag: threading.Event = field(default_factory=threading.Event)


_tasks: Dict[str, MXTask] = {}
_tasks_lock = threading.Lock()


def _cleanup_stale_tasks():
    """Remove tasks older than STALE_TASK_TTL."""
    now = time.time()
    stale = [tid for tid, t in _tasks.items() if now - t.start_time > STALE_TASK_TTL]
    for tid in stale:
        _tasks.pop(tid, None)
    if stale:
        logger.info(f"Cleaned up {len(stale)} stale MX tasks")


def _run_mx_check(task_id: str):
    """Background worker that processes domains and updates task progress."""
    task = _tasks.get(task_id)
    if not task:
        return

    try:
        df, filename = get_dataframe(task.file_id)
        if df is None:
            task.error = "Source data expired. Please re-upload."
            task.status = "error"
            return

        # Normalize domains
        task.status = "running"
        df = df.copy()
        df.loc[:, '_normalized_domain'] = df[task.column].apply(normalize_domain)

        unique_domains = df['_normalized_domain'].dropna().unique().tolist()
        task.total = len(unique_domains)
        task.total_records = len(df)
        task.valid_domains = len(unique_domains)

        if not unique_domains:
            task.error = "No valid domains found in the selected column."
            task.status = "error"
            return

        # Set up resolver
        resolver = dns.resolver.Resolver()
        resolver.timeout = task.timeout
        resolver.lifetime = task.timeout
        resolver.nameservers = ['8.8.8.8', '8.8.4.4', '1.1.1.1', '1.0.0.1']

        # Process domains concurrently
        results: Dict[str, Dict[str, str]] = {}

        with ThreadPoolExecutor(max_workers=task.workers) as executor:
            future_to_domain = {
                executor.submit(get_mail_provider, domain, resolver): domain
                for domain in unique_domains
            }

            for future in as_completed(future_to_domain):
                if task.cancel_flag.is_set():
                    executor.shutdown(wait=False, cancel_futures=True)
                    task.status = "cancelled"
                    return

                domain, provider, mx_records = future.result()
                results[domain] = {
                    'provider': provider,
                    'mx_records': ', '.join(mx_records) if mx_records else ''
                }
                task.completed += 1

        # Map results back to full DataFrame
        df.loc[:, 'mail_provider'] = df['_normalized_domain'].map(
            lambda x: results.get(x, {}).get('provider', 'Not Checked') if pd.notna(x) else 'Invalid'
        )
        df.loc[:, 'mx_records'] = df['_normalized_domain'].map(
            lambda x: results.get(x, {}).get('mx_records', '') if pd.notna(x) else ''
        )
        df = df.drop(columns=['_normalized_domain'])

        # Store result
        task.result_file_id = store_dataframe(df, filename)
        task.provider_stats = df['mail_provider'].value_counts().to_dict()
        task.total_records = len(df)
        task.status = "completed"

        # Save to result buffer
        from lib.config import AppConfig
        config = AppConfig()
        suffix = config.get_app('mx_provider_checker').get('output_suffix', 'mx_checked')
        from lib.file_handler import get_download_filename
        output_filename = get_download_filename(filename or 'data.csv', suffix=suffix)
        ResultBuffer().save(df, 'mx_provider_checker', output_filename)

        logger.info(f"MX check {task_id} completed: {task.total} domains in {time.time() - task.start_time:.1f}s")

    except Exception as e:
        task.error = str(e)
        task.status = "error"
        logger.exception(f"MX check {task_id} failed")


def start_task(file_id: str, column: str, workers: int, timeout: int) -> str:
    """Create and start a background MX check task. Returns task_id."""
    with _tasks_lock:
        _cleanup_stale_tasks()

    # Clamp inputs
    workers = max(1, min(100, workers))
    timeout = max(1, min(30, timeout))

    task_id = str(uuid.uuid4())[:12]
    task = MXTask(
        task_id=task_id,
        file_id=file_id,
        column=column,
        workers=workers,
        timeout=timeout,
    )
    _tasks[task_id] = task

    thread = threading.Thread(target=_run_mx_check, args=(task_id,), daemon=True)
    thread.start()

    return task_id


def get_task(task_id: str) -> Optional[MXTask]:
    return _tasks.get(task_id)


def cancel_task(task_id: str) -> bool:
    task = _tasks.get(task_id)
    if task and task.status == "running":
        task.cancel_flag.set()
        return True
    return False


def auto_detect_column(columns: list) -> Optional[str]:
    """Return the best column name for domain data, or None."""
    for col in columns:
        col_lower = col.lower()
        if any(kw in col_lower for kw in AUTO_DETECT_KEYWORDS):
            return col
    return None
