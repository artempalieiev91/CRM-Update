"""IDS operations: logging, marker file, short/long health tests (CodeIT policy)."""

from __future__ import annotations

import logging
import os
import secrets
import threading
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Callable
from urllib.parse import parse_qs, urlparse

ROOT = Path(__file__).resolve().parent
DEFAULT_MARKER = ROOT / ".crm_update_workspace" / "status.marker"
DEFAULT_HEALTH_PORT = 8502
LOGGER_NAME = "crm_update.ids"

_log = logging.getLogger(LOGGER_NAME)
_started = False
_start_lock = threading.Lock()
_server: ThreadingHTTPServer | None = None


def _env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def marker_path() -> Path:
    raw = _env("CRM_UPDATE_MARKER_PATH")
    return Path(raw) if raw else DEFAULT_MARKER


def health_port() -> int:
    raw = _env("CRM_UPDATE_HEALTH_PORT", str(DEFAULT_HEALTH_PORT))
    try:
        return int(raw)
    except ValueError:
        return DEFAULT_HEALTH_PORT


def health_host() -> str:
    return _env("CRM_UPDATE_HEALTH_HOST", "0.0.0.0") or "0.0.0.0"


def health_api_key() -> str:
    return _env("CRM_UPDATE_HEALTH_API_KEY")


def setup_logging(*, force: bool = False) -> None:
    """Stdout + optional file. Safe to call more than once."""
    logger = logging.getLogger("crm_update")
    if logger.handlers and not force:
        return
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter(
        "%(asctime)s %(levelname)s [%(name)s] %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S%z",
    )
    stream = logging.StreamHandler()
    stream.setFormatter(fmt)
    logger.addHandler(stream)
    log_file = _env("CRM_UPDATE_LOG_FILE")
    if log_file:
        path = Path(log_file)
        path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(path, encoding="utf-8")
        file_handler.setFormatter(fmt)
        logger.addHandler(file_handler)
    logger.propagate = False


def write_marker(text: str) -> None:
    path = marker_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    body = f"{text.strip()}\nupdated: {stamp}\n"
    path.write_text(body, encoding="utf-8")


def mark_ok() -> None:
    write_marker("OK")


def mark_error(message: str) -> None:
    write_marker(f"ERROR: {message.strip()}")


def mark_warning(message: str) -> None:
    write_marker(f"WARNING: {message.strip()}")


def log_key_action(
    action: str,
    *,
    ok: bool,
    detail: str = "",
    update_marker: bool = True,
) -> None:
    """Log a policy key action and optionally refresh the marker file."""
    extra = f" — {detail}" if detail else ""
    if ok:
        _log.info("%s%s", action, extra)
        if update_marker:
            mark_ok()
        return
    _log.error("%s%s", action, extra)
    if update_marker:
        mark_error(f"{action}{extra}")


def _workspace_dir() -> Path:
    return ROOT / ".crm_update_workspace"


def run_short_test() -> str:
    """Overall stance. Returns 'ok' when the process can serve tests."""
    return "ok"


def run_long_test(*, skip_google: bool | None = None) -> tuple[bool, str]:
    """Deeper checks: disk write, marker, optional Google HTTPS."""
    errors: list[str] = []
    workspace = _workspace_dir()
    try:
        workspace.mkdir(parents=True, exist_ok=True)
        probe = workspace / ".ids_long_test"
        probe.write_text("ok\n", encoding="utf-8")
        if probe.read_text(encoding="utf-8").strip() != "ok":
            errors.append("workspace probe file mismatch")
        probe.unlink(missing_ok=True)
    except OSError as exc:
        errors.append(f"cannot write workspace ({workspace}): {exc}")

    try:
        write_marker("OK")
        if not marker_path().is_file():
            errors.append("marker file was not created")
    except OSError as exc:
        errors.append(f"cannot write marker file: {exc}")

    if skip_google is None:
        skip_google = _env("CRM_UPDATE_LONG_TEST_SKIP_GOOGLE").lower() in {
            "1",
            "true",
            "yes",
        }
    if not skip_google:
        try:
            import requests

            r = requests.get("https://docs.google.com", timeout=8)
            if r.status_code >= 500:
                errors.append(f"Google HTTPS status {r.status_code}")
        except Exception as exc:  # noqa: BLE001 — long test must report any failure
            errors.append(f"cannot reach docs.google.com: {exc}")

    if errors:
        msg = "; ".join(errors)
        mark_error(msg)
        return False, msg
    mark_ok()
    return True, "ok"


def _authorized(handler: BaseHTTPRequestHandler) -> bool:
    expected = health_api_key()
    if not expected:
        return True
    provided = handler.headers.get("X-API-Key", "") or ""
    parsed = urlparse(handler.path)
    query = parse_qs(parsed.query)
    if not provided:
        provided = (query.get("api_key") or [""])[0]
    if len(provided) != len(expected):
        return False
    return secrets.compare_digest(provided, expected)


class _HealthHandler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args) -> None:  # noqa: A003
        _log.info("health %s", format % args)

    def _send(self, code: int, body: str) -> None:
        payload = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self) -> None:  # noqa: N802
        path = urlparse(self.path).path.rstrip("/") or "/"
        if path not in {"/health", "/health/short", "/healthz", "/health/long"}:
            self._send(404, "not found")
            return
        if not _authorized(self):
            self._send(401, "unauthorized")
            return
        if path == "/health/long":
            ok, message = run_long_test()
            self._send(200 if ok else 503, message)
            return
        self._send(200, run_short_test())

    def do_HEAD(self) -> None:  # noqa: N802
        self.do_GET()


def start_health_server() -> ThreadingHTTPServer:
    host = health_host()
    port = health_port()
    server = ThreadingHTTPServer((host, port), _HealthHandler)
    thread = threading.Thread(target=server.serve_forever, name="crm-update-health", daemon=True)
    thread.start()
    _log.info("health server listening on http://%s:%s/health", host, port)
    return server


def ensure_ops_started(*, root: Path | None = None) -> None:
    """Idempotent: logging, marker OK, health HTTP server."""
    global _started, _server
    del root  # reserved for callers; workspace is always project ROOT
    with _start_lock:
        if _started:
            return
        setup_logging()
        try:
            _server = start_health_server()
        except OSError as exc:
            log_key_action(
                "health server bind",
                ok=False,
                detail=f"{health_host()}:{health_port()}: {exc}",
            )
            _started = True
            return
        mark_ok()
        _log.info("CRM Update IDS started; marker=%s", marker_path())
        if not health_api_key():
            _log.warning(
                "CRM_UPDATE_HEALTH_API_KEY is empty — health tests are open. "
                "Set the key before production."
            )
        _started = True


def log_disk_write(path: Path | str, *, action: str = "write disk file") -> None:
    log_key_action(action, ok=True, detail=str(path))


def log_disk_failure(path: Path | str, exc: BaseException, *, action: str = "write disk file") -> None:
    log_key_action(action, ok=False, detail=f"{path}: {exc}")


def logged_write(action: str, path: Path | str, writer: Callable[[], None]) -> None:
    try:
        writer()
    except Exception as exc:
        log_disk_failure(path, exc, action=action)
        raise
    log_disk_write(path, action=action)
