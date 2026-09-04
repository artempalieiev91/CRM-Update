# CRM Update

Internal Streamlit service (IDS) that prepares CRM import files from Google Sheets / research CSV: companies, leads (PPL), LinkedIn URLs, activities.

This README is the handoff file required by **CodeIT Common requirements for internally developed services** (v1.0).

## Location (fill in on the server)

| Item | Value |
|------|--------|
| Server | `[HOSTNAME]` |
| Directory | `[/path/to/crm-update]` |
| UI URL | `http://[HOST]:8501` |
| Health (short test) | `http://[HOST]:8502/health` |
| Health (long test) | `http://[HOST]:8502/health/long` |
| Marker file | `[DIR]/.crm_update_workspace/status.marker` (or `CRM_UPDATE_MARKER_PATH`) |

IT: add this service to Infra Wiki and the operability monitor using the URLs above.

## Does it launch automatically?

**Default: no.** It is an interactive UI. Start it with `python launch.py` (or the systemd unit below if IT deploys it as always-on).

If systemd is enabled, treat it as a daemon: short/long tests and the marker file apply.

## How to validate a successful launch

1. Short test (overall stance) — must return `ok`:

```bash
curl -sS -H "X-API-Key: $CRM_UPDATE_HEALTH_API_KEY" http://127.0.0.1:8502/health
```

2. UI responds: open `http://[HOST]:8501` (Streamlit). Stopgap: `GET http://[HOST]:8501/_stcore/health`.

3. Marker file contains a first line `OK`.

4. Long test (disk write + Google HTTPS):

```bash
curl -sS -H "X-API-Key: $CRM_UPDATE_HEALTH_API_KEY" http://127.0.0.1:8502/health/long
```

Returns `ok` (HTTP 200) or an error description (HTTP 503). API key is required when `CRM_UPDATE_HEALTH_API_KEY` is set. Query string `?api_key=` is also accepted.

## Run

Python 3.9+.

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example /etc/crm-update.env   # then set CRM_UPDATE_HEALTH_API_KEY
set -a && source /etc/crm-update.env && set +a
.venv/bin/python launch.py
```

Local without env file:

```bash
.venv/bin/python launch.py
```

UI: http://localhost:8501  
Health: http://localhost:8502/health

## Logs (key actions)

Stdout (journald if systemd). Optional file: `CRM_UPDATE_LOG_FILE`.

Logged actions:

- Google Sheets fetch (start / success / failure)
- Workspace, organizations CRM, persons CRM disk writes
- Export CSV writes (leads, companies, LinkedIn)
- MX DNS lookups
- Health server bind

On failure the marker file is set to `ERROR: …`. On success of a key action it is set back to `OK`.

## Marker file

Agreed default: `.crm_update_workspace/status.marker`

| Contents | Meaning |
|----------|---------|
| `OK` | No error on the last key action |
| `ERROR: …` | Key action failed (Google fetch, disk write, health bind, long test) |
| `WARNING: …` | Reserved for degraded-but-running cases |

## systemd example (always-on)

```ini
[Unit]
Description=CRM Update IDS
After=network.target

[Service]
Type=simple
WorkingDirectory=/path/to/crm-update
EnvironmentFile=/etc/crm-update.env
ExecStart=/path/to/crm-update/.venv/bin/python launch.py
Restart=on-failure
User=crm-update

[Install]
WantedBy=multi-user.target
```

Pass the health URL and API key to IT so they can add the service to the operability monitor. **Do not release until that is done** (company policy).

## Operator notes (українською)

Сервіс готує файли для CRM: компанії, ліди, LinkedIn, activities. У UI: Google / CSV зверху, довідники організацій і persons у sidebar. Health-порти потрібні IT для моніторингу; у повсякденній роботі користуйтесь http://localhost:8501.
