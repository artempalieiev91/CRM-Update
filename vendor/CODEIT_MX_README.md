# CodeIT MX Provider Checker

## Файли

| Файл | Призначення |
|------|-------------|
| `vendor/mx_provider_checker_router.py` | FastAPI-роутер CodeIT (UI, task API) |
| `vendor/mx_checker_codeit_source.py` | Оригінал `lib/mx_checker.py` з Workspace |
| `mx_checker.py` | Та сама DNS/MX-логіка для Streamlit (без `lib.file_handler`) |

## Логіка (з `mx_checker_codeit_source.py`)

- Домен з колонки **Email** (`normalize_domain` → частина після `@`)
- DNS MX через `dnspython`, nameservers `8.8.8.8`, `1.1.1.1`, …
- `PROVIDER_PATTERNS` → `mail_provider` (Google Workspace, Microsoft 365, Other, …)
- CRM: `Microsoft 365` → `outlook`, решта визначених провайдерів → `Other (gmail, etc)`, помилки/немає MX → порожньо
- `workers=30`, `timeout=5` (як у роутері)

## Відмінності від веб-версії CodeIT

- Немає фонових `task_id` / `ResultBuffer` — синхронний `fill_email_providers()` у `crm_export.py` та `app.py`
- Колонка результату — `Organization - Email provider`, не `mail_provider`
