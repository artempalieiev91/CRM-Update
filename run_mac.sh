#!/usr/bin/env bash
# Запуск CRM Update на macOS: venv, залежності, Streamlit.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

PY="${PYTHON:-python3}"
if ! command -v "$PY" &>/dev/null; then
  echo "Потрібен Python 3 (наприклад: brew install python). Команда не знайдена: $PY" >&2
  exit 1
fi

if [[ ! -d .venv ]]; then
  echo "Створення .venv…"
  "$PY" -m venv .venv
fi
# shellcheck source=/dev/null
source .venv/bin/activate
python -m pip install -q -U pip
python -m pip install -q -r requirements.txt

echo "Відкрийте в браузері адресу, яку покаже Streamlit (зазвичай http://localhost:8501)."
exec streamlit run app.py "$@"
