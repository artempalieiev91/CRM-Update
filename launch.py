"""Запуск Streamlit-додатку (зручно для PyCharm: Run → launch.py)."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def main() -> None:
    app = ROOT / "app.py"
    ui_port = os.environ.get("CRM_UPDATE_STREAMLIT_PORT", "8501").strip() or "8501"
    health_port = os.environ.get("CRM_UPDATE_HEALTH_PORT", "8502").strip() or "8502"
    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(app),
        "--server.port",
        ui_port,
        *sys.argv[1:],
    ]
    print("Запуск:", " ".join(cmd))
    print(f"UI:      http://localhost:{ui_port}")
    print(f"Health:  http://localhost:{health_port}/health")
    print(f"Long:    http://localhost:{health_port}/health/long")
    raise SystemExit(subprocess.call(cmd, cwd=ROOT))


if __name__ == "__main__":
    main()
