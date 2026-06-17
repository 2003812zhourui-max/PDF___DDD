from __future__ import annotations

import os
import sys
from pathlib import Path


def load_local_env(env_path: Path) -> None:
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if not key or key in os.environ:
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
            value = value[1:-1]
        os.environ[key] = value


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
CURRENT_WORK_DIR = Path(os.getcwd()).resolve()
load_local_env(BASE_DIR / ".env")
if CURRENT_WORK_DIR != BASE_DIR:
    load_local_env(CURRENT_WORK_DIR / ".env")
DEFAULT_LOG_DIR = BASE_DIR / "logs"
DEFAULT_INPUT_DIR = BASE_DIR / "pdf_downloads"
DEFAULT_OUTPUT_DIR = BASE_DIR / "output" / "pdf"
DEFAULT_DOWNLOAD_LOG = DEFAULT_LOG_DIR / "download_log.csv"
DEFAULT_STORAGE_STATE = BASE_DIR / "wms_storage_state.json"
DEFAULT_DPI = 200
DEFAULT_TIMEOUT = 30
DEFAULT_MAX_PAGES = 1
