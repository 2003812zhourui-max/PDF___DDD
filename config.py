from __future__ import annotations

import os
import sys
from pathlib import Path


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
CURRENT_WORK_DIR = Path(os.getcwd()).resolve()
DEFAULT_LOG_DIR = BASE_DIR / "logs"
DEFAULT_INPUT_DIR = BASE_DIR / "pdf_downloads"
DEFAULT_OUTPUT_DIR = BASE_DIR / "output" / "pdf"
DEFAULT_DOWNLOAD_LOG = DEFAULT_LOG_DIR / "download_log.csv"
DEFAULT_STORAGE_STATE = BASE_DIR / "wms_storage_state.json"
DEFAULT_DPI = 200
DEFAULT_TIMEOUT = 30
DEFAULT_MAX_PAGES = 1
