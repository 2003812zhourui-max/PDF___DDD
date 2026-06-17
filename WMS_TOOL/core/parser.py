from __future__ import annotations

from pathlib import Path


def parse_labels(args, download_manifest_path: Path) -> Path:
    from pipeline import run_ocr

    return run_ocr(args, download_manifest_path)
