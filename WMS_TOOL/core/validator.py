from __future__ import annotations

from pathlib import Path


def export_excel(args, ocr_results_path: Path) -> Path:
    from pipeline import extract_data

    return extract_data(args, ocr_results_path)
