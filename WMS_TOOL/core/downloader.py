from __future__ import annotations

from pathlib import Path


def download_labels(args) -> Path:
    from pipeline import download_pdf

    return download_pdf(args)
