from __future__ import annotations

import csv
import json
from dataclasses import asdict, fields
from pathlib import Path
from typing import Any

from barcode_verify_tracking import VerifyResult
from exporter import export_results
from pdf_download import run_download
from pdf_verify import verify_pdfs
from utils import ensure_dir, log


SUPPORTED_LABEL_SUFFIXES = {".pdf", ".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff"}


def intermediate_dir(args) -> Path:
    base = Path(args.output_dir).expanduser().resolve()
    return base / "_intermediate" / str(args.output_name)


def write_json_file(path: Path, payload: Any) -> Path:
    ensure_dir(path.parent)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def read_json_file(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def files_from_download_log(download_log: Path) -> list[str]:
    if not download_log.exists():
        return []

    paths: list[str] = []
    with download_log.open("r", encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file):
            if str(row.get("status") or "").lower() != "success":
                continue
            for item in str(row.get("filePath") or "").split("|"):
                item = item.strip()
                if item:
                    paths.append(str(Path(item).expanduser().resolve()))
    return list(dict.fromkeys(paths))


def files_from_input_dir(input_dir: Path) -> list[str]:
    if input_dir.is_file():
        return [str(input_dir.resolve())]
    if not input_dir.exists():
        return []

    paths = [
        str(path.resolve())
        for path in sorted(input_dir.rglob("*"))
        if path.is_file() and path.suffix.lower() in SUPPORTED_LABEL_SUFFIXES
    ]
    return paths


def download_pdf(args) -> Path:
    """Step 1: download labels or accept an input directory, then write a path manifest."""
    if args.strict_json:
        raise ValueError("download_pdf() accepts label files, not strict-json results.")

    if args.input_dir and not (args.start_time or args.end_time or args.wh_codes):
        input_dir = Path(args.input_dir).expanduser().resolve()
        download_log = Path(args.download_log).expanduser().resolve()
        skipped = True
    else:
        download_result = run_download(args)
        input_dir = download_result.input_dir
        download_log = download_result.download_log
        skipped = download_result.skipped

    files = files_from_download_log(download_log) or files_from_input_dir(input_dir)
    manifest_path = intermediate_dir(args) / "01_download_manifest.json"
    payload = {
        "step": "download_pdf",
        "input_dir": str(input_dir),
        "download_log": str(download_log),
        "skipped": skipped,
        "files": files,
    }
    write_json_file(manifest_path, payload)
    log(f"download_pdf 中间文件: {manifest_path}")
    return manifest_path


def run_ocr(args, download_manifest_path: Path) -> Path:
    """Step 2: read downloaded file paths and write OCR/text/barcode results."""
    manifest = read_json_file(download_manifest_path)
    input_dir = Path(str(manifest.get("input_dir") or "")).expanduser().resolve()
    download_log = Path(str(manifest.get("download_log") or "")).expanduser().resolve()

    if not input_dir.exists():
        raise FileNotFoundError(f"run_ocr input path does not exist: {input_dir}")

    args.input_dir = str(input_dir)
    args.download_log = str(download_log)
    args.strict_json = ""
    results = verify_pdfs(args, input_dir=input_dir, download_log=download_log)

    ocr_path = intermediate_dir(args) / "02_ocr_results.json"
    payload = {
        "step": "run_ocr",
        "download_manifest": str(download_manifest_path),
        "input_files": manifest.get("files", []),
        "results": [asdict(result) for result in results],
    }
    write_json_file(ocr_path, payload)
    log(f"run_ocr 中间文件: {ocr_path}")
    return ocr_path


def load_verify_results(ocr_results_path: Path) -> list[VerifyResult]:
    payload = read_json_file(ocr_results_path)
    valid_keys = {field.name for field in fields(VerifyResult)}
    results: list[VerifyResult] = []
    for item in payload.get("results", []):
        if not isinstance(item, dict):
            continue
        result_kwargs = {key: item.get(key) for key in valid_keys}
        results.append(VerifyResult(**result_kwargs))
    return results


def extract_data(args, ocr_results_path: Path) -> Path:
    """Step 3: read OCR/text results and export business data."""
    results = load_verify_results(ocr_results_path)
    output_dir = Path(args.output_dir).expanduser().resolve()
    excel_path, json_path = export_results(results, output_dir=output_dir, output_name=args.output_name)

    extract_path = intermediate_dir(args) / "03_extract_manifest.json"
    payload = {
        "step": "extract_data",
        "ocr_results": str(ocr_results_path),
        "excel_path": str(excel_path),
        "json_path": str(json_path),
        "result_count": len(results),
    }
    write_json_file(extract_path, payload)
    log(f"extract_data 中间文件: {extract_path}")
    return extract_path
