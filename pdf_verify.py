from __future__ import annotations

import argparse
from pathlib import Path

from barcode_verify_tracking import (
    VerifyResult,
    error_result,
    load_metadata_index,
    load_rows,
    load_wms_tracking,
    load_zxingcpp,
    print_result,
    print_summary,
    process_row,
)
from config import DEFAULT_MAX_PAGES
from utils import log


def build_verify_namespace(args, input_dir: Path, download_log: Path) -> argparse.Namespace:
    return argparse.Namespace(
        input_dir=str(input_dir),
        strict_json=args.strict_json,
        download_log=str(download_log),
        output_dir=args.output_dir,
        output_name=args.output_name,
        limit=args.limit,
        offset=args.offset,
        rotations=args.rotations,
        dpi=args.dpi,
        max_pages=args.max_pages or DEFAULT_MAX_PAGES,
        timeout=args.timeout,
        ocr=getattr(args, "ocr", False),
        ocr_enabled=getattr(args, "ocr_enabled", getattr(args, "ocr", False)),
        only_source_only=args.only_source_only,
        debug=args.debug,
    )


def verify_pdfs(args, input_dir: Path, download_log: Path) -> list[VerifyResult]:
    verify_args = build_verify_namespace(args, input_dir, download_log)
    wms_records = load_wms_tracking(download_log)
    zxingcpp = load_zxingcpp()

    # 加载 metadata 索引
    metadata_text = str(getattr(args, "metadata", "") or "").strip()
    metadata_path = Path(metadata_text) if metadata_text else None
    if not metadata_path or not metadata_path.exists() or metadata_path.is_dir():
        from config import BASE_DIR
        metadata_path = BASE_DIR / "output" / "download_label_metadata.jsonl"
    if metadata_path.exists():
        meta_by_file, meta_by_dno = load_metadata_index(metadata_path)
        log(f"已加载 metadata 索引: {len(meta_by_file)} 条文件映射")
    else:
        meta_by_file, meta_by_dno = {}, {}
        log("未找到 metadata.jsonl，跳过 metadata 丰富")

    rows = load_rows(verify_args, wms_records, meta_by_file, meta_by_dno)
    if not rows:
        log("没有需要识别的 PDF/图片面单")
        return []

    log(f"开始识别校验 PDF/图片面单，总数: {len(rows)}")
    results: list[VerifyResult] = []
    for index, row in enumerate(rows, start=1):
        try:
            result = process_row(row, verify_args, zxingcpp, wms_records)
        except Exception as exc:
            log(f"单文件处理失败: {row.get('file_name') or row.get('file_path')} reason={exc}")
            result = error_result(row, wms_records, exc)
        results.append(result)
        print_result(index, len(rows), result)
    print_summary(results)
    return results
