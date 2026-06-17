from __future__ import annotations

import argparse
import sys
from pathlib import Path

from config import (
    BASE_DIR,
    CURRENT_WORK_DIR,
    DEFAULT_DPI,
    DEFAULT_DOWNLOAD_LOG,
    DEFAULT_MAX_PAGES,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_STORAGE_STATE,
    DEFAULT_TIMEOUT,
)
from task_pipeline import tracked_download_pdf, tracked_extract_data, tracked_run_ocr
from utils import exe_pause_if_frozen, log


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PDF 面单下载 + 条码识别 + 三信号校验一条龙工具")
    parser.add_argument("--start-time", default="", help="下载开始时间，例如 2026-06-01 00:00:00")
    parser.add_argument("--end-time", default="", help="下载结束时间，例如 2026-06-02 23:59:59")
    parser.add_argument("--wh-codes", default="", help="仓库代码，多个用英文逗号分隔")
    parser.add_argument("--statuses", default="15", help="WMS 状态码，默认 15；待处理通常是 10")
    parser.add_argument("--platforms", default="", help="平台参数预留，当前下载脚本未直接支持时仅记录")
    parser.add_argument("--input-dir", default="", help="已有 PDF 输入目录；传入后跳过下载")
    parser.add_argument("--strict-json", default="", help="已有严格识别 JSON；传入后按 JSON 继续条码补验")
    parser.add_argument("--storage-state", default="", help="WMS 登录态 JSON 文件路径")
    parser.add_argument("--download-log", default=str(DEFAULT_DOWNLOAD_LOG), help="下载日志 CSV，包含 expressNo/filePath")
    parser.add_argument("--download-name", default="", help="下载批次名；不传时默认跟 output-name 一致")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="Excel/JSON 输出目录")
    parser.add_argument("--output-name", default="pdf_label_pipeline_result", help="输出文件名前缀")
    parser.add_argument("--rotations", action="store_true", help="启用 90/180/270 度旋转条码识别")
    parser.add_argument("--dpi", type=int, default=DEFAULT_DPI, help="PDF 渲染 DPI，默认 200")
    parser.add_argument("--max-pages", type=int, default=DEFAULT_MAX_PAGES, help="最多渲染/提取页数，默认 1")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="单 PDF 处理超时秒数，默认 30")
    parser.add_argument("--ocr", action="store_true", help="启用 OCR 辅助识别面单角标")
    parser.add_argument("--limit", type=int, default=0, help="抽样处理数量，0 表示不限")
    parser.add_argument("--offset", type=int, default=0, help="跳过前 N 条，方便断点续跑")
    parser.add_argument("--only-source-only", action="store_true", help="只处理 source_only/低置信度 PDF")
    parser.add_argument("--debug", action="store_true", help="输出详细日志")
    parser.add_argument("--username", default="", help="WMS 登录账号（HTTP 并发模式需要）")
    parser.add_argument("--password", default="", help="WMS 登录密码（HTTP 并发模式需要）")
    parser.add_argument("--workers", type=int, default=8, help="HTTP 并发下载线程数，默认 8")
    parser.add_argument("--download-retries", type=int, default=5, help="单订单下载最大尝试次数，默认 5")
    parser.add_argument("--retry-base-delay", type=float, default=0.8, help="下载/列表重试基础等待秒数，默认 0.8")
    parser.add_argument("--channel", default="", help="物流渠道筛选，例如 TikTok-CBT-US、Upload_Shipping_Label-Speedx")
    parser.add_argument("--metadata", default="", help="metadata.jsonl 路径，默认自动检测 output/download_label_metadata.jsonl")
    parser.add_argument("--task-state", default="", help="任务状态 JSONL 路径；默认写入输出目录的 _intermediate")
    parser.add_argument("--browser-mode", action="store_true", help="使用浏览器兼容模式下载（需要 Playwright）")
    parser.add_argument("--force", action="store_true", help="强制重新下载，不跳过下载日志里已有成功记录")
    return parser.parse_args()


def resolve_storage_state_path(path_text: str) -> Path:
    candidates: list[Path] = []
    if path_text:
        candidates.append(Path(path_text).expanduser().resolve())
    candidates.append(DEFAULT_STORAGE_STATE.resolve())
    candidates.append((CURRENT_WORK_DIR / "wms_storage_state.json").resolve())

    unique_candidates: list[Path] = []
    seen: set[str] = set()
    for path in candidates:
        key = str(path).lower()
        if key not in seen:
            unique_candidates.append(path)
            seen.add(key)

    for path in unique_candidates:
        if path.exists():
            return path
    return unique_candidates[0]


def main() -> int:
    args = parse_args()
    try:
        args.ocr_enabled = bool(args.ocr)
        storage_state = resolve_storage_state_path(args.storage_state)
        args.storage_state = str(storage_state)
        output_dir = Path(args.output_dir).expanduser().resolve()
        log("启动 PDF 面单三步流水线")
        log(f"当前工作目录: {CURRENT_WORK_DIR}")
        log(f"EXE 所在目录/项目目录: {BASE_DIR}")
        log(f"storage_state 实际路径: {storage_state}")
        log(f"storage_state 是否存在: {storage_state.exists()}")
        if args.platforms:
            log(f"平台参数已记录: {args.platforms}")

        download_manifest = tracked_download_pdf(args)
        ocr_results = tracked_run_ocr(args, download_manifest)
        extract_manifest = tracked_extract_data(args, ocr_results)
        log(f"识别输出目录: {output_dir}")
        log(f"三步流水线完成: {extract_manifest}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1
    finally:
        exe_pause_if_frozen()


if __name__ == "__main__":
    raise SystemExit(main())
