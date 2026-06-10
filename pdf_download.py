from __future__ import annotations

import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from config import BASE_DIR, DEFAULT_DOWNLOAD_LOG, DEFAULT_INPUT_DIR, DEFAULT_LOG_DIR
from utils import log


DEFAULT_OUTPUT_NAME = "pdf_label_pipeline_result"


@dataclass
class DownloadResult:
    input_dir: Path
    download_log: Path
    skipped: bool


def has_download_params(args) -> bool:
    return bool(args.start_time or args.end_time or args.wh_codes or getattr(args, "platforms", ""))


def sanitize_name(value: str) -> str:
    value = re.sub(r'[\\/:*?"<>|]+', "_", value or "")
    value = re.sub(r"\s+", "_", value).strip("_")
    return value or "batch"


def batch_name(args) -> str:
    if args.output_name and args.output_name != DEFAULT_OUTPUT_NAME:
        return sanitize_name(args.output_name)
    if args.start_time or args.end_time:
        text = f"{args.start_time}_{args.end_time}"
        text = re.sub(r"[^0-9A-Za-z]+", "_", text).strip("_")
        return text or datetime.now().strftime("batch_%Y%m%d_%H%M%S")
    return datetime.now().strftime("batch_%Y%m%d_%H%M%S")


def default_batch_dir(args) -> Path:
    return DEFAULT_INPUT_DIR / batch_name(args)


def is_default_download_log(path_text: str) -> bool:
    return Path(path_text).expanduser().resolve() == DEFAULT_DOWNLOAD_LOG.resolve()


def resolve_download_log(args, target_dir: Path) -> Path:
    if args.download_log and not is_default_download_log(args.download_log):
        return Path(args.download_log).expanduser().resolve()
    return DEFAULT_LOG_DIR / f"{batch_name(args)}_download_log.csv"


def run_download_http(args) -> DownloadResult:
    """使用 auto_download.py 的纯 HTTP 并发下载（无需浏览器，并发提速）"""
    import os

    input_dir = Path(args.input_dir).expanduser().resolve() if args.input_dir else default_batch_dir(args).resolve()
    input_dir.mkdir(parents=True, exist_ok=True)
    DEFAULT_LOG_DIR.mkdir(parents=True, exist_ok=True)
    download_log = DEFAULT_LOG_DIR / "download_log.csv"

    username = getattr(args, "username", "") or os.environ.get("WMS_USERNAME", "")
    password = getattr(args, "password", "") or os.environ.get("WMS_PASSWORD", "")
    if not username or not password:
        raise RuntimeError("HTTP 并发模式需要账号密码，请设置 --username/--password 或环境变量 WMS_USERNAME/WMS_PASSWORD")

    workers = getattr(args, "workers", 5)

    downloader_args = [
        "auto_download.py",
        "--username", username,
        "--password", password,
        "--pdf-dir", str(input_dir),
        "--workers", str(workers),
    ]
    if args.start_time and args.end_time:
        downloader_args.extend(["--start-time", args.start_time, "--end-time", args.end_time])
    elif getattr(args, "start_time", ""):
        downloader_args.extend(["--start-time", args.start_time])
    if args.wh_codes:
        downloader_args.extend(["--wh-codes", args.wh_codes])
    if getattr(args, "statuses", ""):
        downloader_args.extend(["--status", args.statuses])
    if args.limit and args.limit > 0:
        downloader_args.extend(["--limit", str(args.limit)])
    if getattr(args, "channel", ""):
        downloader_args.extend(["--channel", args.channel])

    log(f"开始 HTTP 并发下载 PDF 面单，保存目录: {input_dir}")
    log(f"并发线程: {workers}")
    display_args = downloader_args[1:].copy()
    if "--password" in display_args:
        password_index = display_args.index("--password") + 1
        if password_index < len(display_args):
            display_args[password_index] = "******"
    log("下载参数: " + " ".join(display_args))

    import auto_download

    old_argv = sys.argv[:]
    try:
        sys.argv = downloader_args
        auto_download.main()
    finally:
        sys.argv = old_argv

    log(f"下载完成，PDF目录: {input_dir}")
    log(f"下载日志: {download_log}")
    return DownloadResult(input_dir=input_dir, download_log=download_log, skipped=False)


def run_download_browser(args) -> DownloadResult:
    """使用原 Playwright 浏览器模式下载（兼容旧流程）"""
    download_requested = has_download_params(args) or not args.input_dir
    input_dir = Path(args.input_dir).expanduser().resolve() if args.input_dir else default_batch_dir(args).resolve()

    if not download_requested:
        log(f"跳过下载，使用已有 PDF 目录: {input_dir}")
        return DownloadResult(input_dir=input_dir, download_log=Path(args.download_log).expanduser().resolve(), skipped=True)

    input_dir.mkdir(parents=True, exist_ok=True)
    DEFAULT_LOG_DIR.mkdir(parents=True, exist_ok=True)
    download_log = resolve_download_log(args, input_dir)
    validation_log = DEFAULT_LOG_DIR / f"{batch_name(args)}_pdf_validation_log.csv"

    downloader_args = [
        "batch_download_wms_pdfs.py",
        "--pdf-dir",
        str(input_dir),
        "--log-file",
        str(download_log),
        "--pdf-validation-log-file",
        str(validation_log),
    ]
    if args.start_time:
        downloader_args.extend(["--start-time", args.start_time])
    if args.end_time:
        downloader_args.extend(["--end-time", args.end_time])
    if args.wh_codes:
        downloader_args.extend(["--wh-codes", args.wh_codes])
    if getattr(args, "statuses", ""):
        downloader_args.extend(["--statuses", args.statuses])
    if getattr(args, "storage_state", ""):
        downloader_args.extend(["--storage-state", args.storage_state])
    if args.limit and args.limit > 0:
        downloader_args.extend(["--total-limit", str(args.limit)])
    if getattr(args, "channel", ""):
        downloader_args.extend(["--channel", args.channel])
    if args.debug:
        downloader_args.append("--browser-mode")

    log(f"开始下载 PDF 面单，保存目录: {input_dir}")
    log("下载参数: " + " ".join(downloader_args[1:]))
    if getattr(sys, "frozen", False):
        import batch_download_wms_pdfs

        old_argv = sys.argv[:]
        try:
            sys.argv = downloader_args
            batch_download_wms_pdfs.main()
        finally:
            sys.argv = old_argv
    else:
        script = BASE_DIR / "batch_download_wms_pdfs.py"
        if not script.exists():
            raise FileNotFoundError(f"下载脚本不存在: {script}")
        command = [sys.executable, str(script), *downloader_args[1:]]
        completed = subprocess.run(command, cwd=str(BASE_DIR), check=False)
        if completed.returncode != 0:
            raise RuntimeError(f"PDF 下载失败，退出码: {completed.returncode}")

    log(f"下载完成，PDF目录: {input_dir}")
    log(f"下载日志: {download_log}")
    return DownloadResult(input_dir=input_dir, download_log=download_log, skipped=False)


def run_download(args) -> DownloadResult:
    """自动选择下载模式：默认 HTTP 并发，传 --browser-mode 用旧浏览器模式"""
    if not has_download_params(args) and args.input_dir:
        input_dir = Path(args.input_dir).expanduser().resolve()
        log(f"跳过下载，使用已有 PDF 目录: {input_dir}")
        return DownloadResult(input_dir=input_dir, download_log=Path(args.download_log).expanduser().resolve(), skipped=True)

    if getattr(args, "browser_mode", False):
        log("使用浏览器模式下载（需要 Playwright）")
        return run_download_browser(args)
    else:
        try:
            return run_download_http(args)
        except Exception as exc:
            log(f"HTTP 并发模式失败: {exc}")
            if has_download_params(args) or not args.input_dir:
                log("回退到浏览器模式...")
                return run_download_browser(args)
            raise
