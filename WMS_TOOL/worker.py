from __future__ import annotations

import contextlib
import json
import os
import re
import sys
from argparse import Namespace
from datetime import datetime
from pathlib import Path
from typing import Any, Callable


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output" / "excel"
DOWNLOAD_DIR = BASE_DIR / "output" / "downloads"
LOG_DIR = BASE_DIR / "logs"
CONFIG_PATH = BASE_DIR / "config" / "default.json"

Callback = Callable[[dict[str, Any]], None]


def load_defaults() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        return {}
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if not key or key in os.environ:
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        os.environ[key] = value


def load_local_envs() -> None:
    for path in (BASE_DIR / ".env", BASE_DIR.parent / ".env", Path.cwd() / ".env"):
        load_env_file(path)


def sanitize_name(value: str) -> str:
    return re.sub(r"[^0-9A-Za-z_]+", "_", value).strip("_") or "wms_run"


def run_name_for(start_time: str, end_time: str) -> str:
    text = f"{start_time}_{end_time}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return sanitize_name(text)


def emit(callback: Callback | None, event: str, **payload: Any) -> None:
    if callback:
        callback({"event": event, **payload})


class CallbackWriter:
    def __init__(self, callback: Callback | None, log_file, base_progress: int, progress_span: int) -> None:
        self.callback = callback
        self.log_file = log_file
        self.base_progress = base_progress
        self.progress_span = progress_span
        self._buffer = ""

    def write(self, text: str) -> int:
        self.log_file.write(text)
        self.log_file.flush()
        self._buffer += text
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            self._handle_line(line.rstrip())
        return len(text)

    def flush(self) -> None:
        self.log_file.flush()

    def _handle_line(self, line: str) -> None:
        if not line:
            return
        emit(self.callback, "log", message=line)
        match = re.search(r"\[(\d+)/(\d+)\]", line)
        if match:
            done = int(match.group(1))
            total = max(int(match.group(2)), 1)
            progress = self.base_progress + int(self.progress_span * min(done / total, 1.0))
            emit(self.callback, "progress", progress=progress)


def resolve_storage_state() -> Path:
    candidates = [
        BASE_DIR / "wms_storage_state.json",
        BASE_DIR.parent / "wms_storage_state.json",
        Path.cwd() / "wms_storage_state.json",
    ]
    for path in candidates:
        if path.exists():
            return path.resolve()
    return candidates[0].resolve()


def make_args(params: dict[str, Any], run_name: str, metadata_path: Path) -> Namespace:
    defaults = load_defaults()
    output_dir = OUTPUT_DIR
    input_dir = DOWNLOAD_DIR / run_name
    download_log = LOG_DIR / f"{run_name}_download_log.csv"
    task_state = LOG_DIR / f"{run_name}_task_state.jsonl"
    return Namespace(
        start_time=str(params.get("start_time") or "").strip(),
        end_time=str(params.get("end_time") or "").strip(),
        wh_codes=str(params.get("warehouse") or params.get("wh_codes") or defaults.get("warehouse") or "US02").strip(),
        statuses=str(params.get("statuses") or defaults.get("statuses") or "10,15,20,30").strip(),
        platforms="",
        input_dir=str(input_dir),
        strict_json="",
        storage_state=str(resolve_storage_state()),
        download_log=str(download_log),
        download_name=run_name,
        output_dir=str(output_dir),
        output_name=run_name,
        rotations=False,
        dpi=int(params.get("dpi") or defaults.get("dpi") or 200),
        max_pages=int(params.get("max_pages") or defaults.get("max_pages") or 1),
        timeout=int(params.get("timeout") or defaults.get("timeout") or 60),
        ocr=False,
        ocr_enabled=False,
        limit=int(params.get("limit") or defaults.get("limit") or 0),
        offset=0,
        only_source_only=False,
        debug=False,
        username=str(params.get("username") or ""),
        password=str(params.get("password") or ""),
        workers=int(params.get("workers") or defaults.get("workers") or 8),
        download_retries=int(params.get("download_retries") or defaults.get("download_retries") or 5),
        retry_base_delay=float(params.get("retry_base_delay") or defaults.get("retry_base_delay") or 0.8),
        channel=str(params.get("channel") or ""),
        metadata=str(metadata_path),
        task_state=str(task_state),
        browser_mode=False,
        force=bool(params.get("force") or False),
    )


def validate_params(params: dict[str, Any]) -> None:
    start_time = str(params.get("start_time") or "").strip()
    end_time = str(params.get("end_time") or "").strip()
    if not start_time or not end_time:
        raise ValueError("start_time 和 end_time 必须填写")
    datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S")
    datetime.strptime(end_time, "%Y-%m-%d %H:%M:%S")


def run_pipeline(params: dict[str, Any], callback: Callback | None = None) -> dict[str, Any]:
    validate_params(params)
    load_local_envs()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    run_name = run_name_for(str(params["start_time"]), str(params["end_time"]))
    metadata_path = LOG_DIR / f"{run_name}_metadata.jsonl"
    os.environ["PDF_DDD_METADATA_JSONL"] = str(metadata_path)
    os.environ["PDF_DDD_METADATA_SUMMARY"] = str(LOG_DIR / f"{run_name}_metadata_summary.xlsx")
    args = make_args(params, run_name, metadata_path)
    log_path = LOG_DIR / f"{run_name}.log"

    emit(callback, "status", status="running")
    emit(callback, "progress", progress=1)

    with log_path.open("a", encoding="utf-8", newline="\n") as log_file:
        writer = CallbackWriter(callback, log_file, base_progress=10, progress_span=55)
        with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
            print(f"任务开始: {run_name}", flush=True)
            print(f"时间范围: {args.start_time} ~ {args.end_time}", flush=True)
            print(f"仓库: {args.wh_codes}; 状态: {args.statuses}; 并发: {args.workers}", flush=True)

            emit(callback, "progress", progress=3)
            try:
                from WMS_TOOL.core.wms_client import query_totals

                totals = query_totals(
                    start_time=args.start_time,
                    end_time=args.end_time,
                    wh_codes=[item.strip() for item in args.wh_codes.split(",") if item.strip()],
                    statuses=[item.strip() for item in args.statuses.split(",") if item.strip()],
                    storage_state=args.storage_state,
                    channel=args.channel,
                )
                print(f"WMS 查询总数: {totals['total']} ({json.dumps(totals['by_warehouse'], ensure_ascii=False)})", flush=True)
            except Exception as exc:
                print(f"WMS 总数查询失败，继续执行下载: {exc}", flush=True)

            emit(callback, "progress", progress=8)
            print("开始下载面单", flush=True)
            from task_pipeline import tracked_download_pdf, tracked_extract_data, tracked_run_ocr
            from pipeline import read_json_file

            download_manifest = tracked_download_pdf(args)
            emit(callback, "progress", progress=65)
            print("开始解析 PDF / 图片并校验条码", flush=True)
            ocr_results = tracked_run_ocr(args, download_manifest)
            emit(callback, "progress", progress=88)
            print("开始生成 Excel", flush=True)
            extract_manifest = tracked_extract_data(args, ocr_results)
            payload = read_json_file(extract_manifest)
            excel_path = str(payload.get("excel_path") or "")
            json_path = str(payload.get("json_path") or "")
            result_count = int(payload.get("result_count") or 0)
            emit(callback, "progress", progress=100)
            print(f"任务完成: {excel_path}", flush=True)

    result = {
        "run_name": run_name,
        "excel_path": excel_path,
        "json_path": json_path,
        "log_path": str(log_path),
        "metadata_path": str(metadata_path),
        "result_count": result_count,
    }
    emit(callback, "done", **result)
    return result
