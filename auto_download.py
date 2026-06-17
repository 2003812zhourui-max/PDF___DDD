"""
WMS 面单自动下载 —— 复用 batch_download_wms_pdfs 的完整 HTTP 通道，支持自动登录 + 并发。

用法：
  set WMS_USERNAME=your_user && set WMS_PASSWORD=your_password
  python auto_download.py --wh-codes US02
  python auto_download.py --username your_user --password your_password --wh-codes US02 --date 2026-06-08 --workers 8
"""

import argparse
import base64
import csv
import hashlib
import json
import os
import sys
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import unquote

import requests
from curl_cffi import requests as curl_requests

from track_key import generate_track_key

# 直接复用原始脚本的 HTTP 通道（已验证可用）
from batch_download_wms_pdfs import (
    session_from_storage_state as _build_session,
    fetch_json_http_with_auth_retry as _fetch_json,
    fetch_json_http as _fetch_json_direct,
    download_pdf_http as _download_pdf,
    append_label_metadata,
)

# ============================
# 常量
# ============================

BASE_URL = "https://omp.xlwms.com"
LOGIN_API = "/gateway/wms/auth/login"
LIST_API = "/gateway/wms/blDelivery/page"
DETAIL_API = "/gateway/wms/blDelivery/detail"
DOWNLOAD_URL_API = "/gateway/wms/appendix/getPreviewAndDownLoadUrl"

BASE_DIR = Path(__file__).resolve().parent
PDF_DIR = BASE_DIR / "pdf_downloads"
LOG_DIR = BASE_DIR / "logs"
LOG_FILE = LOG_DIR / "download_log.csv"
AUTH_STATE_FILE = BASE_DIR / "wms_storage_state.json"

CSV_FIELDS = [
    "deliveryNo", "sourceNo", "expressNo", "customerCode",
    "whCode", "status", "filePath", "error", "downloadedAt",
]
LABEL_SUFFIXES = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"}


# ============================
# 工具函数
# ============================

_log_lock = threading.Lock()
_metadata_lock = threading.Lock()
_success_lock = threading.Lock()
_thread_local = threading.local()
_counter = {"done": 0, "total": 0}


def log(msg: str) -> None:
    with _log_lock:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def progress_log(dno: str, status: str, detail: str = "") -> None:
    with _log_lock:
        _counter["done"] += 1
        suffix = f" ({detail})" if detail else ""
        print(f"[{_counter['done']}/{_counter['total']}] {status} {dno}{suffix}", flush=True)


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def sanitize(value: Any) -> str:
    import re
    text = str(value or "").strip()
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    text = re.sub(r"\s+", "_", text)
    return text or "EMPTY"


def first_non_empty(record: dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = record.get(key)
        if value is not None and value != "":
            return str(value)
    return ""


def compact_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def retry_delay(attempt: int, base_delay: float) -> float:
    # 优化备注：指数退避加轻微抖动，避免 WMS 短时限流/断连后所有线程同时重试。
    jitter = (uuid.uuid4().int % 100) / 1000
    return min(base_delay * (2 ** max(attempt - 1, 0)) + jitter, 8.0)


def is_token_error(error_text: str) -> bool:
    text = str(error_text or "").lower()
    return "401" in text or "unauthorized" in text or "请求未授权" in text or ("token" in text and "expired" in text)


# ============================
# 登录 & Token 管理
# ============================

def _extract_tenant_code(token: str) -> str:
    """从 JWT token 中提取 tenantCode"""
    try:
        p = token.split(".")[1] + "=="
        decoded = json.loads(base64.b64decode(p))
        sub = json.loads(unquote(decoded.get("sub", "{}")))
        return sub.get("tenantCode", "")
    except Exception:
        return ""


def _token_from_storage_state() -> tuple[str, str]:
    """从 wms_storage_state.json 中提取已有的 (token, tenant_code)。
    如果文件不存在或没有 token，返回 ("", "")。
    """
    if not AUTH_STATE_FILE.exists():
        return "", ""
    try:
        state = json.loads(AUTH_STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return "", ""

    origins = state.get("origins", [])
    for origin_entry in origins:
        if not isinstance(origin_entry, dict):
            continue
        if origin_entry.get("origin") != BASE_URL:
            continue
        entries = origin_entry.get("localStorage", [])
        if not isinstance(entries, list):
            continue
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            name = entry.get("name", "")
            if name in ("wms-token", "omp-token"):
                token = entry.get("value", "")
                if token:
                    return token, _extract_tenant_code(token)
    return "", ""


def _test_token_valid(wh_code: str = "US02") -> bool:
    """用轻量级 API 调用测试已有 token 是否仍然有效。
    复用 _build_session 构建完整 session（含 cookie、tenantcode 等），避免漏 header。
    返回 True 表示 token 可用，False 表示已过期。
    """
    try:
        session, auth_values = _build_session(str(AUTH_STATE_FILE), wh_code, "auto")
        # 用一个极小的分页请求来验证 token
        payload = {
            "current": 1, "size": 1, "status": "15",
            "startTime": "2026-01-01 00:00:00", "endTime": "2026-01-01 00:00:01",
            "whCode": wh_code,
        }
        result = _fetch_json(session, auth_values, LIST_API, method="POST", payload=payload)
        return result.get("code") == 200
    except Exception:
        return False


def try_existing_token(wh_code: str = "US02") -> tuple[str, str] | None:
    """尝试复用 wms_storage_state.json 中已有的 token。
    如果 token 有效返回 (token, tenant_code)，否则返回 None。
    """
    token, tenant_code = _token_from_storage_state()
    if not token:
        return None

    log("检测到已有 token，验证有效性...")
    if _test_token_valid(wh_code):
        log(f"已有 token 有效，跳过登录 (tenantCode={tenant_code})")
        return token, tenant_code

    log("已有 token 已过期")
    return None


def wms_login(username: str, password: str) -> tuple[str, str]:
    """HTTP 登录 WMS，返回 (token, tenant_code)"""
    log(f"登录 WMS: {username}")
    resp = curl_requests.post(
        f"{BASE_URL}{LOGIN_API}",
        json={
            "businessType": "wms",
            "deviceFingerprint": hashlib.md5(uuid.uuid4().bytes).hexdigest()[:8],
            "deviceInfo": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36 | 1920x1080 | Asia/Shanghai",
            "loginAccount": username,
            "loginFlowId": hashlib.md5(uuid.uuid4().bytes).hexdigest(),
            "password": password,
        },
        timeout=30,
        headers={"Content-Type": "application/json;charset=UTF-8", "User-Agent": "Mozilla/5.0"},
    )
    if resp.status_code != 200:
        raise RuntimeError(f"登录失败: HTTP {resp.status_code}")

    data = resp.json()
    if data.get("code") != 200:
        raise RuntimeError(f"登录失败: {json.dumps(data, ensure_ascii=False)[:500]}")

    token = data.get("data", {}).get("token", "")
    if not token:
        raise RuntimeError("登录响应没有 token")

    tenant_code = _extract_tenant_code(token)
    log(f"登录成功! tenantCode={tenant_code}")
    return token, tenant_code


def update_storage_state(token: str, tenant_code: str, username: str) -> None:
    """更新 wms_storage_state.json，让原始脚本也能用"""
    state = {}
    if AUTH_STATE_FILE.exists():
        try:
            state = json.loads(AUTH_STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass

    origin = "https://omp.xlwms.com"
    origins = state.get("origins", [])
    found = False
    for o in origins:
        if o.get("origin") == origin:
            found = True
            ls = o.get("localStorage", [])
            names = {e.get("name"): i for i, e in enumerate(ls)}
            for name, value in [("wms-token", token), ("omp-token", token)]:
                if name in names:
                    ls[names[name]]["value"] = value
                else:
                    ls.append({"name": name, "value": value})
            if "wh" not in names:
                ls.append({"name": "wh", "value": json.dumps({"whCode": "US02", "tenantCode": tenant_code})})
            if "language" not in names:
                ls.append({"name": "language", "value": "zh"})
            o["localStorage"] = ls
            break
    if not found:
        origins.append({
            "origin": origin,
            "localStorage": [
                {"name": "wms-token", "value": token},
                {"name": "omp-token", "value": token},
                {"name": "wh", "value": json.dumps({"whCode": "US02", "tenantCode": tenant_code})},
                {"name": "language", "value": "zh"},
            ],
        })
    state["origins"] = origins

    # 确保有基本 cookies
    cookies = state.get("cookies", [])
    cookie_names = {c.get("name") for c in cookies}
    for name, value in [("version", "prod"), ("prod", "always")]:
        if name not in cookie_names:
            cookies.append({"name": name, "value": value, "domain": ".xlwms.com", "path": "/"})
    state["cookies"] = cookies

    AUTH_STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"已更新 {AUTH_STATE_FILE.name}")


# ============================
# WMS HTTP 客户端（复用原始脚本的认证通道）
# ============================

# HTTP 请求和下载直接使用原始脚本的函数（已导入）


# ============================
# 业务逻辑
# ============================

def extract_label_files(detail: dict[str, Any]) -> list[dict[str, str]]:
    labels = []
    fk = first_non_empty(detail, "fileKey")
    fn = first_non_empty(detail, "fileName")
    if fk:
        labels.append({"fileKey": fk, "fileName": fn})
        return labels
    pkg_list = detail.get("packageList")
    if not isinstance(pkg_list, list):
        return labels
    seen = set()
    for pkg in pkg_list:
        if not isinstance(pkg, dict):
            continue
        fk = first_non_empty(pkg, "fileKey")
        fn = first_non_empty(pkg, "fileName")
        if not fk or (fk, fn) in seen:
            continue
        seen.add((fk, fn))
        labels.append({"fileKey": fk, "fileName": fn})
    return labels


def normalize_order(record: dict[str, Any]) -> dict[str, str]:
    return {
        "deliveryNo": first_non_empty(record, "deliveryNo"),
        "sourceNo": first_non_empty(record, "sourceNo"),
        "customerCode": first_non_empty(record, "customerCode"),
        "whCode": first_non_empty(record, "whCode"),
        "expressNo": first_non_empty(record, "expressNo"),
    }


def label_suffix(file_name: str) -> str:
    suffix = Path(str(file_name or "")).suffix.lower()
    return suffix if suffix in LABEL_SUFFIXES else ".pdf"


def make_label_path(order: dict[str, str], count: int, idx: int, file_name: str = "") -> Path:
    dno = sanitize(order["deliveryNo"])
    eno = sanitize(order["expressNo"])
    suffix = f"_{idx}" if count > 1 else ""
    return PDF_DIR / f"{dno}_{eno}{suffix}{label_suffix(file_name)}"


def make_pdf_path(order: dict[str, str], count: int, idx: int) -> Path:
    return make_label_path(order, count, idx)


def load_success_set() -> set[str]:
    if not LOG_FILE.exists():
        return set()
    result = set()
    with LOG_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if row.get("status") == "success" and row.get("deliveryNo"):
                result.add(row["deliveryNo"])
    return result


_csv_lock = threading.Lock()


def append_log(row: dict[str, Any]) -> None:
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with _csv_lock:
        exists = LOG_FILE.exists()
        with LOG_FILE.open("a", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
            if not exists:
                writer.writeheader()
            writer.writerow({field: row.get(field, "") for field in CSV_FIELDS})


def append_metadata_safely(
    pdf_path: Path,
    order: dict[str, str],
    detail: dict[str, Any],
    label: dict[str, Any],
    url_json: dict[str, Any],
    download_url: str,
) -> None:
    try:
        with _metadata_lock:
            append_label_metadata(pdf_path, order, detail, label, url_json, download_url)
    except Exception as exc:
        log(f"metadata 写入失败，继续下载流程: {exc}")


def process_one_order(session, auth_values, order: dict[str, str]) -> list[Path]:
    from urllib.parse import urlencode
    query = urlencode({"deliveryNo": order["deliveryNo"], "customerCode": order["customerCode"],
                       "whCode": order["whCode"]})
    detail_json = request_json(session, auth_values, f"{DETAIL_API}?{query}", method="GET")
    detail = detail_json.get("data")
    if not isinstance(detail, dict):
        raise RuntimeError("详情接口没有返回 data")

    detail_order = {
        "deliveryNo": first_non_empty(detail, "deliveryNo") or order["deliveryNo"],
        "sourceNo": order["sourceNo"],
        "customerCode": first_non_empty(detail, "customerCode") or order["customerCode"],
        "whCode": first_non_empty(detail, "whCode") or order["whCode"],
        "expressNo": first_non_empty(detail, "expressNo") or order["expressNo"],
    }

    labels = extract_label_files(detail)
    if not labels:
        raise RuntimeError("详情接口未找到 fileKey")

    saved = []
    for idx, label in enumerate(labels, start=1):
        q = urlencode({"fileKey": label["fileKey"], "fileName": label["fileName"] or f"{detail_order['expressNo']}.pdf",
                       "customerCode": detail_order["customerCode"], "whCode": detail_order["whCode"]})
        url_json = request_json(session, auth_values, f"{DOWNLOAD_URL_API}?{q}", method="GET")
        dl_url = url_json.get("data", {}).get("downLoadUrl") if isinstance(url_json.get("data"), dict) else ""
        if not dl_url:
            raise RuntimeError("下载链接接口没有返回 downLoadUrl")
        pdf_path = make_label_path(detail_order, len(labels), idx, str(label.get("fileName") or ""))
        download_pdf(session, dl_url, pdf_path)
        append_metadata_safely(pdf_path, detail_order, detail, label, url_json, dl_url)
        saved.append(pdf_path)
    return saved


def download_task(
    session,
    auth_values,
    order: dict[str, str],
    success_set: set[str],
    force: bool,
    max_attempts: int = 3,
    retry_base_delay: float = 1.0,
) -> dict:
    dno = order["deliveryNo"]
    with _success_lock:
        already_success = dno in success_set
    if not force and already_success:
        return {"deliveryNo": dno, "status": "skipped"}

    last_error = ""
    for attempt in range(1, max_attempts + 1):
        try:
            paths = process_one_order(session, auth_values, order)
            fp = "|".join(str(p.resolve()) for p in paths)
            append_log({**order, "status": "success", "filePath": fp, "error": "", "downloadedAt": now_text()})
            with _success_lock:
                success_set.add(dno)
            return {"deliveryNo": dno, "status": "success", "filePath": fp}
        except Exception as exc:
            last_error = str(exc)
            if is_token_error(last_error):
                return {"deliveryNo": dno, "status": "token_expired", "error": last_error}
            if attempt < max_attempts:
                time.sleep(retry_delay(attempt, retry_base_delay))

    append_log({**order, "status": "failed", "filePath": "", "error": last_error, "downloadedAt": now_text()})
    return {"deliveryNo": dno, "status": "failed", "error": last_error}


# ============================
# 主流程
# ============================

def yesterday_range() -> tuple[str, str]:
    d = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    return f"{d} 00:00:00", f"{d} 23:59:59"


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="WMS 面单自动下载 (HTTP + 并发)")
    p.add_argument("--username", default=os.environ.get("WMS_USERNAME", ""))
    p.add_argument("--password", default=os.environ.get("WMS_PASSWORD", ""))
    p.add_argument("--wh-codes", default="US02")
    p.add_argument("--date", default="")
    p.add_argument("--start-time", default="")
    p.add_argument("--end-time", default="")
    p.add_argument("--status", default="15")
    p.add_argument("--size", type=int, default=100)
    p.add_argument("--max-pages", type=int, default=0, help="最大页数；0 表示不限，直到接口没有更多记录")
    p.add_argument("--limit", type=int, default=0, help="最多下载订单数；0 表示不限")
    p.add_argument("--workers", type=int, default=8, help="下载并发线程数；服务器上可按带宽/API限流调整")
    p.add_argument("--download-retries", type=int, default=5, help="单订单下载最大尝试次数，默认 5")
    p.add_argument("--retry-base-delay", type=float, default=0.8, help="下载/列表重试基础等待秒数，默认 0.8")
    p.add_argument("--pdf-dir", default=str(PDF_DIR))
    p.add_argument("--log-file", default=str(LOG_FILE))
    p.add_argument("--channel", default="", help="物流渠道筛选，例如 TikTok-CBT-US、Upload_Shipping_Label-Speedx")
    p.add_argument("--force", action="store_true")
    return p.parse_args()


def process_one(session, auth_values, order: dict[str, str]) -> list[Path]:
    """用原始脚本的函数处理单个订单"""
    from urllib.parse import urlencode
    q = urlencode({"deliveryNo": order["deliveryNo"], "customerCode": order["customerCode"],
                   "whCode": order["whCode"]})
    detail_json = _fetch_json(session, auth_values, f"{DETAIL_API}?{q}", method="GET")
    detail = detail_json.get("data")
    if not isinstance(detail, dict):
        raise RuntimeError("详情接口没有返回 data")

    d_order = {
        "deliveryNo": first_non_empty(detail, "deliveryNo") or order["deliveryNo"],
        "sourceNo": order["sourceNo"],
        "customerCode": first_non_empty(detail, "customerCode") or order["customerCode"],
        "whCode": first_non_empty(detail, "whCode") or order["whCode"],
        "expressNo": first_non_empty(detail, "expressNo") or order["expressNo"],
    }

    labels = extract_label_files(detail)
    if not labels:
        raise RuntimeError("详情接口未找到 fileKey")

    saved = []
    for idx, label in enumerate(labels, start=1):
        q2 = urlencode({"fileKey": label["fileKey"], "fileName": label["fileName"] or f"{d_order['expressNo']}.pdf",
                        "customerCode": d_order["customerCode"], "whCode": d_order["whCode"]})
        url_json = _fetch_json(session, auth_values, f"{DOWNLOAD_URL_API}?{q2}", method="GET")
        dl_url = url_json.get("data", {}).get("downLoadUrl") if isinstance(url_json.get("data"), dict) else ""
        if not dl_url:
            raise RuntimeError("下载链接接口没有返回 downLoadUrl")
        pdf_path = make_label_path(d_order, len(labels), idx, str(label.get("fileName") or ""))
        _download_pdf(session, dl_url, pdf_path)
        append_metadata_safely(pdf_path, d_order, detail, label, url_json, dl_url)
        saved.append(pdf_path)
    return saved


def download_task(
    session,
    auth_values,
    order: dict[str, str],
    success_set: set[str],
    force: bool,
    max_attempts: int = 3,
    retry_base_delay: float = 1.0,
) -> dict:
    dno = order["deliveryNo"]
    with _success_lock:
        already_success = dno in success_set
    if not force and already_success:
        return {"deliveryNo": dno, "status": "skipped"}

    last_error = ""
    for attempt in range(1, max_attempts + 1):
        try:
            paths = process_one(session, auth_values, order)
            fp = "|".join(str(p.resolve()) for p in paths)
            append_log({**order, "status": "success", "filePath": fp, "error": "", "downloadedAt": now_text()})
            with _success_lock:
                success_set.add(dno)
            return {"deliveryNo": dno, "status": "success", "filePath": fp}
        except Exception as exc:
            last_error = str(exc)
            if is_token_error(last_error):
                return {"deliveryNo": dno, "status": "token_expired", "error": last_error}
            if attempt < max_attempts:
                time.sleep(retry_delay(attempt, retry_base_delay))

    append_log({**order, "status": "failed", "filePath": "", "error": last_error, "downloadedAt": now_text()})
    return {"deliveryNo": dno, "status": "failed", "error": last_error}


def _refresh_session(args, wh_code: str) -> tuple[Any, list[str]]:
    """重新登录并重建 session，返回 (session, auth_values)"""
    existing = try_existing_token(wh_code)
    if existing:
        token, tenant_code = existing
    else:
        if not args.password:
            raise RuntimeError("token 已过期且未提供密码，无法刷新登录态")
        token, tenant_code = wms_login(args.username, args.password)
        update_storage_state(token, tenant_code, args.username)
    session, auth_values = _build_session(str(AUTH_STATE_FILE), wh_code, "auto")
    log("登录态已刷新")
    return session, auth_values


def clear_thread_client() -> None:
    if hasattr(_thread_local, "clients"):
        _thread_local.clients = {}


def get_thread_client(wh_code: str, version: int = 0) -> tuple[Any, list[str]]:
    # 优化备注：requests.Session 不是为多线程共享设计的；按线程缓存 session，减少连接争用并保留连接复用速度。
    clients = getattr(_thread_local, "clients", None)
    client_version = getattr(_thread_local, "client_version", None)
    if client_version != version:
        clients = {}
        _thread_local.clients = clients
        _thread_local.client_version = version
    if clients is None:
        clients = {}
        _thread_local.clients = clients
        _thread_local.client_version = version
    if wh_code not in clients:
        clients[wh_code] = _build_session(str(AUTH_STATE_FILE), wh_code, "auto")
    session, auth_values = clients[wh_code]
    session.headers["whcode"] = wh_code
    return session, auth_values


def fetch_order_page_with_retry(
    session,
    auth_values,
    payload: dict[str, Any],
    max_attempts: int,
    retry_base_delay: float,
) -> dict[str, Any]:
    # 优化备注：列表页偶发 ConnectionReset 时只重试当前页，避免整批任务从头失败。
    last_error = ""
    for attempt in range(1, max_attempts + 1):
        try:
            return _fetch_json(session, auth_values, LIST_API, method="POST", payload=payload)
        except Exception as exc:
            last_error = str(exc)
            if is_token_error(last_error) or attempt >= max_attempts:
                raise
            log(f"列表接口失败，重试 {attempt + 1}/{max_attempts}: {last_error[:120]}")
            time.sleep(retry_delay(attempt, retry_base_delay))
    raise RuntimeError(last_error)


def write_download_summary(
    args,
    start_time: str,
    end_time: str,
    stats: dict[str, int],
    elapsed: float,
    total_orders: int,
) -> Path:
    # 优化备注：summary JSON 是给飞书/服务器消费的稳定接口，不依赖解析控制台日志。
    summary_path = LOG_FILE.with_name(f"{LOG_FILE.stem}_summary.json")
    payload = {
        "start_time": start_time,
        "end_time": end_time,
        "wh_codes": [c.strip() for c in args.wh_codes.split(",") if c.strip()],
        "statuses": [s.strip() for s in args.status.split(",") if s.strip()],
        "channel": args.channel,
        "workers": args.workers,
        "download_retries": args.download_retries,
        "total_orders": total_orders,
        "stats": stats,
        "elapsed_seconds": round(elapsed, 3),
        "pdf_dir": str(PDF_DIR.resolve()),
        "download_log": str(LOG_FILE.resolve()),
        "written_at": now_text(),
    }
    summary_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary_path


def main() -> None:
    global PDF_DIR, LOG_FILE, _counter

    args = parse_args()

    PDF_DIR = Path(args.pdf_dir)
    LOG_FILE = Path(args.log_file)
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

    if args.start_time and args.end_time:
        start_time, end_time = args.start_time, args.end_time
    elif args.date:
        start_time, end_time = f"{args.date} 00:00:00", f"{args.date} 23:59:59"
    else:
        start_time, end_time = yesterday_range()

    wh_codes = [c.strip() for c in args.wh_codes.split(",") if c.strip()]
    status_values = [s.strip() for s in args.status.split(",") if s.strip()]
    if not status_values:
        status_values = ["15"]
    success_set = load_success_set()

    log(f"时间: {start_time} ~ {end_time}")
    log(f"仓库: {wh_codes}")
    log(f"状态: {status_values}")
    log(f"并发: {args.workers}")
    log(f"下载重试: {args.download_retries} 次，基础退避: {args.retry_base_delay}s")
    if args.channel:
        log(f"物流渠道: {args.channel}")

    # 优先复用已有 token，避免不必要的登录请求
    wh_code_for_test = wh_codes[0] if wh_codes else "US02"
    existing = try_existing_token(wh_code_for_test)
    if existing:
        token, tenant_code = existing
    else:
        if not args.password:
            raise SystemExit("已有 token 已过期且未提供密码，请设置 --password 或 WMS_PASSWORD 环境变量")
        token, tenant_code = wms_login(args.username, args.password)
        update_storage_state(token, tenant_code, args.username)

    # 用原始脚本的函数构建 session（已验证可用）
    session, auth_values = _build_session(str(AUTH_STATE_FILE), wh_codes[0], "auto")

    # 收集订单（用原始脚本的 fetch_json）
    all_orders: list[tuple[str, dict[str, str]]] = []
    for wh_code in wh_codes:
        session.headers["whcode"] = wh_code
        for status in status_values:
            log(f"获取 {wh_code} 状态 {status} 订单...")
            page = 1
            while args.max_pages <= 0 or page <= args.max_pages:
                payload = {
                    "appendixFlag": "", "areaCodes": [], "categoryIdList": [], "cellNos": [],
                    "codeType": "barcode", "countKind": "orderWeight", "countryRegionCodes": "",
                    "current": page, "customerCodes": "", "endTime": end_time,
                    "expressFlag": "", "expressPrintStatus": "", "forecastStatus": "",
                    "logisticsCarrier": "", "logisticsChannel": args.channel, "orderCount": "",
                    "orderNoType": "sourceNo", "orderSourceList": [], "productPackType": "",
                    "receiver": "", "relatedReturnOrder": "", "salesPlatform": "",
                    "size": args.size, "skuQtyStrList": [], "sourceNoLists": [],
                    "startTime": start_time, "status": status, "timeType": "createTime",
                    "unitMark": 0, "varietyType": "", "weightCountEnd": "",
                    "weightCountStart": "", "whCode": wh_code, "withVas": "",
                }
                data = fetch_order_page_with_retry(
                    session,
                    auth_values,
                    payload,
                    max_attempts=max(1, args.download_retries),
                    retry_base_delay=max(0.1, args.retry_base_delay),
                )
                records = data.get("data", {}).get("records", []) if isinstance(data.get("data"), dict) else []
                if not records:
                    break
                for r in records:
                    o = normalize_order(r)
                    if o["deliveryNo"]:
                        all_orders.append((wh_code, o))
                log(f"  status={status} page={page}: {len(records)} 条")
                if len(records) < args.size:
                    break
                page += 1

    # 过滤
    seen = set()
    filtered = []
    for wc, o in all_orders:
        dno = o["deliveryNo"]
        if dno in seen:
            continue
        seen.add(dno)
        if not args.force and dno in success_set:
            continue
        filtered.append((wc, o))

    if not filtered:
        log("没有需要下载的订单")
        return

    if args.limit and args.limit > 0:
        filtered = filtered[:args.limit]
    _counter["total"] = len(filtered)
    _counter["done"] = 0

    log(f"开始下载: {len(filtered)} 单, {args.workers} 线程")
    t0 = time.time()
    stats = {"ok": 0, "fail": 0, "skip": 0, "refresh": 0}

    _refresh_lock = threading.Lock()
    _refresh_state = {"version": 0}
    _max_refresh = 3  # 最多自动刷新 3 次

    def do_task(wc, o):
        for attempt in range(2):  # 最多重试 1 次（刷新后重试）
            task_session, task_auth = get_thread_client(wc, _refresh_state["version"])
            result = download_task(
                task_session,
                task_auth,
                o,
                success_set,
                args.force,
                max_attempts=max(1, args.download_retries),
                retry_base_delay=max(0.1, args.retry_base_delay),
            )
            if result["status"] != "token_expired":
                return result
            # token 过期，尝试刷新
            with _refresh_lock:
                # 双重检查：其他线程可能已经刷新过了
                if stats["refresh"] >= _max_refresh:
                    return result
                stats["refresh"] += 1
                log(f"Token 过期，自动刷新 ({stats['refresh']}/{_max_refresh})...")
                try:
                    _refresh_session(args, wc)
                    _refresh_state["version"] += 1
                    clear_thread_client()
                except Exception as refresh_err:
                    log(f"刷新失败: {refresh_err}")
                    return result
            # 刷新成功，重试当前订单
        return result

    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futs = {pool.submit(do_task, wc, o): o for wc, o in filtered}
        for fut in as_completed(futs):
            o = futs[fut]
            try:
                r = fut.result()
            except Exception as e:
                append_log({**o, "status": "failed", "filePath": "", "error": str(e), "downloadedAt": now_text()})
                r = {"deliveryNo": o["deliveryNo"], "status": "failed", "error": str(e)}
            s = r["status"]
            dno = r["deliveryNo"]
            if s == "token_expired":
                # 刷新次数用尽
                stats["fail"] += 1
                append_log({**o, "status": "failed", "filePath": "", "error": r.get("error", "token expired"), "downloadedAt": now_text()})
                progress_log(dno, "[FAIL]", "token expired, max refresh reached")
            elif s == "success":
                stats["ok"] += 1
                progress_log(dno, "[OK]", r.get("filePath", "").split("\\")[-1])
            elif s == "skipped":
                stats["skip"] += 1
                progress_log(dno, "[SKIP]")
            else:
                stats["fail"] += 1
                progress_log(dno, "[FAIL]", r.get("error", "")[:60])

    elapsed = time.time() - t0
    log(f"完成 ({elapsed:.1f}s) ok={stats['ok']} fail={stats['fail']} skip={stats['skip']} refresh={stats['refresh']}")
    summary_path = write_download_summary(args, start_time, end_time, stats, elapsed, len(filtered))
    log(f"下载汇总 JSON: {summary_path.resolve()}")
    log(f"PDF: {PDF_DIR.resolve()}")
    log(f"Log: {LOG_FILE.resolve()}")


if __name__ == "__main__":
    main()
