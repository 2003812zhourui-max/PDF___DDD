from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Any

import requests


BASE_DIR = Path(__file__).resolve().parents[1]
FEISHU_API = "https://open.feishu.cn/open-apis"


def load_local_env(path: Path) -> None:
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


def feishu_post_json(url: str, payload: dict[str, Any], token: str | None = None) -> dict[str, Any]:
    headers = {"Content-Type": "application/json; charset=utf-8"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    response = requests.post(url, headers=headers, json=payload, timeout=30)
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        raise RuntimeError(f"Feishu HTTP error {response.status_code}: {response.text}") from exc
    data = response.json()
    if data.get("code") != 0:
        raise RuntimeError(f"Feishu API error: {data}")
    return data


def get_tenant_access_token(app_id: str, app_secret: str) -> str:
    data = feishu_post_json(
        f"{FEISHU_API}/auth/v3/tenant_access_token/internal",
        {"app_id": app_id, "app_secret": app_secret},
    )
    token = data.get("tenant_access_token")
    if not token:
        raise RuntimeError(f"tenant_access_token missing: {data}")
    return str(token)


def file_type_for(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".xls", ".xlsx"}:
        return "xls"
    if suffix in {".doc", ".docx"}:
        return "doc"
    if suffix in {".ppt", ".pptx"}:
        return "ppt"
    if suffix == ".pdf":
        return "pdf"
    if suffix == ".csv":
        return "stream"
    return "stream"


def upload_file(path: Path, token: str) -> str:
    url = f"{FEISHU_API}/im/v1/files"
    headers = {"Authorization": f"Bearer {token}"}
    with path.open("rb") as file_obj:
        files = {"file": (path.name, file_obj, "application/octet-stream")}
        data = {"file_type": file_type_for(path), "file_name": path.name}
        response = requests.post(url, headers=headers, data=data, files=files, timeout=120)
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        raise RuntimeError(f"Feishu HTTP error {response.status_code}: {response.text}") from exc
    payload = response.json()
    if payload.get("code") != 0:
        raise RuntimeError(f"Feishu upload error: {payload}")
    file_key = (payload.get("data") or {}).get("file_key")
    if not file_key:
        raise RuntimeError(f"file_key missing: {payload}")
    return str(file_key)


def send_file_message(chat_id: str, file_key: str, token: str) -> dict[str, Any]:
    return feishu_post_json(
        f"{FEISHU_API}/im/v1/messages?receive_id_type=chat_id",
        {
            "receive_id": chat_id,
            "msg_type": "file",
            "content": json.dumps({"file_key": file_key}, ensure_ascii=False),
        },
        token,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Upload a file to Feishu and send it to a chat.")
    parser.add_argument("file", help="Local file path to upload.")
    return parser.parse_args()


def main() -> int:
    load_local_env(BASE_DIR / ".env")
    args = parse_args()
    path = Path(args.file).expanduser().resolve()
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(path)

    app_id = os.environ.get("FEISHU_APP_ID", "").strip()
    app_secret = os.environ.get("FEISHU_APP_SECRET", "").strip()
    chat_id = os.environ.get("FEISHU_CHAT_ID", "").strip()
    missing = [name for name, value in [("FEISHU_APP_ID", app_id), ("FEISHU_APP_SECRET", app_secret), ("FEISHU_CHAT_ID", chat_id)] if not value]
    if missing:
        raise RuntimeError(f"Missing required env vars: {', '.join(missing)}")

    token = get_tenant_access_token(app_id, app_secret)
    file_key = upload_file(path, token)
    result = send_file_message(chat_id, file_key, token)
    print(json.dumps({"file_key": file_key, "send_result": result}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
