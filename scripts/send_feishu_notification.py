from __future__ import annotations

import argparse
import base64
import hashlib
import hmac
import json
import os
import time
from pathlib import Path
from typing import Any

import requests


BASE_DIR = Path(__file__).resolve().parents[1]


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


def feishu_sign(secret: str, timestamp: str) -> str:
    string_to_sign = f"{timestamp}\n{secret}".encode("utf-8")
    digest = hmac.new(string_to_sign, b"", hashlib.sha256).digest()
    return base64.b64encode(digest).decode("utf-8")


def build_card(title: str, status: str, fields: list[tuple[str, str]], note: str = "") -> dict[str, Any]:
    color = "green"
    if status == "warning":
        color = "yellow"
    elif status == "error":
        color = "red"

    elements: list[dict[str, Any]] = []
    if fields:
        elements.append(
            {
                "tag": "div",
                "fields": [
                    {
                        "is_short": True,
                        "text": {"tag": "lark_md", "content": f"**{name}**\n{value}"},
                    }
                    for name, value in fields
                ],
            }
        )
    if note:
        elements.append({"tag": "hr"})
        elements.append({"tag": "div", "text": {"tag": "lark_md", "content": note}})

    return {
        "msg_type": "interactive",
        "card": {
            "config": {"wide_screen_mode": True},
            "header": {
                "template": color,
                "title": {"tag": "plain_text", "content": title},
            },
            "elements": elements,
        },
    }


def send_message(payload: dict[str, Any], webhook_url: str, secret: str = "") -> dict[str, Any]:
    if secret:
        timestamp = str(int(time.time()))
        payload["timestamp"] = timestamp
        payload["sign"] = feishu_sign(secret, timestamp)

    response = requests.post(webhook_url, json=payload, timeout=20)
    response.raise_for_status()
    data = response.json()
    if data.get("code") not in (0, None):
        raise RuntimeError(f"Feishu webhook error: {data}")
    return data


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Send PDF_DDD notification to Feishu webhook.")
    parser.add_argument("--title", default="PDF_DDD 测试通知")
    parser.add_argument("--status", choices=["success", "warning", "error"], default="success")
    parser.add_argument("--note", default="")
    parser.add_argument("--field", action="append", default=[], help="Field in name=value format; can repeat.")
    return parser.parse_args()


def main() -> int:
    load_local_env(BASE_DIR / ".env")
    args = parse_args()
    webhook_url = os.environ.get("FEISHU_WEBHOOK_URL", "").strip()
    secret = os.environ.get("FEISHU_WEBHOOK_SECRET", "").strip()
    if not webhook_url:
        raise RuntimeError("FEISHU_WEBHOOK_URL is missing. Add it to .env first.")

    fields: list[tuple[str, str]] = []
    for item in args.field:
        if "=" in item:
            name, value = item.split("=", 1)
            fields.append((name.strip(), value.strip()))

    payload = build_card(args.title, args.status, fields, args.note)
    result = send_message(payload, webhook_url, secret)
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
