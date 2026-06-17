from __future__ import annotations

import json
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from utils import ensure_dir


TASK_NAMESPACE = uuid.UUID("8f4ef30f-1a27-4ab6-83ef-e98f7d28a1a5")


def now_timestamp() -> str:
    return datetime.now().isoformat(timespec="seconds")


def task_state_path(args) -> Path:
    explicit_path = str(getattr(args, "task_state", "") or "").strip()
    if explicit_path:
        return Path(explicit_path).expanduser().resolve()

    output_dir = Path(args.output_dir).expanduser().resolve()
    return output_dir / "_intermediate" / str(args.output_name) / "task_state.jsonl"


def make_task_id(file_path: str) -> str:
    path_text = str(Path(file_path).expanduser().resolve()) if file_path else "download"
    return str(uuid.uuid5(TASK_NAMESPACE, path_text.lower()))


def append_task_state(
    path: Path,
    *,
    task_id: str,
    file_path: str,
    stage: str,
    status: str,
    error_type: str = "",
    error_msg: str = "",
) -> None:
    ensure_dir(path.parent)
    row = {
        "task_id": task_id,
        "file_path": file_path,
        "stage": stage,
        "status": status,
        "error_type": error_type,
        "error_msg": error_msg,
        "timestamp": now_timestamp(),
    }
    with path.open("a", encoding="utf-8", newline="\n") as file:
        file.write(json.dumps(row, ensure_ascii=False) + "\n")


def task_map_from_manifest(manifest: dict[str, Any]) -> dict[str, str]:
    tasks = manifest.get("tasks", [])
    mapping: dict[str, str] = {}
    if isinstance(tasks, list):
        for task in tasks:
            if not isinstance(task, dict):
                continue
            file_path = str(task.get("file_path") or "")
            task_id = str(task.get("task_id") or "")
            if file_path and task_id:
                mapping[str(Path(file_path).expanduser().resolve()).lower()] = task_id
    return mapping
