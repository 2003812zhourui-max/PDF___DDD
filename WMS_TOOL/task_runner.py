from __future__ import annotations

import threading
import traceback
import uuid
from copy import deepcopy
from datetime import datetime
from typing import Any

from .worker import run_pipeline


_tasks: dict[str, dict[str, Any]] = {}
_lock = threading.Lock()


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _append_log(task: dict[str, Any], message: str) -> None:
    task.setdefault("log", []).append(f"[{_now()}] {message}")
    if len(task["log"]) > 2000:
        task["log"] = task["log"][-2000:]


def _update_task(task_id: str, **updates: Any) -> None:
    with _lock:
        task = _tasks.get(task_id)
        if not task:
            return
        for key, value in updates.items():
            if key == "log":
                _append_log(task, str(value))
            else:
                task[key] = value
        task["updated_at"] = _now()


def _callback(task_id: str, event: dict[str, Any]) -> None:
    event_name = str(event.get("event") or "")
    if event_name == "log":
        _update_task(task_id, log=str(event.get("message") or ""))
    elif event_name == "progress":
        progress = max(0, min(100, int(event.get("progress") or 0)))
        _update_task(task_id, progress=progress)
    elif event_name == "status":
        _update_task(task_id, status=str(event.get("status") or "running"))
    elif event_name == "done":
        _update_task(
            task_id,
            status="done",
            progress=100,
            result={key: value for key, value in event.items() if key != "event"},
            log=f"Excel 已生成: {event.get('excel_path') or ''}",
        )


def _run_task(task_id: str, params: dict[str, Any]) -> None:
    try:
        _update_task(task_id, status="running", progress=0, log="任务启动")
        result = run_pipeline(params, callback=lambda event: _callback(task_id, event))
        _update_task(task_id, status="done", progress=100, result=result)
    except Exception as exc:
        _update_task(
            task_id,
            status="failed",
            error=str(exc),
            log=f"任务失败: {exc}",
        )
        _update_task(task_id, log=traceback.format_exc())


def start_task(params: dict[str, Any]) -> str:
    task_id = str(uuid.uuid4())
    task = {
        "task_id": task_id,
        "progress": 0,
        "status": "running",
        "log": [],
        "result": None,
        "error": "",
        "created_at": _now(),
        "updated_at": _now(),
    }
    with _lock:
        _tasks[task_id] = task
    thread = threading.Thread(target=_run_task, args=(task_id, dict(params)), daemon=True)
    thread.start()
    return task_id


def get_task(task_id: str) -> dict[str, Any] | None:
    with _lock:
        task = _tasks.get(task_id)
        return deepcopy(task) if task else None
