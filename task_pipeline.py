from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, TypeVar

from pipeline import download_pdf, extract_data, read_json_file, run_ocr, write_json_file
from task_state import append_task_state, make_task_id, task_map_from_manifest, task_state_path
from utils import log


T = TypeVar("T")
RETRY_LIMIT = 3


def retry_step(stage: str, action: Callable[[], T]) -> T:
    last_error: Exception | None = None
    for attempt in range(1, RETRY_LIMIT + 1):
        try:
            if attempt > 1:
                log(f"{stage} retry {attempt}/{RETRY_LIMIT}")
            return action()
        except Exception as exc:
            last_error = exc
            log(f"{stage} failed attempt {attempt}/{RETRY_LIMIT}: {exc}")
    assert last_error is not None
    raise last_error


def normalize_file_path(path_text: str) -> str:
    return str(Path(path_text).expanduser().resolve())


def attach_tasks_to_manifest(manifest_path: Path, state_path: Path) -> list[dict[str, str]]:
    manifest = read_json_file(manifest_path)
    tasks: list[dict[str, str]] = []
    for file_path in manifest.get("files", []):
        normalized_path = normalize_file_path(str(file_path))
        task_id = make_task_id(normalized_path)
        tasks.append({"task_id": task_id, "file_path": normalized_path})
        append_task_state(
            state_path,
            task_id=task_id,
            file_path=normalized_path,
            stage="DOWNLOAD",
            status="CREATED",
        )
        append_task_state(
            state_path,
            task_id=task_id,
            file_path=normalized_path,
            stage="DOWNLOAD",
            status="DOWNLOADED",
        )

    manifest["tasks"] = tasks
    manifest["task_state"] = str(state_path)
    write_json_file(manifest_path, manifest)
    return tasks


def tracked_download_pdf(args) -> Path:
    state_path = task_state_path(args)
    try:
        manifest_path = retry_step("DOWNLOAD", lambda: download_pdf(args))
        attach_tasks_to_manifest(manifest_path, state_path)
        return manifest_path
    except Exception as exc:
        task_id = make_task_id(f"{getattr(args, 'output_name', 'download')}:download")
        append_task_state(
            state_path,
            task_id=task_id,
            file_path="",
            stage="DOWNLOAD",
            status="FAILED",
            error_type="DOWNLOAD_FAILED",
            error_msg=str(exc),
        )
        raise


def classify_ocr_failure(result: dict[str, Any]) -> tuple[str, str]:
    reason = " ".join(
        str(result.get(key) or "")
        for key in ("reason", "reason_zh", "barcode_error", "barcode_note", "ocr_reason")
    ).lower()
    suffix = Path(str(result.get("file_path") or result.get("file_name") or "")).suffix.lower()
    if "non-pdf" in reason or "unsupported file format" in reason:
        return "NOT_PDF", reason
    if suffix and suffix != ".pdf" and suffix not in {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff"}:
        return "NOT_PDF", reason
    return "OCR_FAILED", reason


def attach_tasks_to_ocr_results(ocr_path: Path, download_manifest_path: Path, state_path: Path) -> None:
    manifest = read_json_file(download_manifest_path)
    ocr_payload = read_json_file(ocr_path)
    task_map = task_map_from_manifest(manifest)
    completed_task_ids: set[str] = set()
    ocr_done_task_ids: set[str] = set()

    results = ocr_payload.get("results", [])
    if isinstance(results, list):
        for result in results:
            if not isinstance(result, dict):
                continue
            file_path = normalize_file_path(str(result.get("file_path") or ""))
            task_id = task_map.get(file_path.lower()) or make_task_id(file_path)
            result["task_id"] = task_id
            completed_task_ids.add(task_id)
            verify_status = str(result.get("verify_status") or "")
            if verify_status == "download_file_error":
                error_type, error_msg = classify_ocr_failure(result)
                append_task_state(
                    state_path,
                    task_id=task_id,
                    file_path=file_path,
                    stage="OCR",
                    status="FAILED",
                    error_type=error_type,
                    error_msg=error_msg,
                )
            else:
                append_task_state(
                    state_path,
                    task_id=task_id,
                    file_path=file_path,
                    stage="OCR",
                    status="OCR_DONE",
                )
                ocr_done_task_ids.add(task_id)

    ocr_payload["tasks"] = manifest.get("tasks", [])
    ocr_payload["processed_task_ids"] = sorted(completed_task_ids)
    ocr_payload["ocr_done_task_ids"] = sorted(ocr_done_task_ids)
    ocr_payload["task_state"] = str(state_path)
    write_json_file(ocr_path, ocr_payload)


def tracked_run_ocr(args, download_manifest_path: Path) -> Path:
    state_path = task_state_path(args)
    try:
        ocr_path = retry_step("OCR", lambda: run_ocr(args, download_manifest_path))
        attach_tasks_to_ocr_results(ocr_path, download_manifest_path, state_path)
        return ocr_path
    except Exception as exc:
        manifest = read_json_file(download_manifest_path)
        for task in manifest.get("tasks", []):
            if not isinstance(task, dict):
                continue
            append_task_state(
                state_path,
                task_id=str(task.get("task_id") or ""),
                file_path=str(task.get("file_path") or ""),
                stage="OCR",
                status="FAILED",
                error_type="OCR_FAILED",
                error_msg=str(exc),
            )
        raise


def tracked_extract_data(args, ocr_results_path: Path) -> Path:
    state_path = task_state_path(args)
    ocr_payload = read_json_file(ocr_results_path)
    try:
        extract_path = extract_data(args, ocr_results_path)
        extract_payload = read_json_file(extract_path)
        ocr_done_task_ids = set(ocr_payload.get("ocr_done_task_ids", []))
        task_by_id = {
            str(task.get("task_id") or ""): task
            for task in ocr_payload.get("tasks", [])
            if isinstance(task, dict)
        }
        for task_id in sorted(ocr_done_task_ids):
            task = task_by_id.get(task_id, {})
            append_task_state(
                state_path,
                task_id=task_id,
                file_path=str(task.get("file_path") or ""),
                stage="EXTRACT",
                status="EXTRACT_DONE",
            )
        extract_payload["tasks"] = ocr_payload.get("tasks", [])
        extract_payload["task_state"] = str(state_path)
        write_json_file(extract_path, extract_payload)
        return extract_path
    except Exception as exc:
        for task in ocr_payload.get("tasks", []):
            if not isinstance(task, dict):
                continue
            append_task_state(
                state_path,
                task_id=str(task.get("task_id") or ""),
                file_path=str(task.get("file_path") or ""),
                stage="EXTRACT",
                status="FAILED",
                error_type="PARSE_FAILED",
                error_msg=str(exc),
            )
        raise
