from __future__ import annotations

import json
import re
import unicodedata
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from barcode_verify_tracking import TRACKING_STATUS_STRONG, VerifyResult, match_template_text
from config import BASE_DIR
from utils import ensure_dir, log


METADATA_JSONL = BASE_DIR / "output" / "download_label_metadata.jsonl"
UNKNOWN = "UNKNOWN"
NORMAL_LABEL_TEMPLATE = "普通面单"

REVIEW_STATUSES = {
    "review_conflict",
    "review_unknown",
    "barcode_failed",
    "source_only_low_confidence",
    "ocr_needed",
    "download_file_error",
}

STATUS_ZH = {
    "auto_pass_triple_verified": "自动通过：条码、PDF文字层、下载数据一致",
    "auto_pass_barcode_verified": "自动通过：条码与下载数据一致",
    "auto_pass_text_verified": "自动通过：PDF文字层与下载数据一致",
    "auto_pass_verified_prefix": "可通过：USPS 前缀匹配",
    "source_only_low_confidence": "低置信度：仅下载数据/文件名可用",
    "review_conflict": "需要复核：多来源结果冲突",
    "review_unknown": "需要复核：无法识别",
    "barcode_failed": "条码识别失败",
    "ocr_needed": "需要 OCR 深度识别",
    "download_file_error": "下载文件异常",
}

CONFIDENCE_ZH = {
    "high": "高",
    "medium": "中",
    "low": "低",
    "review": "需复核",
    "高": "高",
    "中": "中",
    "低": "低",
    "需复核": "需复核",
}

BUSINESS_COLUMNS: list[tuple[str, str, str]] = [
    ("订单号", "order_no", "raw"),
    ("追踪号", "tracking_no", "raw"),
    ("文件名", "file_name", "raw"),
    ("文件路径", "file_path", "raw"),
    ("承运商", "carrier", "raw"),
    ("面单类型", "label_template", "raw"),
    ("模板细分", "template_subdivision", "raw"),
    ("识别依据", "recognition_basis", "raw"),
    ("识别置信度", "recognition_confidence", "raw"),
    ("备注", "template_note", "raw"),
    ("是否人工复核", "need_review", "yes_no"),
    ("最终状态", "verify_status_zh", "raw"),
    ("内部状态", "verify_status", "raw"),
    ("文件类型", "file_type", "raw"),
    ("文件格式", "file_format", "raw"),
    ("条码追踪号", "barcode_tracking", "raw"),
    ("PDF文字层追踪号", "pdf_text_tracking", "raw"),
    ("文件名追踪号", "filename_tracking", "raw"),
    ("下载订单号", "download_order_no", "raw"),
    ("下载波次号", "download_wave_no", "raw"),
    ("下载仓库", "download_warehouse", "raw"),
    ("条码是否成功", "barcode_success", "yes_no"),
    ("条码候选", "all_barcode_candidates", "raw"),
    ("条码类型", "decoded_formats", "raw"),
    ("是否需要OCR", "ocr_needed", "yes_no"),
    ("OCR原因", "ocr_reason", "raw"),
    ("OCR优先级", "ocr_priority", "raw"),
    ("耗时秒", "processing_seconds", "raw"),
    ("metadata承运商", "meta_carrier", "raw"),
    ("metadata承运商冲突", "meta_carrier_conflict", "yes_no"),
    ("metadata承运商备注", "meta_carrier_note", "raw"),
    ("metadata渠道", "meta_channel", "raw"),
    ("metadata渠道代码", "meta_channel_code", "raw"),
    ("metadata客户代码", "meta_customer_code", "raw"),
    ("metadata数据来源", "meta_source", "raw"),
]


@dataclass
class TemplateDecision:
    carrier: str = UNKNOWN
    label_template: str = NORMAL_LABEL_TEMPLATE
    template_subdivision: str = ""
    matched_text: str = ""
    matched_rule: str = ""
    source: str = "默认规则"
    confidence: str = "中"
    note: str = "未命中 0024/CBT/CBS，按普通面单处理"
    need_review: bool = False

    @property
    def recognized(self) -> bool:
        return self.label_template in {"0024", "CBT", "CBS"}


def normalize_text(value: object) -> str:
    return str(value or "").strip()


def normalize_for_match(value: object) -> str:
    text = unicodedata.normalize("NFKC", normalize_text(value)).upper()
    return re.sub(r"\s+", " ", text).strip()


def compact_text(value: object) -> str:
    text = normalize_for_match(value)
    return re.sub(r"[^0-9A-Z\u4e00-\u9fff]+", "", text)


def first_non_empty(*values: object) -> str:
    for value in values:
        text = normalize_text(value)
        if text and text.upper() != UNKNOWN:
            return text
    return ""


def confidence_to_zh(value: object, default: str = "低") -> str:
    text = normalize_text(value)
    return CONFIDENCE_ZH.get(text, default)


def status_to_zh(result: VerifyResult) -> str:
    return STATUS_ZH.get(result.verify_status, first_non_empty(result.verify_status_zh, result.verify_status))


def normalize_carrier(value: object) -> str:
    text = normalize_for_match(value)
    if not text or text == UNKNOWN:
        return ""
    rules = [
        ("USPS", ("USPS", "POSTAL")),
        ("FedEx", ("FEDEX", "FDX")),
        ("UPS", ("UPS",)),
        ("GOFO", ("GOFO", "GFUS")),
        ("SwiftX", ("SWIFTX", "SWX")),
        ("UniUni", ("UNIUNI", "UNI UNI", "UUS")),
        ("OnTrac", ("ONTRAC",)),
        ("LaserShip", ("LASERSHIP", "LASER SHIP", "1LS")),
        ("SpeedX", ("SPEEDX", "SPX")),
        ("Yanwen", ("YANWEN", "YW")),
    ]
    for carrier, tokens in rules:
        if any(token in text for token in tokens):
            return carrier
    return normalize_text(value)


def normalize_cbt_sub(raw: str) -> str:
    value = raw.upper().replace("O", "0")
    match = re.search(r"([AB])0?([1-5])", value)
    if not match:
        return ""
    return f"{match.group(1)}0{match.group(2)}"


def template_subdivision(template_code: str, template_sub_code: str, marker: str) -> str:
    if template_code == "0024":
        return template_sub_code if template_sub_code in {"01", "02"} else ""
    return ""


def find_template(text: object, source: str, confidence: str, carrier_hint: str = "") -> TemplateDecision | None:
    raw = normalize_text(text)
    if not raw:
        return None

    match = match_template_text(raw, source, confidence)
    if match.template_code:
        subdivision = template_subdivision(match.template_code, match.template_sub_code, match.template_marker)
        return TemplateDecision(
            carrier=carrier_hint or UNKNOWN,
            label_template=match.template_code,
            template_subdivision=subdivision,
            matched_text=match.template_marker,
            matched_rule=f"{match.template_code}模板规则",
            source=source,
            confidence=confidence,
            note=f"{source} 命中 {match.template_marker}",
            need_review=match.template_code == "0024" and subdivision == UNKNOWN,
        )
    return None

    normalized = normalize_for_match(raw)
    compact = compact_text(raw)

    cbt_match = re.search(r"CBT([AB][0O]?[1-5])", compact)
    if cbt_match:
        sub_code = normalize_cbt_sub(cbt_match.group(1))
        return TemplateDecision(
            carrier=carrier_hint or UNKNOWN,
            label_template="CBT",
            template_subdivision=sub_code or UNKNOWN,
            matched_text=f"CBT {sub_code}".strip(),
            matched_rule="CBT编号规则",
            source=source,
            confidence=confidence,
            note=f"{source} 命中 CBT {sub_code}" if sub_code else f"{source} 命中 CBT",
            need_review=not bool(sub_code),
        )

    if "TTCBT" in compact or re.search(r"\bTT[-_\s]*CBT\b", normalized):
        return TemplateDecision(
            carrier=carrier_hint or UNKNOWN,
            label_template="CBT",
            template_subdivision=UNKNOWN,
            matched_text="TT-CBT",
            matched_rule="TT-CBT规则",
            source=source,
            confidence=confidence,
            note=f"{source} 命中 TT-CBT，但未识别到 CBT 细分编号",
            need_review=True,
        )

    if "LAXCBS" in compact or re.search(r"\bCBS\b", normalized):
        marker = "LAX-CBS" if "LAXCBS" in compact else "CBS"
        return TemplateDecision(
            carrier=carrier_hint or UNKNOWN,
            label_template="CBS",
            template_subdivision=marker,
            matched_text=marker,
            matched_rule="CBS规则",
            source=source,
            confidence=confidence,
            note=f"{source} 命中 {marker}",
            need_review=False,
        )

    code_match = re.search(r"0024(0?[0-9]{1,2})?", compact)
    if code_match:
        raw_sub = code_match.group(1) or ""
        sub_code = ""
        if raw_sub:
            try:
                sub_code = f"{int(raw_sub.replace('O', '0')):02d}"
            except ValueError:
                sub_code = ""
        marker = f"0024 {sub_code}".strip()
        return TemplateDecision(
            carrier=carrier_hint or UNKNOWN,
            label_template="0024",
            template_subdivision=sub_code or UNKNOWN,
            matched_text=marker,
            matched_rule="0024规则",
            source=source,
            confidence=confidence,
            note=f"{source} 命中 {marker}",
            need_review=False,
        )

    return None


def read_metadata_rows(path: Path = METADATA_JSONL) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    rows: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as file:
        for line in file:
            try:
                item = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(item, dict):
                rows.append(item)
    return rows


def metadata_index(rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    index: dict[str, dict[str, Any]] = {}
    for row in rows:
        for key in (row.get("file_path"), row.get("file_name")):
            text = normalize_text(key)
            if text:
                index[text.lower()] = row
        file_path = normalize_text(row.get("file_path"))
        if file_path:
            try:
                index[str(Path(file_path).resolve()).lower()] = row
            except OSError:
                pass
    return index


def metadata_for_result(result: VerifyResult, index: dict[str, dict[str, Any]]) -> dict[str, Any]:
    candidates = [normalize_text(result.file_path), normalize_text(result.file_name)]
    try:
        candidates.append(str(Path(result.file_path).resolve()))
    except OSError:
        pass
    for candidate in candidates:
        if candidate.lower() in index:
            return index[candidate.lower()]
    return {}


def dict_or_empty(value: object) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def flatten_named_values(prefix: str, value: object) -> Iterable[tuple[str, str]]:
    if isinstance(value, dict):
        for key, item in value.items():
            child_prefix = f"{prefix}.{key}" if prefix else str(key)
            yield from flatten_named_values(child_prefix, item)
    elif isinstance(value, list):
        for index, item in enumerate(value):
            yield from flatten_named_values(f"{prefix}[{index}]", item)
    else:
        text = normalize_text(value)
        if text:
            yield prefix, text


def metadata_named_values(metadata: dict[str, Any]) -> list[tuple[str, str]]:
    if not metadata:
        return []
    fields = dict_or_empty(metadata.get("source_fields"))
    raw_item = dict_or_empty(metadata.get("raw_item"))
    priority_keys = [
        "fileName",
        "fileKey",
        "carrier",
        "channelName",
        "logisticsChannelName",
        "deliveryChannelName",
        "serviceName",
        "logisticsChannel",
        "channelGroupCode",
        "channelGroupName",
        "productName",
    ]
    values: list[tuple[str, str]] = []
    for key in priority_keys:
        for container_name, container in (("metadata", metadata), ("metadata.source_fields", fields), ("metadata.raw_item", raw_item)):
            if key in container:
                values.append((f"{container_name}.{key}", normalize_text(container.get(key))))
    for key in ["carrier_hint", "channel_hint", "service_hint", "template_hint", "order_no", "tracking_no", "wh_code"]:
        values.append((f"metadata.{key}", normalize_text(metadata.get(key))))
    values.extend(flatten_named_values("metadata.source_fields", fields))
    values.extend(flatten_named_values("metadata.raw_item", raw_item))
    return [(name, value) for name, value in values if value]


def carrier_from_metadata(metadata: dict[str, Any]) -> str:
    for _name, value in metadata_named_values(metadata):
        carrier = normalize_carrier(value)
        if carrier:
            return carrier
    return ""


def template_from_metadata(metadata: dict[str, Any], carrier_hint: str) -> TemplateDecision | None:
    for name, value in metadata_named_values(metadata):
        decision = find_template(value, name, "高", carrier_hint)
        if decision:
            decision.note = f"{name} 命中 {decision.matched_text}"
            return decision
    return None


def template_from_filename(result: VerifyResult, carrier_hint: str) -> TemplateDecision | None:
    for source, text in [
        ("文件名", result.file_name),
        ("文件路径", result.file_path),
    ]:
        decision = find_template(text, source, "高", carrier_hint)
        if decision:
            decision.note = f"{source} 命中 {decision.matched_text}"
            return decision
    return None


def template_from_verify_result(result: VerifyResult, carrier_hint: str) -> TemplateDecision | None:
    template_code = first_non_empty(result.template_code)
    if template_code not in {"0024", "CBT", "CBS"}:
        return None
    raw_sub_code = first_non_empty(getattr(result, "template_sub_code", ""))
    marker = first_non_empty(result.template_marker, template_code)
    sub_code = template_subdivision(template_code, raw_sub_code, marker)
    display_marker = marker if template_code == "0024" else template_code
    source = first_non_empty(result.template_source, "PDF/OCR识别")
    if "OCR" in source.upper():
        confidence = "中"
    else:
        confidence = confidence_to_zh(result.template_confidence, default="高")
    decision = TemplateDecision(
        carrier=carrier_hint or UNKNOWN,
        label_template=template_code,
        template_subdivision=sub_code,
        matched_text=display_marker,
        matched_rule=f"{template_code}识别结果",
        source=source,
        confidence=confidence,
        note=f"{source} 命中 {display_marker}",
        need_review=(template_code == "0024" and sub_code == UNKNOWN),
    )
    return decision


def make_template_decision(result: VerifyResult, metadata: dict[str, Any]) -> TemplateDecision:
    carrier = first_non_empty(
        normalize_carrier(result.last_mile_carrier),
        normalize_carrier(result.carrier),
        carrier_from_metadata(metadata),
    ) or UNKNOWN

    verify_decision = template_from_verify_result(result, carrier)
    if verify_decision and "OCR" not in normalize_for_match(verify_decision.source):
        if verify_decision.carrier == UNKNOWN:
            verify_decision.carrier = carrier
        return verify_decision

    for detector in (
        lambda: template_from_metadata(metadata, carrier),
        lambda: template_from_filename(result, carrier),
    ):
        decision = detector()
        if decision:
            if decision.carrier == UNKNOWN:
                decision.carrier = carrier
            return decision

    if verify_decision:
        if verify_decision.carrier == UNKNOWN:
            verify_decision.carrier = carrier
        return verify_decision

    return TemplateDecision(
        carrier=carrier,
        label_template=NORMAL_LABEL_TEMPLATE,
        template_subdivision="",
        source="默认规则",
        confidence="中",
        note="未命中 0024/CBT/CBS，按普通面单处理",
        need_review=False,
    )


def tracking_no(result: VerifyResult, metadata: dict[str, Any]) -> str:
    return first_non_empty(
        result.barcode_tracking,
        result.pdf_text_tracking,
        result.source_tracking,
        result.filename_tracking,
        metadata.get("tracking_no"),
    )


def order_no(result: VerifyResult, metadata: dict[str, Any]) -> str:
    fields = dict_or_empty(metadata.get("source_fields"))
    return first_non_empty(
        result.download_order_no,
        metadata.get("order_no"),
        fields.get("deliveryNo"),
        fields.get("sourceNo"),
        fields.get("platformOrderNo"),
        fields.get("referOrderNo"),
    )


def result_to_business_row(result: VerifyResult, metadata: dict[str, Any]) -> dict[str, object]:
    decision = make_template_decision(result, metadata)
    row = asdict(result)
    row.update(
        {
            "order_no": order_no(result, metadata),
            "tracking_no": tracking_no(result, metadata),
            "carrier": decision.carrier,
            "label_template": decision.label_template,
            "template_subdivision": decision.template_subdivision,
            "recognition_basis": decision.note,
            "recognition_confidence": decision.confidence,
            "template_note": decision.note,
            "template_matched_text": decision.matched_text,
            "template_matched_rule": decision.matched_rule,
            "template_source": decision.source,
            "verify_status_zh": status_to_zh(result),
            "need_review": bool(result.need_review or result.ocr_needed or decision.need_review),
        }
    )
    return row


def display_value(value: object, mode: str) -> object:
    if mode == "yes_no":
        return "是" if bool(value) else "否"
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    return value


def export_results(results: list[VerifyResult], output_dir: Path, output_name: str) -> tuple[Path, Path]:
    ensure_dir(output_dir)
    excel_path = output_dir / f"{output_name}.xlsx"
    json_path = output_dir / f"{output_name}.json"
    write_excel(results, excel_path)
    write_json(results, json_path)
    log(f"Excel输出: {excel_path}")
    log(f"JSON输出: {json_path}")
    return excel_path, json_path


def append_result_sheet(ws, rows: list[dict[str, object]]) -> None:
    ws.append([label for label, _key, _mode in BUSINESS_COLUMNS])
    for row in rows:
        ws.append([display_value(row.get(key, ""), mode) for _label, key, mode in BUSINESS_COLUMNS])
    format_result_sheet(ws)


def format_result_sheet(ws) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    strong_fill = PatternFill("solid", fgColor="C6E0B4")
    weak_fill = PatternFill("solid", fgColor="FFF2CC")
    review_fill = PatternFill("solid", fgColor="F8CBAD")
    keys = [key for _label, key, _mode in BUSINESS_COLUMNS]
    status_col = keys.index("verify_status") + 1
    review_col = keys.index("need_review") + 1

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill

    for row_index in range(2, ws.max_row + 1):
        status = str(ws.cell(row_index, status_col).value or "")
        need_review = str(ws.cell(row_index, review_col).value or "") == "是"
        fill = review_fill if need_review else strong_fill if status in TRACKING_STATUS_STRONG else weak_fill
        for cell in ws[row_index]:
            cell.fill = fill

    autosize(ws)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def autosize(ws) -> None:
    for column_cells in ws.columns:
        column_letter = get_column_letter(column_cells[0].column)
        max_length = max(len(str(cell.value or "")) for cell in column_cells)
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 70)


def count_by(rows: list[dict[str, object]], key: str) -> Counter[str]:
    return Counter(str(row.get(key) or UNKNOWN) for row in rows)


def append_counter_section(ws, title: str, headers: tuple[str, str], counter: Counter[str]) -> None:
    ws.append([])
    ws.append([title])
    ws[ws.max_row][0].font = Font(bold=True)
    ws.append(list(headers))
    for name, count in sorted(counter.items(), key=lambda item: (-item[1], item[0])):
        ws.append([name or UNKNOWN, count])


def append_summary_sheet(ws, rows: list[dict[str, object]]) -> None:
    total = len(rows)
    auto_pass = sum(1 for row in rows if row.get("verify_status") in TRACKING_STATUS_STRONG)
    need_review = sum(1 for row in rows if bool(row.get("need_review")))

    ws.append(["统计项", "数值"])
    ws[1][0].font = Font(bold=True)
    ws[1][1].font = Font(bold=True)
    for label, value in [
        ("总文件数", total),
        ("自动通过数量", auto_pass),
        ("需要复核数量", need_review),
        ("条码识别成功数量", sum(1 for row in rows if bool(row.get("barcode_success")))),
        ("OCR触发数量", sum(1 for row in rows if bool(row.get("ocr_needed")))),
        ("0024数量", sum(1 for row in rows if row.get("label_template") == "0024")),
        ("CBT数量", sum(1 for row in rows if row.get("label_template") == "CBT")),
        ("CBS数量", sum(1 for row in rows if row.get("label_template") == "CBS")),
        ("普通面单数量", sum(1 for row in rows if row.get("label_template") == NORMAL_LABEL_TEMPLATE)),
    ]:
        ws.append([label, value])

    append_counter_section(ws, "按承运商统计", ("承运商", "数量"), count_by(rows, "carrier"))
    append_counter_section(ws, "按面单类型统计", ("面单类型", "数量"), count_by(rows, "label_template"))
    append_counter_section(ws, "按模板细分统计", ("模板细分", "数量"), count_by(rows, "template_subdivision"))

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for row in ws.iter_rows():
        if row[0].row == 1 or str(row[0].value or "").startswith("按"):
            for cell in row:
                cell.font = Font(bold=True)
                cell.fill = header_fill
    autosize(ws)


def business_rows(results: list[VerifyResult]) -> list[dict[str, object]]:
    metadata_rows = read_metadata_rows()
    index = metadata_index(metadata_rows)
    return [result_to_business_row(result, metadata_for_result(result, index)) for result in results]


def write_excel(results: list[VerifyResult], path: Path) -> None:
    rows = business_rows(results)
    wb = Workbook()

    ws_all = wb.active
    ws_all.title = "全部结果"
    append_result_sheet(ws_all, rows)

    ws_review = wb.create_sheet("异常复核")
    append_result_sheet(ws_review, [row for row in rows if bool(row.get("need_review"))])

    ws_summary = wb.create_sheet("统计汇总")
    append_summary_sheet(ws_summary, rows)

    try:
        wb.save(path)
    except Exception as exc:
        raise RuntimeError(f"Excel 写入失败: {path}; {exc}") from exc


def write_json(results: list[VerifyResult], path: Path) -> None:
    metadata_rows = read_metadata_rows()
    index = metadata_index(metadata_rows)
    data: list[dict[str, object]] = []
    for result in results:
        metadata = metadata_for_result(result, index)
        decision = make_template_decision(result, metadata)
        row = asdict(result)
        row["business_template"] = asdict(decision)
        row["business_need_review"] = bool(result.need_review or result.ocr_needed or decision.need_review)
        row["business_order_no"] = order_no(result, metadata)
        row["business_tracking_no"] = tracking_no(result, metadata)
        data.append(row)
    try:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(f"JSON 写入失败: {path}; {exc}") from exc
