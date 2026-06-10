from __future__ import annotations

import argparse
import csv
import json
import re
import signal
import sys
import time
import unicodedata
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps


LOCAL_DEPS = Path(__file__).resolve().parent / ".codex_deps"
if LOCAL_DEPS.exists():
    sys.path.insert(0, str(LOCAL_DEPS))


TRACKING_STATUS_STRONG = {
    "auto_pass_triple_verified",
    "auto_pass_barcode_verified",
    "auto_pass_text_verified",
    "auto_pass_verified_prefix",
}
REVIEW_STATUSES = {"review_conflict", "review_unknown", "barcode_failed"}
SOURCE_ONLY_STATUSES = {
    "auto_pass_source_only",
    "auto_pass_source_consistent",
    "source_only_low_confidence",
    "review_conflict",
    "review_unknown",
    "barcode_failed",
    "ocr_needed",
    "download_file_error",
}
OCR_QUEUE_STATUSES = {"ocr_needed", "download_file_error"}
SUPPORTED_IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff"}
SUPPORTED_INPUT_SUFFIXES = {".pdf", *SUPPORTED_IMAGE_SUFFIXES}
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
}
NEGATIVE_CONTEXT = (
    "PHONE",
    "TEL",
    "ACCOUNT",
    "REFERENCE",
    "REF",
    "ORDER",
    "SKU",
    "QTY",
    "WEIGHT",
    "ZIP",
    "POSTAL",
    "INVOICE",
    "CUSTOMER",
)
FEDEX_CONTEXT = ("FEDEX", "TRK#", "TRK #", "TRACKING ID", "TRACKING NUMBER")


@dataclass
class Candidate:
    carrier: str
    value: str
    source: str
    score: int
    context: str


@dataclass
class TemplateMatch:
    template_code: str = ""
    template_sub_code: str = ""
    template_marker: str = ""
    template_source: str = ""
    template_confidence: str = ""


@dataclass
class VerifyResult:
    file_name: str
    file_path: str
    file_type: str
    file_format: str
    is_image_label: bool
    carrier: str
    last_mile_carrier: str
    template_code: str
    template_sub_code: str
    template_marker: str
    carrier_display: str
    template_source: str
    template_confidence: str
    source_tracking: str
    filename_tracking: str
    pdf_text_tracking: str
    barcode_tracking: str
    download_order_no: str
    download_wave_no: str
    download_warehouse: str
    ocr_tracking: str
    all_text_candidates: str
    all_barcode_candidates: str
    verify_status: str
    verify_status_zh: str
    confidence: str
    confidence_zh: str
    reason: str
    reason_zh: str
    need_review: bool
    barcode_success: bool
    barcode_page_numbers: str
    barcode_processing_seconds: float
    barcode_matches_source: bool
    barcode_matches_pdf_text: bool
    barcode_verify_note: str
    barcode_error: str
    pdf_text_success: bool
    image_barcode_result: str
    image_processing_note: str
    ocr_needed: bool
    ocr_reason: str
    ocr_priority: str
    processing_seconds: float
    decoded_formats: str
    decoded_raw_values: str
    page_count: int | None
    recognized_at: str
    meta_carrier: str = ""
    meta_channel: str = ""
    meta_channel_code: str = ""
    meta_customer_code: str = ""
    meta_source: str = ""
    meta_carrier_conflict: bool = False
    meta_carrier_note: str = ""


TRACKING_PATTERNS: dict[str, re.Pattern[str]] = {
    "UPS": re.compile(r"(?<![0-9A-Z])1Z[0-9A-Z]{16}(?![0-9A-Z])", re.I),
    "LaserShip": re.compile(r"(?<![0-9A-Z])1LS[A-Z0-9]{9,24}(?![0-9A-Z])", re.I),
    "SwiftX": re.compile(r"(?<![0-9A-Z])SWX\d{12,24}(?![0-9A-Z])", re.I),
    "GOFO": re.compile(r"(?<![0-9A-Z])(?:GFUS|YT)[A-Z0-9]{10,24}(?![0-9A-Z])", re.I),
    "UniUni": re.compile(r"(?<![0-9A-Z])UUS[A-Z0-9]{8,24}(?![0-9A-Z])", re.I),
    "SpeedX": re.compile(r"(?<![0-9A-Z])SPX[A-Z0-9]{8,28}(?![0-9A-Z])", re.I),
    "Yanwen": re.compile(r"(?<![0-9A-Z])YW[A-Z0-9]{8,28}(?![0-9A-Z])", re.I),
    "OnTrac": re.compile(r"(?<![0-9A-Z])D100\d{8,20}(?![0-9A-Z])", re.I),
    "USPS": re.compile(r"(?<!\d)(?:92|93|94)\d(?:[\s-]?\d){19,25}(?!\d)", re.I),
    "FedEx": re.compile(r"(?<!\d)(?:\d[\s-]?){12}(?!\d)|(?<!\d)(?:\d{15}|\d{20}|\d{22})(?!\d)"),
}

TEMPLATE_RULES: list[tuple[str, str, re.Pattern[str]]] = [
    ("CBT", "CBT", re.compile(r"\b(?:TT[-\s]?CBT|CBT\s*[AB]\s*\d{2})\b", re.I)),
    ("CBS", "CBS", re.compile(r"\b(?:LAX[-\s]?CBS|CBS)\b", re.I)),
    ("0024", "0024", re.compile(r"\b0024(?:\s*0?[12])?\b", re.I)),
    ("天翼", "天翼", re.compile(r"天翼", re.I)),
    ("官方", "官方", re.compile(r"官方", re.I)),
    ("达通", "达通", re.compile(r"达通", re.I)),
    ("联递", "联递", re.compile(r"联递", re.I)),
    ("鲸准", "鲸准", re.compile(r"鲸准", re.I)),
    ("平台面单", "平台面单", re.compile(r"平台面单", re.I)),
]

OCR_ENGINE: Any | None | bool = None
NORMAL_LABEL_SUBDIVISION = "正常面单"
VALID_0024_SUB_CODES = {"01", "02"}


def normalize(value: str) -> str:
    return re.sub(r"[^0-9A-Za-z]", "", value or "").upper()


def unique(values: list[str]) -> list[str]:
    return list(dict.fromkeys(value for value in values if value))


def yes_no(value: bool) -> str:
    return "是" if value else "否"


def file_format(path: Path) -> str:
    return path.suffix.lower().lstrip(".")


def file_type(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return "PDF"
    if suffix in SUPPORTED_IMAGE_SUFFIXES:
        return "图片"
    return "未知"


def is_image_file(path: Path) -> bool:
    return path.suffix.lower() in SUPPORTED_IMAGE_SUFFIXES


def is_supported_input(path: Path) -> bool:
    return path.suffix.lower() in SUPPORTED_INPUT_SUFFIXES


def status_to_zh(status: str) -> str:
    return STATUS_ZH.get(status, status)


def confidence_to_zh(confidence: str) -> str:
    return CONFIDENCE_ZH.get(confidence, confidence)


def normalize_marker(value: str) -> str:
    text = re.sub(r"\s+", " ", value or "").strip()
    text = text.replace(" - ", "-").replace(" -", "-").replace("- ", "-")
    return text.upper() if re.search(r"[A-Za-z]", text) else text


def normalize_0024_sub_code(raw: str) -> str:
    clean = re.sub(r"[^0-9O]", "", raw.upper()).replace("O", "0")
    if not clean:
        return ""
    try:
        code = f"{int(clean):02d}"
    except ValueError:
        return ""
    return code if code in VALID_0024_SUB_CODES else ""


def template_marker(template_code: str, template_sub_code: str = "") -> str:
    if template_code == "0024" and template_sub_code:
        return f"0024-{template_sub_code}"
    return template_code


def normalize_template_text(text: str) -> tuple[str, str]:
    normalized = unicodedata.normalize("NFKC", text or "").upper()
    normalized = normalized.replace("–", "-").replace("—", "-").replace("_", "-")
    spaced = re.sub(r"\s+", " ", normalized).strip()
    compact = re.sub(r"[^0-9A-Z\u4e00-\u9fff]+", "", normalized)
    return spaced, compact


def match_template_text(text: str, source: str, confidence: str) -> TemplateMatch:
    if not text:
        return TemplateMatch()

    spaced, compact = normalize_template_text(text)

    if "CBT" in compact or re.search(r"\bTT[-\s]*CBT\b|\bCBT\b", spaced, re.I):
        return TemplateMatch(
            template_code="CBT",
            template_sub_code=NORMAL_LABEL_SUBDIVISION,
            template_marker="CBT",
            template_source=source,
            template_confidence=confidence,
        )

    if "LAXCBS" in compact or re.search(r"\bCBS\b", spaced, re.I):
        return TemplateMatch(
            template_code="CBS",
            template_sub_code=NORMAL_LABEL_SUBDIVISION,
            template_marker="CBS",
            template_source=source,
            template_confidence=confidence,
        )

    code_0024_match = re.search(r"(?<![0-9A-Z])0024(?:\s*[-_/]?\s*([0O]?[12]))?(?![0-9A-Z])", spaced, re.I)
    if code_0024_match:
        sub_code = normalize_0024_sub_code(code_0024_match.group(1) or "")
        return TemplateMatch(
            template_code="0024",
            template_sub_code=sub_code,
            template_marker=template_marker("0024", sub_code),
            template_source=source,
            template_confidence=confidence,
        )

    return TemplateMatch()


def match_template_corner_text(text: str, source: str, confidence: str) -> TemplateMatch:
    match = match_template_text(text, source, confidence)
    if match.template_code and (match.template_code != "0024" or match.template_sub_code):
        return match

    _spaced, compact = normalize_template_text(text)
    if re.search(r"[AB]0[1-5]", compact):
        return TemplateMatch(
            template_code="CBT",
            template_sub_code=NORMAL_LABEL_SUBDIVISION,
            template_marker="CBT",
            template_source=source,
            template_confidence=confidence,
        )

    index_0024 = compact.find("0024")
    if index_0024 >= 0:
        tail = compact[index_0024 + 4 : index_0024 + 16]
        sub_match = re.search(r"0?([12])", tail)
        sub_code = normalize_0024_sub_code(sub_match.group(0) if sub_match else "")
        if sub_code:
            return TemplateMatch(
                template_code="0024",
                template_sub_code=sub_code,
                template_marker=template_marker("0024", sub_code),
                template_source=source,
                template_confidence=confidence,
            )

    return match


def merge_template_match(primary: TemplateMatch, fallback: TemplateMatch) -> TemplateMatch:
    return primary if primary.template_code else fallback


def load_ocr_engine() -> Any | None:
    global OCR_ENGINE
    if OCR_ENGINE is False:
        return None
    if OCR_ENGINE is not None:
        return OCR_ENGINE
    try:
        import pytesseract  # type: ignore

        OCR_ENGINE = pytesseract
        return pytesseract
    except Exception:
        OCR_ENGINE = False
        return None


def ocr_image(image: Image.Image) -> str:
    engine = load_ocr_engine()
    if engine is None:
        return ""
    try:
        return str(engine.image_to_string(image, lang="chi_sim+eng") or "")
    except Exception:
        try:
            return str(engine.image_to_string(image, lang="eng") or "")
        except Exception:
            return ""


def ocr_template_corner(image: Image.Image) -> str:
    engine = load_ocr_engine()
    if engine is None:
        return ""
    width, height = image.size
    crops = [
        image.crop((0, int(height * 0.86), int(width * 0.22), int(height * 0.99))),
        image.crop((0, int(height * 0.84), int(width * 0.38), height)),
    ]
    config = "--psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    texts: list[str] = []
    for crop in crops:
        gray = ImageOps.grayscale(crop)
        bw = gray.point(lambda value: 255 if value > 150 else 0)
        variants = [crop, bw.convert("RGB")]
        for variant in variants:
            scaled = variant.resize((variant.width * 3, variant.height * 3))
            try:
                texts.append(str(engine.image_to_string(scaled, lang="eng", config=config) or ""))
            except Exception:
                continue
    return "\n".join(texts)


def crop_bottom(image: Image.Image, ratio: float = 0.30) -> Image.Image:
    width, height = image.size
    top = max(0, int(height * (1 - ratio)))
    return image.crop((0, top, width, height))


def template_ocr_regions(image: Image.Image) -> list[tuple[str, Image.Image]]:
    width, height = image.size
    return [
        ("底部OCR", image.crop((0, int(height * 0.65), width, height))),
        ("顶部OCR", image.crop((0, 0, width, int(height * 0.25)))),
        ("右上OCR", image.crop((int(width * 0.55), 0, width, int(height * 0.40)))),
        ("左下OCR", image.crop((0, int(height * 0.60), int(width * 0.55), height))),
        ("全页OCR", image),
    ]


def first_page_images_for_template(path: Path, dpi: int, is_image_label: bool) -> tuple[list[Image.Image], str]:
    if is_image_label:
        images, _page_count, error = load_image_file(path)
    else:
        images, _page_count, error = render_pdf_images(path, dpi=dpi, max_pages=1)
    return [image for _page_no, image in images[:1]], error


def detect_template(
    path: Path,
    pdf_text: str,
    decoded_raw_values: list[str],
    dpi: int,
    is_image_label: bool,
    ocr_enabled: bool = False,
    allow_full_page_ocr: bool = True,
) -> TemplateMatch:
    text_chunks = [
        (pdf_text, "PDF文字层", "高"),
        ("\n".join(decoded_raw_values), "条码内容", "中"),
    ]
    for text, source, confidence in text_chunks:
        match = match_template_text(text, source, confidence)
        if match.template_code:
            return match

    images, image_error = first_page_images_for_template(path, dpi=dpi, is_image_label=is_image_label)
    if not image_error and images and load_ocr_engine() is not None:
        for image in images:
            match = match_template_corner_text(ocr_template_corner(image), "左下模板OCR", "中")
            if match.template_code:
                return match

    if not ocr_enabled:
        return TemplateMatch()

    if not image_error and images and load_ocr_engine() is not None:
        for region_name in ["底部OCR", "顶部OCR", "右上OCR", "左下OCR", "全页OCR"]:
            if region_name == "全页OCR" and not allow_full_page_ocr:
                continue
            region_texts: list[str] = []
            for image in images:
                for current_name, crop in template_ocr_regions(image):
                    if current_name == region_name:
                        region_texts.append(ocr_image(crop))
                        break
            match = match_template_text("\n".join(region_texts), region_name, "中")
            if match.template_code:
                return match

    return TemplateMatch()


def carrier_display(last_mile_carrier: str, template_code: str, template_sub_code: str = "") -> str:
    template_text = "-".join(part for part in [template_code, template_sub_code] if part)
    if last_mile_carrier and template_text:
        return f"{template_text}-{last_mile_carrier}" if template_code == "平台面单" else f"{last_mile_carrier}-{template_text}"
    return last_mile_carrier or template_text or ""


def reason_to_zh(
    status: str,
    barcode_success: bool,
    source_tracking: str,
    barcode_values: list[str],
    text_values: list[str],
    is_image_label: bool,
    file_error: str = "",
) -> str:
    if status == "auto_pass_triple_verified":
        return "条码、PDF文字层和下载数据三者一致，自动通过"
    if status == "auto_pass_barcode_verified":
        if source_tracking and source_tracking in barcode_values:
            return "条码反读成功，且与下载数据一致，自动通过"
        return "条码反读成功，且与PDF文字层一致，自动通过"
    if status == "auto_pass_text_verified":
        return "PDF文字层识别成功，但未识别到条码，建议抽查"
    if status == "auto_pass_verified_prefix":
        return "USPS文字层追踪号存在粘连，前缀与下载数据一致，可通过"
    if status == "source_only_low_confidence":
        return "仅下载数据或文件名可用，缺少条码/PDF文字层强校验，需复核"
    if status == "review_conflict":
        return "条码、PDF文字层或下载数据不一致，需人工复核"
    if status == "barcode_failed":
        return "条码识别失败，建议进入OCR深度识别或人工复核"
    if status == "ocr_needed":
        if is_image_label:
            if barcode_success:
                return "图片面单条码已反读，但未提取到有效追踪号，建议OCR深度识别"
            return "图片面单条码识别失败，建议OCR深度识别"
        return "未识别到有效追踪号，建议OCR深度识别"
    if status == "download_file_error":
        return f"下载文件异常，无法打开或格式不支持{(': ' + file_error) if file_error else ''}"
    if status == "review_unknown":
        return "未识别到有效追踪号，建议OCR深度识别"
    if is_image_label and barcode_success:
        return "下载文件为图片格式，已按图片面单处理"
    return "需人工复核"


def barcode_note(
    barcode_success: bool,
    barcode_values: list[str],
    source_tracking: str,
    text_values: list[str],
    barcode_error: str,
) -> str:
    if not barcode_success:
        return f"条码未反读成功{(': ' + barcode_error) if barcode_error else ''}"
    if not barcode_values:
        return "条码通过 zxing-cpp 反读成功，但未提取到有效追踪号"
    if source_tracking and source_tracking in barcode_values:
        return "条码通过 zxing-cpp 反读成功，且与下载数据一致"
    if any(value in text_values for value in barcode_values):
        return "条码通过 zxing-cpp 反读成功，且与PDF文字层一致"
    if source_tracking:
        return "条码通过 zxing-cpp 反读成功，但与下载数据不一致"
    return "条码通过 zxing-cpp 反读成功"


def ocr_decision(
    status: str,
    carrier: str,
    source_tracking: str,
    text_values: list[str],
    barcode_values: list[str],
    barcode_decoded: bool,
    is_image_label: bool,
    file_error: str,
) -> tuple[bool, str, str]:
    if status in TRACKING_STATUS_STRONG:
        return False, "强校验已通过", ""
    if status == "download_file_error" or file_error:
        return True, "下载文件异常，无法打开或格式不支持", "高"
    if status == "review_conflict":
        return True, "多来源结果冲突", "高"
    if is_image_label and not barcode_values:
        if barcode_decoded:
            return True, "图片条码已反读但未提取到有效追踪号", "高"
        return True, "图片条码识别失败", "高"
    if status in {"barcode_failed", "ocr_needed"}:
        return True, "条码识别失败", "高"
    if status == "source_only_low_confidence":
        return True, "低置信度，仅下载数据/文件名可用", "中"
    if not text_values and not barcode_values:
        return True, "PDF文字层缺失且条码失败", "高"
    if carrier == "UNKNOWN" or not source_tracking:
        return True, "无法识别承运商或追踪号", "中"
    if status == "review_unknown":
        return True, "无法识别有效追踪号", "中"
    return False, "不需要OCR", ""


def infer_carrier(value: str) -> str:
    tracking = normalize(value)
    if tracking.startswith("1Z"):
        return "UPS"
    if tracking.startswith("1LS"):
        return "LaserShip"
    if tracking.startswith("SWX"):
        return "SwiftX"
    if tracking.startswith(("GFUS", "YT")):
        return "GOFO"
    if tracking.startswith("UUS"):
        return "UniUni"
    if tracking.startswith("SPX"):
        return "SpeedX"
    if tracking.startswith("YW"):
        return "Yanwen"
    if tracking.startswith("D100"):
        return "OnTrac"
    if re.fullmatch(r"9\d{20,30}", tracking):
        return "USPS"
    if re.fullmatch(r"\d{12}|\d{15}|\d{20}|\d{22}", tracking):
        return "FedEx"
    return "UNKNOWN"


def normalize_carrier_name(value: str) -> str:
    text = normalize(value)
    if not text or text in {"UNKNOWN", "OTHER"}:
        return ""
    aliases = {
        "USPS": "USPS",
        "POSTAL": "USPS",
        "FEDEX": "FedEx",
        "FDX": "FedEx",
        "UPS": "UPS",
        "GOFO": "GOFO",
        "GFUS": "GOFO",
        "SWIFTX": "SwiftX",
        "SWX": "SwiftX",
        "UNIUNI": "UniUni",
        "UUS": "UniUni",
        "LASERSHIP": "LaserShip",
        "1LS": "LaserShip",
        "SPEEDX": "SpeedX",
        "SPX": "SpeedX",
        "YANWEN": "Yanwen",
        "YW": "Yanwen",
        "ONTRAC": "OnTrac",
    }
    return aliases.get(text, value.strip())


def load_zxingcpp() -> Any | None:
    try:
        import zxingcpp  # type: ignore

        return zxingcpp
    except ImportError:
        print("Missing dependency: zxing-cpp", flush=True)
        print("Install with one of these commands:", flush=True)
        print("  python -m pip install --target .codex_deps zxing-cpp", flush=True)
        print("  python -m pip install zxing-cpp", flush=True)
        return None


def context_window(text: str, start: int, end: int, size: int = 90) -> str:
    return text[max(0, start - size) : min(len(text), end + size)]


def score_candidate(
    carrier: str,
    value: str,
    source: str,
    context: str,
    source_tracking: str,
    barcode_values: list[str] | None = None,
) -> int:
    score = 0
    context_upper = context.upper()
    barcode_values = barcode_values or []
    if value == source_tracking:
        score += 50
    if value in barcode_values:
        score += 80
    if any(token in context_upper for token in ("TRK#", "TRK #", "TRACKING ID", "TRACKING NUMBER")):
        score += 40
    if carrier != "UNKNOWN" and carrier.upper() in context_upper:
        score += 20
    if any(token in context_upper for token in NEGATIVE_CONTEXT):
        score -= 50
    if source in {"filename", "wms", "barcode"}:
        score += 20
    return score


def extract_candidates(
    text: str,
    source: str,
    carrier_hint: str = "",
    source_tracking: str = "",
    barcode_values: list[str] | None = None,
) -> list[Candidate]:
    carriers = [carrier_hint] if carrier_hint and carrier_hint != "UNKNOWN" else list(TRACKING_PATTERNS)
    candidates: list[Candidate] = []
    seen: set[tuple[str, str, str]] = set()
    for carrier in carriers:
        pattern = TRACKING_PATTERNS.get(carrier)
        if not pattern:
            continue
        for match in pattern.finditer(text or ""):
            value = normalize(match.group(0))
            if not value:
                continue
            if carrier == "FedEx":
                if value.startswith(("92", "93", "94")):
                    continue
                if source == "pdf_text":
                    ctx = context_window(text, match.start(), match.end())
                    if not any(token in ctx.upper() for token in FEDEX_CONTEXT):
                        continue
            ctx = context_window(text, match.start(), match.end()) if source == "pdf_text" else source
            key = (carrier, value, source)
            if key in seen:
                continue
            seen.add(key)
            candidates.append(
                Candidate(
                    carrier=carrier,
                    value=value,
                    source=source,
                    score=score_candidate(carrier, value, source, ctx, source_tracking, barcode_values),
                    context=ctx.replace("\n", " ")[:180],
                )
            )
    return candidates


def parse_tracking_list(value: str) -> list[str]:
    values: list[str] = []
    for part in re.split(r"[,;\s|]+", value or ""):
        normalized = normalize(part)
        if normalized:
            values.append(normalized)
    return unique(values)


def list_input_files(input_dir: Path) -> list[Path]:
    if input_dir.is_file():
        return [input_dir] if is_supported_input(input_dir) else []
    if not input_dir.exists():
        raise FileNotFoundError(f"Input path does not exist: {input_dir}")
    return sorted(path for path in input_dir.rglob("*") if path.is_file() and is_supported_input(path))


def read_json(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        raise FileNotFoundError(f"JSON path does not exist: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise RuntimeError(f"Failed to read JSON: {path}; {exc}") from exc
    if not isinstance(data, list):
        raise ValueError(f"JSON must contain a list: {path}")
    return data


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    for encoding in ("utf-8-sig", "utf-8", "gbk"):
        try:
            with path.open("r", newline="", encoding=encoding) as file:
                return list(csv.DictReader(file))
        except UnicodeDecodeError:
            continue
    return []


def load_wms_tracking(path: Path) -> dict[str, str]:
    records: dict[str, str] = {}
    for row in read_csv_rows(path):
        tracking = normalize(row.get("expressNo", ""))
        file_path = row.get("filePath", "")
        if not tracking:
            continue
        records.setdefault(tracking.lower(), tracking)
        if file_path:
            records[Path(file_path).name.lower()] = tracking
    return records


def source_tracking_from_row(row: dict[str, Any], wms_records: dict[str, str]) -> str:
    for key in ("log_tracking_no", "source_tracking", "tracking_no", "filename_tracking_no"):
        value = normalize(str(row.get(key) or ""))
        if value:
            return value
    file_path = Path(str(row.get("file_path") or ""))
    by_name = wms_records.get(file_path.name.lower())
    if by_name:
        return by_name
    filename_candidates = extract_candidates(file_path.stem, "filename")
    return filename_candidates[0].value if filename_candidates else ""


def load_metadata_index(path: Path) -> tuple[dict[str, dict[str, str]], dict[str, dict[str, str]]]:
    """从 metadata.jsonl 构建两个索引: {filename: record} 和 {delivery_no: record}"""
    by_filename: dict[str, dict[str, str]] = {}
    by_delivery_no: dict[str, dict[str, str]] = {}
    if not path.exists():
        return by_filename, by_delivery_no
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                row = json.loads(line)
            except json.JSONDecodeError:
                continue
            if not isinstance(row, dict):
                continue
            source_fields = row.get("source_fields") if isinstance(row.get("source_fields"), dict) else {}
            recognition = row.get("metadata_recognition") if isinstance(row.get("metadata_recognition"), dict) else {}
            meta = {
                "carrier": recognition.get("carrier", "") or source_fields.get("logisticsCarrier", ""),
                "channel": recognition.get("channel", "") or source_fields.get("logisticsChannelName", "")
                           or source_fields.get("logisticsChannel", ""),
                "channel_code": source_fields.get("logisticsChannel", ""),
                "customer_code": source_fields.get("customerCode", ""),
                "wh_code": source_fields.get("whCode", "") or row.get("wh_code", ""),
                "source_no": source_fields.get("sourceNo", ""),
                "platform_order_no": source_fields.get("platformOrderNo", ""),
                "tracking_no": row.get("tracking_no", ""),
                "order_no": row.get("order_no", ""),
            }
            fname = row.get("file_name", "")
            if fname:
                by_filename[fname.lower()] = meta
            dno = meta["order_no"]
            if dno:
                by_delivery_no[dno.lower()] = meta
    return by_filename, by_delivery_no


def _extract_delivery_no_from_filename(stem: str) -> str:
    """从文件名提取 deliveryNo，格式如 DO9852606100RT 或 DO9852606100RT_xxx"""
    m = re.match(r"^(DO[A-Z0-9]+)", stem, re.IGNORECASE)
    return m.group(1) if m else ""


def enrich_row_from_metadata(
    row: dict[str, Any],
    meta_by_file: dict[str, dict[str, str]],
    meta_by_dno: dict[str, dict[str, str]],
) -> None:
    """用 metadata 索引丰富 row，注入 meta_carrier/meta_channel 等字段。"""
    fname = str(row.get("file_name") or Path(str(row.get("file_path") or "")).name).lower()
    meta = meta_by_file.get(fname)
    if not meta:
        # 从文件名提取 delivery_no 再反查
        stem = Path(fname).stem
        dno = _extract_delivery_no_from_filename(stem)
        if dno:
            meta = meta_by_dno.get(dno.lower())
    if meta:
        row["meta_carrier"] = meta.get("carrier", "")
        row["meta_channel"] = meta.get("channel", "")
        row["meta_channel_code"] = meta.get("channel_code", "")
        row["meta_customer_code"] = meta.get("customer_code", "")
        row["meta_wh_code"] = meta.get("wh_code", "")
        row["meta_source"] = "metadata"


def rows_from_input_dir(input_dir: Path, wms_records: dict[str, str],
                        meta_by_file: dict[str, dict[str, str]] | None = None,
                        meta_by_dno: dict[str, dict[str, str]] | None = None) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for path in list_input_files(input_dir):
        filename_candidates = extract_candidates(path.stem, "filename")
        filename_tracking = filename_candidates[0].value if filename_candidates else ""
        source_tracking = wms_records.get(path.name.lower()) or filename_tracking
        row = {
            "file_name": path.name,
            "file_path": str(path.resolve()),
            "tracking_no": source_tracking,
            "filename_tracking_no": filename_tracking,
            "log_tracking_no": wms_records.get(source_tracking.lower(), ""),
            "pdf_text_tracking_no": "",
            "verification_status": "",
        }
        if meta_by_file is not None and meta_by_dno is not None:
            enrich_row_from_metadata(row, meta_by_file, meta_by_dno)
        rows.append(row)
    return rows


def read_pdf_text(path: Path, max_pages: int) -> tuple[str, int | None, str]:
    try:
        parts: list[str] = []
        with fitz.open(str(path)) as doc:
            page_count = doc.page_count
            for page in doc[: max(max_pages, 1)]:
                parts.append(page.get_text("text") or "")
        return "\n".join(parts), page_count, ""
    except Exception as exc:
        return "", None, str(exc)


def render_pdf_images(path: Path, dpi: int, max_pages: int) -> tuple[list[tuple[int, Image.Image]], int | None, str]:
    try:
        images: list[tuple[int, Image.Image]] = []
        zoom = dpi / 72
        with fitz.open(str(path)) as doc:
            page_count = doc.page_count
            for page_index, page in enumerate(doc[: max(max_pages, 1)], start=1):
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                images.append((page_index, Image.frombytes("RGB", (pix.width, pix.height), pix.samples)))
        return images, page_count, ""
    except Exception as exc:
        return [], None, str(exc)


def load_image_file(path: Path) -> tuple[list[tuple[int, Image.Image]], int | None, str]:
    try:
        with Image.open(path) as image:
            return [(1, image.convert("RGB").copy())], 1, ""
    except Exception as exc:
        return [], None, str(exc)


def image_variants(image: Image.Image, rotations: bool, enhanced: bool) -> list[Image.Image]:
    variants = [image]
    if enhanced:
        gray = ImageOps.grayscale(image)
        bw = gray.point(lambda value: 255 if value > 160 else 0, mode="1").convert("RGB")
        variants.extend([gray.convert("RGB"), bw])
    if rotations:
        rotated: list[Image.Image] = []
        for variant in variants:
            rotated.extend([variant.rotate(90, expand=True), variant.rotate(180, expand=True), variant.rotate(270, expand=True)])
        variants.extend(rotated)
    return variants


def decode_images(zxingcpp: Any, images: list[tuple[int, Image.Image]], rotations: bool, enhanced: bool) -> tuple[list[dict[str, str]], str]:
    decoded: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    errors: list[str] = []
    for page_no, image in images:
        for variant in image_variants(image, rotations=rotations, enhanced=enhanced):
            try:
                barcodes = zxingcpp.read_barcodes(variant)
            except Exception as exc:
                errors.append(str(exc))
                continue
            for barcode in barcodes:
                text = str(getattr(barcode, "text", "") or "")
                fmt = str(getattr(barcode, "format", "") or "")
                if not text:
                    continue
                key = (str(page_no), fmt, text)
                if key not in seen:
                    seen.add(key)
                    decoded.append({"page": str(page_no), "format": fmt, "text": text})
    return decoded, "; ".join(unique(errors)[:3])


class Timeout:
    def __init__(self, seconds: int):
        self.seconds = seconds
        self._old_handler: Any = None

    def __enter__(self) -> None:
        if self.seconds <= 0 or not hasattr(signal, "SIGALRM"):
            return
        self._old_handler = signal.signal(signal.SIGALRM, self._handle_timeout)
        signal.alarm(self.seconds)

    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> None:
        if self.seconds > 0 and hasattr(signal, "SIGALRM"):
            signal.alarm(0)
            if self._old_handler is not None:
                signal.signal(signal.SIGALRM, self._old_handler)

    def _handle_timeout(self, _signum: int, _frame: Any) -> None:
        raise TimeoutError(f"timeout after {self.seconds} seconds")


def decode_barcode_for_file(
    zxingcpp: Any | None,
    path: Path,
    dpi: int,
    max_pages: int,
    rotations: bool,
    timeout: int,
) -> tuple[list[dict[str, str]], int | None, str]:
    if zxingcpp is None:
        return [], None, "zxing-cpp is not installed"
    try:
        with Timeout(timeout):
            if is_image_file(path):
                images, page_count, render_error = load_image_file(path)
            else:
                images, page_count, render_error = render_pdf_images(path, dpi=dpi, max_pages=max_pages)
            if render_error:
                return [], page_count, f"download_file_error: {render_error}"
            decoded, decode_error = decode_images(zxingcpp, images, rotations=False, enhanced=False)
            if decoded:
                return decoded, page_count, decode_error
            decoded, decode_error = decode_images(zxingcpp, images, rotations=False, enhanced=True)
            if decoded:
                return decoded, page_count, decode_error
            if rotations:
                decoded, decode_error = decode_images(zxingcpp, images, rotations=True, enhanced=True)
            return decoded, page_count, decode_error
    except TimeoutError:
        return [], None, f"timeout after {timeout} seconds"
    except Exception as exc:
        return [], None, str(exc)


def choose_best(candidates: list[Candidate]) -> Candidate | None:
    if not candidates:
        return None
    return sorted(candidates, key=lambda item: item.score, reverse=True)[0]


def conflict(values_by_source: dict[str, list[str]], source_tracking: str) -> tuple[bool, str]:
    known: dict[str, set[str]] = {
        source: set(values)
        for source, values in values_by_source.items()
        if values
    }
    sources = list(known)
    for i, left_source in enumerate(sources):
        for right_source in sources[i + 1 :]:
            left_values = known[left_source]
            right_values = known[right_source]
            if source_tracking and source_tracking in left_values and source_tracking in right_values:
                continue
            if left_values.isdisjoint(right_values):
                return True, f"{left_source}={','.join(left_values)} conflicts with {right_source}={','.join(right_values)}"
    return False, ""


def determine_status(
    source_tracking: str,
    text_values: list[str],
    barcode_values: list[str],
    carrier: str,
    barcode_error: str,
) -> tuple[str, str, str, bool]:
    source_values = [source_tracking] if source_tracking else []
    prefix_text = [
        value
        for value in text_values
        if source_tracking
        and carrier == "USPS"
        and value.startswith(source_tracking)
        and 0 < len(value) - len(source_tracking) <= 4
    ]
    if prefix_text and source_tracking and (not barcode_values or source_tracking in barcode_values):
        return "auto_pass_verified_prefix", "high", "USPS text layer has glued digits, prefix matched source tracking", False

    is_conflict, conflict_reason = conflict(
        {"source": source_values, "pdf_text": text_values, "barcode": barcode_values},
        source_tracking,
    )
    if is_conflict:
        return "review_conflict", "low", conflict_reason, True

    if not source_tracking and not text_values and not barcode_values:
        return "review_unknown", "low", "no reliable tracking number found", True

    if source_tracking:
        text_exact = source_tracking in text_values
        barcode_exact = source_tracking in barcode_values
        if text_exact and barcode_exact:
            return "auto_pass_triple_verified", "high", "source, PDF text and barcode all matched", False
        if barcode_exact:
            return "auto_pass_barcode_verified", "high", "source tracking matched barcode result", False
        if text_exact:
            return "auto_pass_text_verified", "high", "source tracking matched PDF text layer", False
        if prefix_text:
            return "auto_pass_verified_prefix", "high", "USPS text layer has glued digits, prefix matched source tracking", False
        if barcode_error:
            return "source_only_low_confidence", "low", f"source only; barcode not verified: {barcode_error}", True
        return "source_only_low_confidence", "low", "source only; PDF text and barcode did not verify tracking", True

    if len(unique(text_values + barcode_values)) == 1 and text_values and barcode_values:
        return "auto_pass_barcode_verified", "medium", "PDF text and barcode matched but no source tracking", False
    return "review_unknown", "low", "no source tracking to prove final result", True


def values_from_candidates(candidates: list[Candidate]) -> list[str]:
    return unique([candidate.value for candidate in candidates])


def error_result(row: dict[str, Any], wms_records: dict[str, str], exc: Exception | str) -> VerifyResult:
    file_path = Path(str(row.get("file_path") or row.get("pdf_path") or ""))
    error_text = str(exc)
    current_file_type = file_type(file_path)
    current_file_format = file_format(file_path)
    image_label = is_image_file(file_path)
    return VerifyResult(
        file_name=str(row.get("file_name") or file_path.name),
        file_path=str(file_path),
        file_type=current_file_type,
        file_format=current_file_format,
        is_image_label=image_label,
        carrier="UNKNOWN",
        last_mile_carrier="UNKNOWN",
        template_code="",
        template_sub_code="",
        template_marker="",
        carrier_display="UNKNOWN",
        template_source="",
        template_confidence="",
        source_tracking=source_tracking_from_row(row, wms_records),
        filename_tracking="",
        pdf_text_tracking="",
        barcode_tracking="",
        download_order_no=str(row.get("delivery_no") or row.get("deliveryNo") or ""),
        download_wave_no=str(row.get("source_no") or row.get("sourceNo") or row.get("waveNo") or ""),
        download_warehouse=str(row.get("whCode") or row.get("wh_code") or ""),
        ocr_tracking="",
        all_text_candidates="",
        all_barcode_candidates="",
        verify_status="download_file_error",
        verify_status_zh=status_to_zh("download_file_error"),
        confidence="review",
        confidence_zh=confidence_to_zh("review"),
        reason=f"file processing failed: {error_text}",
        reason_zh=reason_to_zh("download_file_error", False, "", [], [], image_label, file_error=error_text),
        need_review=True,
        barcode_success=False,
        barcode_page_numbers="",
        barcode_processing_seconds=0,
        barcode_matches_source=False,
        barcode_matches_pdf_text=False,
        barcode_verify_note=f"条码未反读成功: {error_text}",
        barcode_error=error_text,
        pdf_text_success=False,
        image_barcode_result="",
        image_processing_note="图片文件异常" if image_label else "",
        ocr_needed=True,
        ocr_reason="下载文件异常，无法打开或格式不支持",
        ocr_priority="高",
        processing_seconds=0,
        decoded_formats="",
        decoded_raw_values="",
        page_count=None,
        recognized_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        meta_carrier=str(row.get("meta_carrier") or ""),
        meta_channel=str(row.get("meta_channel") or ""),
        meta_channel_code=str(row.get("meta_channel_code") or ""),
        meta_customer_code=str(row.get("meta_customer_code") or ""),
        meta_source=str(row.get("meta_source") or ""),
        meta_carrier_conflict=False,
        meta_carrier_note="",
    )


def process_row(row: dict[str, Any], args: argparse.Namespace, zxingcpp: Any | None, wms_records: dict[str, str]) -> VerifyResult:
    start_time = time.perf_counter()
    file_path = Path(str(row.get("file_path") or row.get("pdf_path") or ""))
    current_file_type = file_type(file_path)
    current_file_format = file_format(file_path)
    image_label = is_image_file(file_path)
    source_tracking = source_tracking_from_row(row, wms_records)
    filename_candidates = extract_candidates(file_path.stem, "filename")
    filename_tracking = filename_candidates[0].value if filename_candidates else normalize(str(row.get("filename_tracking_no") or ""))
    download_order_no = str(row.get("delivery_no") or row.get("deliveryNo") or "")
    download_wave_no = str(row.get("source_no") or row.get("sourceNo") or row.get("waveNo") or "")
    download_warehouse = str(row.get("whCode") or row.get("wh_code") or "")
    meta_carrier = str(row.get("meta_carrier") or "")
    meta_channel = str(row.get("meta_channel") or "")
    meta_channel_code = str(row.get("meta_channel_code") or "")
    meta_customer_code = str(row.get("meta_customer_code") or "")
    meta_source = str(row.get("meta_source") or "")
    carrier = infer_carrier(source_tracking)
    reason_parts: list[str] = []
    file_error = ""

    if not file_path.exists():
        file_error = "file does not exist"
    elif not is_supported_input(file_path):
        file_error = f"unsupported file format: {file_path.suffix}"

    text = ""
    page_count: int | None = None
    text_error = ""
    strict_text_values = parse_tracking_list(str(row.get("pdf_text_tracking_no") or ""))
    text_candidates: list[Candidate] = []
    text_values = strict_text_values
    if not file_error and current_file_type == "PDF":
        text, page_count, text_error = read_pdf_text(file_path, max_pages=args.max_pages)
        text_candidates = extract_candidates(text, "pdf_text", carrier_hint=carrier, source_tracking=source_tracking)
        text_values = unique(strict_text_values + values_from_candidates(text_candidates))
        if text_error:
            reason_parts.append(f"PDF text extraction error: {text_error}")
    elif image_label:
        page_count = 1
        reason_parts.append("下载文件为图片格式，已按图片面单处理")
    pdf_text_success = bool(text_values)

    barcode_start = time.perf_counter()
    if file_error:
        decoded_items, barcode_page_count, barcode_error = [], page_count, file_error
    else:
        decoded_items, barcode_page_count, barcode_error = decode_barcode_for_file(
            zxingcpp=zxingcpp,
            path=file_path,
            dpi=args.dpi,
            max_pages=args.max_pages,
            rotations=args.rotations,
            timeout=args.timeout,
        )
    barcode_processing_seconds = round(time.perf_counter() - barcode_start, 3)
    if barcode_page_count is not None:
        page_count = barcode_page_count
    if barcode_error.startswith("download_file_error:"):
        file_error = barcode_error.replace("download_file_error:", "", 1).strip()
    decoded_raw_values = [item["text"] for item in decoded_items]
    barcode_page_numbers = ",".join(unique([str(item.get("page") or "") for item in decoded_items]))
    barcode_candidates: list[Candidate] = []
    for decoded_value in decoded_raw_values:
        decoded_compact = normalize(decoded_value)
        if source_tracking and source_tracking in decoded_compact:
            barcode_candidates.append(
                Candidate(
                    carrier=carrier,
                    value=source_tracking,
                    source="barcode",
                    score=score_candidate(carrier, source_tracking, "barcode", decoded_compact, source_tracking),
                    context=decoded_compact[:180],
                )
            )
        barcode_candidates.extend(
            extract_candidates(
                decoded_compact,
                "barcode",
                carrier_hint=carrier,
                source_tracking=source_tracking,
            )
        )
    barcode_values = values_from_candidates(barcode_candidates)
    barcode_success = bool(decoded_items)
    if barcode_error and not barcode_success:
        reason_parts.append(barcode_error)
    barcode_matches_source = bool(source_tracking and source_tracking in barcode_values)
    barcode_matches_pdf_text = any(value in text_values for value in barcode_values)

    if carrier == "UNKNOWN":
        inferred = infer_carrier((source_tracking or (text_values[0] if text_values else "") or (barcode_values[0] if barcode_values else "")))
        carrier = inferred

    if file_error:
        verify_status, confidence, status_reason, need_review = (
            "download_file_error",
            "review",
            "download file error",
            True,
        )
    else:
        verify_status, confidence, status_reason, need_review = determine_status(
            source_tracking=source_tracking,
            text_values=text_values,
            barcode_values=barcode_values,
            carrier=carrier,
            barcode_error=barcode_error,
        )
    reason_parts.insert(0, status_reason)
    if verify_status == "source_only_low_confidence" and barcode_error and "timeout" in barcode_error.lower():
        verify_status = "barcode_failed"
        need_review = True

    ocr_needed, ocr_reason, ocr_priority = ocr_decision(
        status=verify_status,
        carrier=carrier,
        source_tracking=source_tracking,
        text_values=text_values,
        barcode_values=barcode_values,
        barcode_decoded=barcode_success,
        is_image_label=image_label,
        file_error=file_error,
    )
    if ocr_needed and verify_status == "review_unknown":
        verify_status = "ocr_needed"
        confidence = "review"
        need_review = True
    if ocr_needed and image_label and not barcode_values and verify_status == "source_only_low_confidence":
        confidence = "review"
        need_review = True

    last_mile_carrier = infer_carrier(
        (barcode_values[0] if barcode_values else "")
        or (text_values[0] if text_values else "")
        or source_tracking
        or filename_tracking
    )
    if last_mile_carrier == "UNKNOWN":
        last_mile_carrier = carrier
    if carrier == "UNKNOWN" and last_mile_carrier != "UNKNOWN":
        carrier = last_mile_carrier

    final_carrier_for_compare = last_mile_carrier if last_mile_carrier != "UNKNOWN" else carrier
    normalized_meta_carrier = normalize_carrier_name(meta_carrier)
    meta_carrier_conflict = bool(
        normalized_meta_carrier
        and final_carrier_for_compare != "UNKNOWN"
        and normalized_meta_carrier != final_carrier_for_compare
    )
    meta_carrier_note = ""
    if meta_carrier_conflict:
        meta_carrier_note = f"metadata={normalized_meta_carrier}; tracking={final_carrier_for_compare}"
    elif normalized_meta_carrier:
        meta_carrier_note = f"metadata={normalized_meta_carrier}"

    allow_full_page_template_ocr = verify_status not in TRACKING_STATUS_STRONG
    template_match = TemplateMatch()
    if not file_error:
        template_match = detect_template(
            path=file_path,
            pdf_text=text,
            decoded_raw_values=decoded_raw_values,
            dpi=args.dpi,
            is_image_label=image_label,
            ocr_enabled=bool(getattr(args, "ocr_enabled", getattr(args, "ocr", False))),
            allow_full_page_ocr=allow_full_page_template_ocr,
        )
    display_carrier = carrier_display(last_mile_carrier, template_match.template_code, template_match.template_sub_code)

    best_text = source_tracking if source_tracking in text_values else (choose_best(text_candidates).value if choose_best(text_candidates) else "")
    best_barcode = source_tracking if source_tracking in barcode_values else (choose_best(barcode_candidates).value if choose_best(barcode_candidates) else "")
    processing_seconds = round(time.perf_counter() - start_time, 3)
    reason_zh = reason_to_zh(
        verify_status,
        barcode_success,
        source_tracking,
        barcode_values,
        text_values,
        image_label,
        file_error=file_error,
    )
    barcode_verify_note = barcode_note(
        barcode_success=barcode_success,
        barcode_values=barcode_values,
        source_tracking=source_tracking,
        text_values=text_values,
        barcode_error=barcode_error,
    )
    image_processing_note = ""
    image_barcode_result = ""
    if image_label:
        image_barcode_result = ",".join(barcode_values) if barcode_values else " | ".join(decoded_raw_values)
        if barcode_values:
            image_processing_note = "图片面单条码识别成功，并提取到有效追踪号"
        elif barcode_success:
            image_processing_note = "图片面单条码反读成功，但未提取到有效追踪号，建议OCR深度识别"
        else:
            image_processing_note = "图片面单条码识别失败，建议OCR深度识别"

    return VerifyResult(
        file_name=str(row.get("file_name") or file_path.name),
        file_path=str(file_path),
        file_type=current_file_type,
        file_format=current_file_format,
        is_image_label=image_label,
        carrier=carrier,
        last_mile_carrier=last_mile_carrier,
        template_code=template_match.template_code,
        template_sub_code=template_match.template_sub_code,
        template_marker=template_match.template_marker,
        carrier_display=display_carrier,
        template_source=template_match.template_source,
        template_confidence=template_match.template_confidence,
        source_tracking=source_tracking,
        filename_tracking=filename_tracking,
        pdf_text_tracking=best_text,
        barcode_tracking=best_barcode,
        download_order_no=download_order_no,
        download_wave_no=download_wave_no,
        download_warehouse=download_warehouse,
        ocr_tracking="",
        all_text_candidates=",".join(text_values),
        all_barcode_candidates=",".join(barcode_values),
        verify_status=verify_status,
        verify_status_zh=status_to_zh(verify_status),
        confidence=confidence,
        confidence_zh=confidence_to_zh(confidence),
        reason="; ".join(unique([part for part in reason_parts if part])),
        reason_zh=reason_zh,
        need_review=need_review,
        barcode_success=barcode_success,
        barcode_page_numbers=barcode_page_numbers,
        barcode_processing_seconds=barcode_processing_seconds,
        barcode_matches_source=barcode_matches_source,
        barcode_matches_pdf_text=barcode_matches_pdf_text,
        barcode_verify_note=barcode_verify_note,
        barcode_error="" if barcode_success else barcode_error,
        pdf_text_success=pdf_text_success,
        image_barcode_result=image_barcode_result,
        image_processing_note=image_processing_note,
        ocr_needed=ocr_needed,
        ocr_reason=ocr_reason,
        ocr_priority=ocr_priority,
        processing_seconds=processing_seconds,
        decoded_formats=",".join(sorted({str(item.get("format") or "") for item in decoded_items if item.get("format")})),
        decoded_raw_values=" | ".join(decoded_raw_values),
        page_count=page_count,
        recognized_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        meta_carrier=meta_carrier,
        meta_channel=meta_channel,
        meta_channel_code=meta_channel_code,
        meta_customer_code=meta_customer_code,
        meta_source=meta_source,
        meta_carrier_conflict=meta_carrier_conflict,
        meta_carrier_note=meta_carrier_note,
    )


def write_excel(results: list[VerifyResult], path: Path) -> None:
    from exporter import write_excel as write_chinese_excel

    path.parent.mkdir(parents=True, exist_ok=True)
    write_chinese_excel(results, path)


def write_json(results: list[VerifyResult], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        path.write_text(json.dumps([asdict(result) for result in results], ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(f"Failed to write JSON: {path}; {exc}") from exc


def print_result(index: int, total: int, result: VerifyResult) -> None:
    print(f"[{index}/{total}] file={result.file_name}", flush=True)
    print(f"  carrier={result.carrier}", flush=True)
    print(f"  last_mile_carrier={result.last_mile_carrier}", flush=True)
    print(
        f"  template={result.template_code} sub_code={result.template_sub_code} marker={result.template_marker} "
        f"source={result.template_source} confidence={result.template_confidence}",
        flush=True,
    )
    print(f"  carrier_display={result.carrier_display}", flush=True)
    print(f"  source_tracking={result.source_tracking}", flush=True)
    print(f"  pdf_text_tracking={result.pdf_text_tracking}", flush=True)
    print(f"  barcode_tracking={result.barcode_tracking}", flush=True)
    print(f"  barcode_formats={result.decoded_formats}", flush=True)
    print(f"  barcode_pages={result.barcode_page_numbers}", flush=True)
    print(f"  barcode_seconds={result.barcode_processing_seconds}", flush=True)
    print(f"  barcode_note={result.barcode_verify_note}", flush=True)
    print(f"  status={result.verify_status}", flush=True)
    print(f"  status_zh={result.verify_status_zh}", flush=True)
    print(f"  reason={result.reason}", flush=True)
    print(f"  ocr_needed={result.ocr_needed} reason={result.ocr_reason}", flush=True)
    if result.verify_status == "barcode_failed" or (not result.barcode_success and result.barcode_error):
        print(f"[barcode_failed] file={result.file_name} reason={result.barcode_error or result.reason}", flush=True)
    if result.verify_status == "review_conflict":
        print(
            f"[review_conflict] file={result.file_name} "
            f"source={result.source_tracking} text={result.all_text_candidates} barcode={result.all_barcode_candidates}",
            flush=True,
        )


def print_summary(results: list[VerifyResult]) -> None:
    counts: dict[str, int] = {}
    for result in results:
        counts[result.verify_status] = counts.get(result.verify_status, 0) + 1
    total = len(results)
    strong = sum(counts.get(status, 0) for status in TRACKING_STATUS_STRONG)
    auto_pass = strong
    need_review = sum(1 for result in results if result.need_review)
    print(f"total={total}")
    for status in [
        "auto_pass_triple_verified",
        "auto_pass_barcode_verified",
        "auto_pass_text_verified",
        "auto_pass_verified_prefix",
        "source_only_low_confidence",
        "review_conflict",
        "review_unknown",
        "barcode_failed",
        "ocr_needed",
        "download_file_error",
    ]:
        print(f"{status}={counts.get(status, 0)}")
    print(f"strong_verified_rate={(strong / total * 100) if total else 0:.2f}%")
    print(f"auto_pass_rate={(auto_pass / total * 100) if total else 0:.2f}%")
    print(f"need_review_rate={(need_review / total * 100) if total else 0:.2f}%")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Strict PDF label tracking verification with barcode decoding.")
    parser.add_argument("--input-dir", default="", help="PDF folder or single PDF. Used when --strict-json is not provided.")
    parser.add_argument("--strict-json", default="", help="Previous strict tracking JSON to enrich with barcode verification.")
    parser.add_argument("--download-log", default="download_log.csv", help="Optional WMS CSV containing expressNo/filePath.")
    parser.add_argument("--output-dir", default=str(Path("output") / "pdf"))
    parser.add_argument("--output-name", default="barcode_tracking_verify")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--offset", type=int, default=0)
    parser.add_argument("--rotations", action="store_true")
    parser.add_argument("--dpi", type=int, default=200)
    parser.add_argument("--max-pages", type=int, default=1)
    parser.add_argument("--timeout", type=int, default=30)
    parser.add_argument("--ocr", action="store_true", help="Enable bottom/full-page OCR for weak records.")
    parser.add_argument("--only-source-only", action="store_true")
    parser.add_argument("--debug", action="store_true")
    return parser.parse_args(argv)


def load_rows(args: argparse.Namespace, wms_records: dict[str, str],
              meta_by_file: dict[str, dict[str, str]] | None = None,
              meta_by_dno: dict[str, dict[str, str]] | None = None) -> list[dict[str, Any]]:
    if args.strict_json:
        rows = read_json(Path(args.strict_json))
    elif args.input_dir:
        rows = rows_from_input_dir(Path(args.input_dir), wms_records, meta_by_file, meta_by_dno)
    else:
        raise ValueError("Provide --strict-json or --input-dir")

    if args.only_source_only:
        rows = [
            row
            for row in rows
            if str(row.get("verification_status") or row.get("verify_status") or "") in SOURCE_ONLY_STATUSES
        ]
    if args.offset > 0:
        rows = rows[args.offset :]
    if args.limit > 0:
        rows = rows[: args.limit]
    return rows


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    args.ocr_enabled = bool(args.ocr)
    zxingcpp = load_zxingcpp()
    wms_records = load_wms_tracking(Path(args.download_log))
    try:
        rows = load_rows(args, wms_records)
    except Exception as exc:
        print(f"Failed to load inputs: {exc}", file=sys.stderr)
        return 2
    if not rows:
        print("No rows to process.")
        return 0

    results: list[VerifyResult] = []
    total = len(rows)
    for index, row in enumerate(rows, start=1):
        try:
            result = process_row(row, args, zxingcpp, wms_records)
        except Exception as exc:
            result = error_result(row, wms_records, exc)
        results.append(result)
        print_result(index, total, result)

    output_dir = Path(args.output_dir)
    excel_path = output_dir / f"{args.output_name}.xlsx"
    json_path = output_dir / f"{args.output_name}.json"
    try:
        write_excel(results, excel_path)
        write_json(results, json_path)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 3
    print_summary(results)
    print(f"excel={excel_path.resolve()}")
    print(f"json={json_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
