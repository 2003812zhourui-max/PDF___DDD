from __future__ import annotations

import argparse
import json
from collections import defaultdict
from dataclasses import asdict, dataclass
from pathlib import Path

from barcode_verify_tracking import (
    SUPPORTED_INPUT_SUFFIXES,
    detect_template,
    is_image_file,
    read_pdf_text,
)
from config import DEFAULT_DPI, DEFAULT_MAX_PAGES


SCORED_LABELS = {"0024-01", "0024-02", "CBT", "CBS"}


@dataclass
class SampleEvalResult:
    file_name: str
    file_path: str
    expected: str
    actual: str
    template_code: str
    template_sub_code: str
    template_source: str
    template_confidence: str
    passed: bool
    scored: bool
    note: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Evaluate label template recognition on samples/* folders.")
    parser.add_argument("--samples-dir", default="samples", help="Sample root folder. Default: samples")
    parser.add_argument("--output", default="output/evaluation/sample_eval.json", help="JSON report path")
    parser.add_argument("--dpi", type=int, default=DEFAULT_DPI, help="PDF render DPI for OCR")
    parser.add_argument("--max-pages", type=int, default=DEFAULT_MAX_PAGES, help="PDF text pages to read")
    parser.add_argument("--ocr", action="store_true", help="Enable full template OCR fallback")
    parser.add_argument("--include-unknown", action="store_true", help="Score unknown folder as expecting no template")
    parser.add_argument("--fail-on-mismatch", action="store_true", help="Exit with code 1 if scored samples mismatch")
    return parser.parse_args()


def expected_matches(expected: str, actual: str, include_unknown: bool) -> tuple[bool, bool, str]:
    expected = expected.upper()
    actual = actual.upper()

    if expected in {"0024-01", "0024-02"}:
        return actual == expected, True, "0024 samples require exact 01/02 subdivision"
    if expected in {"CBT", "CBS"}:
        return actual == expected, True, "normal labels only require template type"
    if expected == "UNKNOWN" and include_unknown:
        return actual == "", True, "unknown samples expect no template match"
    return False, False, "folder is not scored"


def display_actual(template_code: str, template_sub_code: str, template_marker: str) -> str:
    if template_code == "0024":
        return template_marker if template_sub_code in {"01", "02"} else "0024"
    return template_code


def list_sample_files(samples_dir: Path) -> list[tuple[str, Path]]:
    files: list[tuple[str, Path]] = []
    for folder in sorted(path for path in samples_dir.iterdir() if path.is_dir()):
        expected = folder.name.upper()
        for path in sorted(folder.rglob("*")):
            if path.is_file() and path.suffix.lower() in SUPPORTED_INPUT_SUFFIXES:
                files.append((expected, path))
    return files


def evaluate_file(expected: str, path: Path, args: argparse.Namespace) -> SampleEvalResult:
    text = ""
    if path.suffix.lower() == ".pdf":
        text, _page_count, _error = read_pdf_text(path, max_pages=args.max_pages)

    match = detect_template(
        path=path,
        pdf_text=text,
        decoded_raw_values=[],
        dpi=args.dpi,
        is_image_label=is_image_file(path),
        ocr_enabled=bool(args.ocr),
        allow_full_page_ocr=True,
    )
    actual = display_actual(match.template_code, match.template_sub_code, match.template_marker)
    passed, scored, note = expected_matches(expected, actual, include_unknown=bool(args.include_unknown))
    return SampleEvalResult(
        file_name=path.name,
        file_path=str(path.resolve()),
        expected=expected,
        actual=actual,
        template_code=match.template_code,
        template_sub_code=match.template_sub_code,
        template_source=match.template_source,
        template_confidence=match.template_confidence,
        passed=passed,
        scored=scored,
        note=note,
    )


def print_summary(results: list[SampleEvalResult]) -> None:
    grouped: dict[str, list[SampleEvalResult]] = defaultdict(list)
    for result in results:
        grouped[result.expected].append(result)

    total_scored = sum(1 for result in results if result.scored)
    total_passed = sum(1 for result in results if result.scored and result.passed)
    rate = (total_passed / total_scored * 100) if total_scored else 0.0

    print("\nSample evaluation")
    print(f"Scored: {total_passed}/{total_scored} ({rate:.2f}%)")
    print("")
    print(f"{'Expected':<12} {'Pass':>7} {'Total':>7} {'Rate':>8}")
    print("-" * 38)
    for expected in sorted(grouped):
        scored = [result for result in grouped[expected] if result.scored]
        if not scored:
            print(f"{expected:<12} {'-':>7} {len(grouped[expected]):>7} {'skip':>8}")
            continue
        passed = sum(1 for result in scored if result.passed)
        current_rate = passed / len(scored) * 100
        print(f"{expected:<12} {passed:>7} {len(scored):>7} {current_rate:>7.2f}%")

    failures = [result for result in results if result.scored and not result.passed]
    if failures:
        print("\nMismatches")
        for result in failures:
            actual = result.actual or "EMPTY"
            print(f"- {result.expected} expected, got {actual}: {result.file_name}")


def main() -> int:
    args = parse_args()
    samples_dir = Path(args.samples_dir).expanduser().resolve()
    if not samples_dir.exists():
        raise SystemExit(f"Samples directory does not exist: {samples_dir}")

    results = [evaluate_file(expected, path, args) for expected, path in list_sample_files(samples_dir)]
    print_summary(results)

    output_path = Path(args.output).expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "samples_dir": str(samples_dir),
        "results": [asdict(result) for result in results],
    }
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nReport: {output_path}")

    has_mismatch = any(result.scored and not result.passed for result in results)
    return 1 if args.fail_on_mismatch and has_mismatch else 0


if __name__ == "__main__":
    raise SystemExit(main())
