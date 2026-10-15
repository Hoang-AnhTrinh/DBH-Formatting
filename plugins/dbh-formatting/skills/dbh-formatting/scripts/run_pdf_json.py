"""Create page-oriented JSON for a PDF with selective rendering.

The script is intentionally conservative: it never infers a new drawing name
when the page policy does not identify one. Unknown pages are marked for review.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path

from pypdf import PdfReader


def estimate_tokens(text: str) -> int:
    # A transparent rough estimate for budgeting only; not an API billing value.
    return max(0, (len(text) + 3) // 4)


def text_quality(text: str) -> str:
    if not text.strip():
        return "empty"
    replacement = text.count("\ufffd")
    control = sum(1 for char in text if ord(char) < 32 and char not in "\n\r\t")
    ratio = (replacement + control) / max(len(text), 1)
    if ratio > 0.01 or len(text.strip()) < 20:
        return "low"
    return "usable"


def normalize_name(value: str) -> str:
    return re.sub(r"\s+", "", value or "").replace("（", "(").replace("）", ")")


def load_policy(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def policy_by_page(policy: dict) -> dict[int, dict]:
    result = {}
    for index, item in enumerate(policy.get("pages", []), start=1):
        page_number = int(item.get("page_number", index))
        result[page_number] = item
    return result


def find_pdftoppm(explicit: str | None) -> str | None:
    candidates = [explicit, shutil.which("pdftoppm")]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return candidate
    return None


def render_page(pdftoppm: str, pdf_path: Path, page_number: int, dpi: int, output_path: Path) -> tuple[str, int]:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prefix = output_path.with_suffix("")
    command = [
        pdftoppm,
        "-f", str(page_number),
        "-l", str(page_number),
        "-r", str(dpi),
        "-png",
        "-singlefile",
        str(pdf_path),
        str(prefix),
    ]
    completed = subprocess.run(command, capture_output=True, text=True, check=False)
    actual = prefix.with_suffix(".png")
    if completed.returncode != 0 or not actual.exists():
        return "failed", 0
    return "created", actual.stat().st_size


def page_content_present(page) -> bool:
    try:
        return page.get_contents() is not None
    except Exception:
        return False


def build(args: argparse.Namespace) -> dict:
    started = time.perf_counter()
    pdf_path = Path(args.pdf).resolve()
    output_dir = Path(args.output).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    pages_dir = output_dir / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)

    policy = load_policy(Path(args.policy).resolve()) if args.policy else {"pages": []}
    by_page = policy_by_page(policy)
    reader = PdfReader(str(pdf_path))
    page_count = len(reader.pages)
    expected_count = len(policy.get("pages", []))
    warnings: list[str] = []
    if expected_count and expected_count != page_count:
        warnings.append(f"page_count_mismatch: policy={expected_count}, pdf={page_count}")
    policy_source = policy.get("source_document")
    if policy_source and policy_source != pdf_path.name:
        warnings.append(f"policy_source_mismatch: policy={policy_source}, pdf={pdf_path.name}")
    policy_mismatch = bool(warnings)

    pdftoppm = find_pdftoppm(args.pdftoppm)
    pages = []
    search_records = []
    deep_count = 0
    image_bytes = 0
    native_chars = 0
    estimated_tokens = 0

    for page_number, page in enumerate(reader.pages, start=1):
        page_started = time.perf_counter()
        item = by_page.get(page_number)
        known = item is not None
        mode = item.get("mode", args.default_mode) if item else args.default_mode
        drawing_no = item.get("drawing_no") if item else None
        drawing_name = item.get("drawing_name") if item else None
        text = ""
        extraction_status = "not_requested"
        if mode == "deep_processing":
            deep_count += 1
            try:
                text = page.extract_text() or ""
                extraction_status = "extracted"
            except Exception as exc:
                extraction_status = f"failed: {type(exc).__name__}: {exc}"
        quality = text_quality(text)
        native_chars += len(text)
        estimated_tokens += estimate_tokens(text)
        page_record = {
            "page_number": page_number,
            "drawing_no": drawing_no,
            "drawing_name": drawing_name,
            "mode": mode,
            "needs_review": (not known) or policy_mismatch,
            "source": {
                "file": str(pdf_path),
                "file_name": pdf_path.name,
                "page_size_pt": {"width": float(page.mediabox.width), "height": float(page.mediabox.height)},
                "rotation": int(page.get("/Rotate", 0) or 0),
                "content_stream_present": page_content_present(page),
            },
            "extraction": {
                "native_text": {
                    "status": extraction_status,
                    "quality": quality,
                    "characters": len(text),
                    "estimated_input_tokens": estimate_tokens(text),
                    "text": text if mode == "deep_processing" else None,
                },
                "vector_or_native_content": {
                    "status": "metadata_only" if mode != "deep_processing" else "preserved_as_pdf_source_reference",
                    "source_page_reference": f"{pdf_path.name}#page={page_number}",
                },
                "image": {
                    "status": "not_requested",
                    "path": None,
                    "dpi": None,
                    "bytes": 0,
                },
                "ocr": {
                    "status": "selective_regions_only",
                    "regions": item.get("ocr_regions", []) if item else [],
                },
            },
            "records": [],
        }
        review_reasons = []
        if not known:
            review_reasons.append("drawing_name_or_page_not_in_policy")
        if policy_mismatch:
            review_reasons.append("policy_does_not_match_source_document")
        if review_reasons:
            page_record["review_reason"] = review_reasons
        if mode == "deep_processing":
            if pdftoppm:
                relative_image = Path("pages") / f"page-{page_number:03d}.png"
                image_status, bytes_written = render_page(
                    pdftoppm, pdf_path, page_number, args.deep_dpi, output_dir / relative_image
                )
                page_record["extraction"]["image"] = {
                    "status": image_status,
                    "path": relative_image.as_posix() if image_status == "created" else None,
                    "dpi": args.deep_dpi if image_status == "created" else None,
                    "bytes": bytes_written,
                }
                image_bytes += bytes_written
            else:
                page_record["extraction"]["image"]["status"] = "pdftoppm_not_found"
            if quality == "low":
                page_record["needs_review"] = True
                existing = page_record.get("review_reason", [])
                if isinstance(existing, str):
                    existing = [existing]
                page_record["review_reason"] = existing + ["needs_visual_or_ocr_review"]
        page_record["timing_ms"] = round((time.perf_counter() - page_started) * 1000, 1)
        page_json = output_dir / "pages" / f"page-{page_number:03d}.json"
        page_json.write_text(json.dumps(page_record, ensure_ascii=False, indent=2), encoding="utf-8")
        pages.append(page_record)
        if mode == "deep_processing":
            search_records.append({
                "page_number": page_number,
                "drawing_no": drawing_no,
                "drawing_name": drawing_name,
                "text": text,
                "quality": quality,
            })

    manifest = {
        "schema": "dbh-formatting/page-manifest@0.1",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "source": {"file": str(pdf_path), "file_name": pdf_path.name, "size_bytes": pdf_path.stat().st_size},
        "policy": {"file": str(Path(args.policy).resolve()) if args.policy else None, "version": policy.get("version")},
        "summary": {
            "page_count": page_count,
            "policy_page_count": expected_count,
            "deep_processing_pages": deep_count,
            "metadata_only_pages": page_count - deep_count,
            "native_text_characters": native_chars,
            "estimated_input_tokens": estimated_tokens,
            "rendered_image_bytes": image_bytes,
            "warnings": warnings,
        },
        "timing_ms": round((time.perf_counter() - started) * 1000, 1),
        "pages": pages,
    }
    (output_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (output_dir / "search_index.json").write_text(json.dumps({"schema": "dbh-formatting/search-index@0.1", "records": search_records}, ensure_ascii=False, indent=2), encoding="utf-8")
    review_queue = {
        "schema": "dbh-formatting/review-queue@0.1",
        "source": str(pdf_path),
        "instructions": "Confirm drawing_no, drawing_name and mode for each listed page before deep extraction on a new hồ sơ.",
        "pages": [
            {
                "page_number": page["page_number"],
                "drawing_no": page.get("drawing_no"),
                "drawing_name": page.get("drawing_name"),
                "mode": page["mode"],
                "reasons": page.get("review_reason", []),
            }
            for page in pages
            if page.get("needs_review")
        ],
    }
    (output_dir / "review_queue.json").write_text(json.dumps(review_queue, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--policy", required=False)
    parser.add_argument("--output", required=True)
    parser.add_argument("--pdftoppm", required=False)
    parser.add_argument("--deep-dpi", type=int, default=150)
    parser.add_argument("--default-mode", choices=["deep_processing", "metadata_only"], default="metadata_only")
    args = parser.parse_args()
    manifest = build(args)
    print(json.dumps({"source": manifest["source"], "summary": manifest["summary"], "timing_ms": manifest["timing_ms"]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

