"""Parse one PDF containing exactly two ordered final-week schedule blocks."""
from __future__ import annotations

import re
import tempfile
from dataclasses import dataclass
from pathlib import Path

import fitz
import pdfplumber

from src.modules.ocr_runtime import run_tesseract_tsv
from src.modules.pdf_schedule_parser import (
    WEEKDAYS,
    _canonical_weekday,
    _match_roster_names,
    _unknown_candidates,
)


WEEK_HEADING_RE = re.compile(r"^(?:第\s*)?(\d+)\s*周$")
NUMBER_RE = re.compile(r"^\d+$")


@dataclass(frozen=True)
class FinalWeekBlock:
    """One ordered week section inside a combined final-weeks PDF."""

    slot: int
    week_number: int
    schedule: dict[str, list[str]]
    matched_count: int
    unknown_names: list[str]


def parse_final_weeks_pdf(
    pdf_path: str,
    total_names: list[str],
    corrections: dict[str, str] | None = None,
) -> list[FinalWeekBlock]:
    """Return the two final-week blocks in visual order."""
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {pdf_path}")
    if path.suffix.lower() != ".pdf":
        raise ValueError("期末周排班文件仅支持 PDF。")

    native_error = None
    try:
        blocks = _analyze_native_pdf(path, total_names, corrections)
        _validate_two_blocks(blocks)
        return blocks
    except ValueError as exc:
        native_error = exc

    try:
        blocks = _analyze_scanned_pdf(path, total_names, corrections)
        _validate_two_blocks(blocks)
        return blocks
    except ValueError as ocr_error:
        if str(native_error).startswith("PDF 文件无法读取或已损坏"):
            raise native_error
        raise ValueError(f"期末周排班 PDF 解析失败：{native_error}；OCR 回退失败：{ocr_error}") from ocr_error


def scan_final_weeks_pdf(
    pdf_path: str,
    total_names: list[str],
    corrections: dict[str, str] | None = None,
) -> list[FinalWeekBlock]:
    """Scan the combined PDF while retaining per-slot warning ownership."""
    return parse_final_weeks_pdf(pdf_path, total_names, corrections)


def _analyze_native_pdf(
    path: Path,
    total_names: list[str],
    corrections: dict[str, str] | None,
) -> list[FinalWeekBlock]:
    positioned_pages = []
    try:
        with pdfplumber.open(path) as pdf:
            for page_index, page in enumerate(pdf.pages):
                words = page.extract_words(
                    x_tolerance=2,
                    y_tolerance=2,
                    keep_blank_chars=False,
                ) or []
                positioned_pages.append((page_index, words))
    except (OSError, ValueError) as exc:
        raise ValueError(f"PDF 文件无法读取或已损坏: {exc}") from exc
    except Exception as exc:
        raise ValueError(f"PDF 原生文字解析失败: {exc}") from exc
    return _analyze_positioned_pages(positioned_pages, total_names, corrections)


def _analyze_scanned_pdf(
    path: Path,
    total_names: list[str],
    corrections: dict[str, str] | None,
) -> list[FinalWeekBlock]:
    positioned_pages = []
    try:
        document = fitz.open(path)
    except Exception as exc:
        raise ValueError(f"PDF 文件无法读取或已损坏: {exc}") from exc

    try:
        with tempfile.TemporaryDirectory(prefix="tonban-final-weeks-ocr-") as temp_dir:
            for page_index, page in enumerate(document):
                image_path = Path(temp_dir) / f"page-{page_index + 1}.png"
                pixmap = page.get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72), alpha=False)
                pixmap.save(image_path)
                words = []
                for token in run_tesseract_tsv(image_path):
                    left = float(token.get("left", 0))
                    top = float(token.get("top", 0))
                    width = float(token.get("width", 0))
                    height = float(token.get("height", 0))
                    words.append({
                        "text": str(token.get("text", "")),
                        "x0": left,
                        "x1": left + width,
                        "top": top,
                        "bottom": top + height,
                    })
                positioned_pages.append((page_index, words))
    finally:
        document.close()
    return _analyze_positioned_pages(positioned_pages, total_names, corrections)


def _analyze_positioned_pages(
    positioned_pages: list[tuple[int, list[dict]]],
    total_names: list[str],
    corrections: dict[str, str] | None,
) -> list[FinalWeekBlock]:
    headings = []
    for page_index, words in positioned_pages:
        for heading in _find_week_headings(words):
            headings.append({"page": page_index, **heading})
    headings.sort(key=lambda item: (item["page"], item["top"], item["x0"]))

    blocks = []
    for index, heading in enumerate(headings):
        page_index = heading["page"]
        page_words = positioned_pages[page_index][1]
        next_top = None
        for later in headings[index + 1:]:
            if later["page"] == page_index:
                next_top = later["top"]
                break
            if later["page"] > page_index:
                break
        block_words = [
            word for word in page_words
            if float(word.get("top", 0)) >= heading["top"] - 3
            and (next_top is None or float(word.get("top", 0)) < next_top - 3)
        ]
        note_tops = [
            float(word.get("top", 0))
            for word in block_words
            if str(word.get("text", "")).strip().startswith("备注")
        ]
        if note_tops:
            notes_top = min(note_tops)
            block_words = [word for word in block_words if float(word.get("top", 0)) < notes_top]
        schedule, unknown_names = _parse_block_words(block_words, total_names, corrections)
        blocks.append(FinalWeekBlock(
            slot=len(blocks) + 1,
            week_number=heading["week_number"],
            schedule=schedule,
            matched_count=sum(len(names) for names in schedule.values()),
            unknown_names=unknown_names,
        ))
    return blocks


def _find_week_headings(words: list[dict]) -> list[dict]:
    headings = []
    used_ids = set()
    for word in words:
        text = re.sub(r"\s+", "", str(word.get("text", "")))
        direct = WEEK_HEADING_RE.fullmatch(text)
        if direct:
            headings.append({
                "week_number": int(direct.group(1)),
                "top": float(word.get("top", 0)),
                "x0": float(word.get("x0", 0)),
            })
            used_ids.add(id(word))
            continue
        if not NUMBER_RE.fullmatch(text):
            continue
        candidates = []
        for following in words:
            if id(following) in used_ids:
                continue
            if re.sub(r"\s+", "", str(following.get("text", ""))) != "周":
                continue
            same_line = abs(float(word.get("top", 0)) - float(following.get("top", 0))) <= 4
            gap = float(following.get("x0", 0)) - float(word.get("x1", word.get("x0", 0)))
            if same_line and -2 <= gap <= 18:
                candidates.append((abs(gap), following))
        if candidates:
            following = min(candidates, key=lambda item: item[0])[1]
            headings.append({
                "week_number": int(text),
                "top": min(float(word.get("top", 0)), float(following.get("top", 0))),
                "x0": float(word.get("x0", 0)),
            })
            used_ids.update({id(word), id(following)})
    return sorted(headings, key=lambda item: (item["top"], item["x0"]))


def _parse_block_words(
    words: list[dict],
    total_names: list[str],
    corrections: dict[str, str] | None,
) -> tuple[dict[str, list[str]], list[str]]:
    schedule = {day: [] for day in WEEKDAYS}
    headers = []
    for word in words:
        day = _canonical_weekday(word.get("text"))
        if day is not None:
            headers.append((day, word))
    if not headers:
        return schedule, []

    header_top = min(float(word.get("top", 0)) for _, word in headers)
    header_bottom = max(float(word.get("bottom", word.get("top", 0))) for _, word in headers)
    header_centers = {
        day: (float(word.get("x0", 0)) + float(word.get("x1", word.get("x0", 0)))) / 2
        for day, word in headers
        if abs(float(word.get("top", 0)) - header_top) <= 6
    }
    if not header_centers:
        return schedule, []

    ordered_headers = sorted(header_centers.items(), key=lambda item: item[1])
    centers = [center for _, center in ordered_headers]
    gaps = [right - left for left, right in zip(centers, centers[1:]) if right > left]
    typical_gap = sorted(gaps)[len(gaps) // 2] if gaps else 120.0
    left_bound = centers[0] - typical_gap * 0.55
    right_bound = centers[-1] + typical_gap * 0.55

    body_words = []
    for word in words:
        top = float(word.get("top", 0))
        center = (float(word.get("x0", 0)) + float(word.get("x1", word.get("x0", 0)))) / 2
        if top <= header_bottom + 2 or center < left_bound or center > right_bound:
            continue
        day = min(header_centers, key=lambda item: abs(header_centers[item] - center))
        body_words.append((top, float(word.get("x0", 0)), day, str(word.get("text", ""))))

    lines: list[dict] = []
    for top, x0, day, text in sorted(body_words):
        target = None
        for line in lines:
            if abs(top - line["top"]) <= 4:
                target = line
                break
        if target is None:
            target = {"top": top, "items": []}
            lines.append(target)
        target["items"].append((x0, day, text))

    unknown_names = []
    seen_unknown = set()
    for line in lines:
        per_day = {}
        for x0, day, text in sorted(line["items"]):
            per_day.setdefault(day, []).append(text)
        for day, texts in per_day.items():
            line_text = " ".join(texts)
            schedule[day].extend(_match_roster_names(line_text, total_names, corrections))
            for candidate in _unknown_candidates(line_text, total_names, corrections):
                if candidate not in seen_unknown:
                    seen_unknown.add(candidate)
                    unknown_names.append(candidate)
    return schedule, unknown_names


def _validate_two_blocks(blocks: list[FinalWeekBlock]) -> None:
    if len(blocks) != 2:
        weeks = "、".join(f"第{block.week_number}周" for block in blocks) or "无"
        raise ValueError(
            f"期末周排班 PDF 必须恰好包含两个周次区块，当前识别到 {len(blocks)} 个（{weeks}）。"
        )
    if blocks[0].week_number == blocks[1].week_number:
        raise ValueError("期末周1和期末周2识别到相同周次，请检查 PDF。")
    for block in blocks:
        if block.week_number <= 0:
            raise ValueError(f"期末周{block.slot}的周次必须是正整数。")
        if block.matched_count <= 0:
            raise ValueError(
                f"期末周{block.slot}（第{block.week_number}周）未识别到任何总名单内姓名。"
            )
