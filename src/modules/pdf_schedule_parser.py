"""Parse weekly schedule tables from native-text and scanned PDF files."""
from __future__ import annotations

import re
import tempfile
from io import BytesIO
from pathlib import Path

import fitz
import numpy as np
import pdfplumber
from PIL import Image

from src.modules.ocr_runtime import run_tesseract_tsv
from src.modules.word_parser import (
    NAME_REGEX,
    WEEKDAY_PATTERNS,
    _compact_text,
    _is_weekday_annotation,
    _match_roster_names,
    _unknown_candidates,
)
from src.utils.helpers import edit_distance


WEEKDAYS = tuple(WEEKDAY_PATTERNS)


def parse_pdf_schedule(pdf_path: str, total_names: list[str], corrections: dict | None = None) -> dict[str, list[str]]:
    """Parse a weekly schedule PDF, preserving repeated name occurrences."""
    analysis = _analyze_pdf(pdf_path, total_names, corrections)
    if analysis["matched_count"] <= 0:
        raise ValueError("PDF 排班表中未识别到任何总名单内姓名，请检查表格或扫描清晰度。")
    return analysis["schedule"]


def scan_pdf_schedule(pdf_path: str, total_names: list[str], corrections: dict | None = None) -> dict:
    """Scan a weekly schedule PDF for matched and unknown names."""
    analysis = _analyze_pdf(pdf_path, total_names, corrections)
    return {
        "matched_count": analysis["matched_count"],
        "unknown_names": analysis["unknown_names"],
    }


def _analyze_pdf(pdf_path: str, total_names: list[str], corrections: dict | None = None) -> dict:
    try:
        native = _analyze_native_pdf(pdf_path, total_names, corrections)
        if native["matched_count"] > 0:
            return native
    except ValueError as exc:
        if str(exc).startswith("PDF 文件无法读取或已损坏"):
            raise
    return _analyze_scanned_pdf(pdf_path, total_names, corrections)


def _analyze_native_pdf(pdf_path: str, total_names: list[str], corrections: dict | None = None) -> dict:
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {pdf_path}")

    schedule = {day: [] for day in WEEKDAYS}
    unknown_names: list[str] = []
    seen_unknown: set[str] = set()
    found_weekday_table = False

    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables() or []:
                    header_index, columns = _find_weekday_header(table)
                    if header_index is None:
                        continue
                    found_weekday_table = True
                    for row in table[header_index + 1:]:
                        if _is_notes_row(row):
                            break
                        for day, column in columns.items():
                            if column >= len(row):
                                continue
                            cell_text = row[column] or ""
                            schedule[day].extend(_match_roster_names(cell_text, total_names, corrections))
                            for candidate in _unknown_candidates(cell_text, total_names, corrections):
                                if candidate not in seen_unknown:
                                    seen_unknown.add(candidate)
                                    unknown_names.append(candidate)
    except (OSError, ValueError) as exc:
        raise ValueError(f"PDF 文件无法读取或已损坏: {exc}") from exc
    except Exception as exc:
        raise ValueError(f"PDF 排班表解析失败: {exc}") from exc

    if not found_weekday_table:
        raise ValueError("PDF 中未找到包含周一至周五列的排班表。")

    return {
        "schedule": schedule,
        "matched_count": sum(len(names) for names in schedule.values()),
        "unknown_names": unknown_names,
    }


def _analyze_scanned_pdf(pdf_path: str, total_names: list[str], corrections: dict | None = None) -> dict:
    schedule = {day: [] for day in WEEKDAYS}
    unknown_names: list[str] = []
    seen_unknown: set[str] = set()
    parsed_pages = 0
    last_error = None
    try:
        document = fitz.open(pdf_path)
    except Exception as exc:
        raise ValueError(f"PDF 文件无法读取或已损坏: {exc}") from exc

    try:
        with tempfile.TemporaryDirectory(prefix="tonban-ocr-") as temp_dir:
            for page_index, page in enumerate(document):
                image_path = Path(temp_dir) / f"page-{page_index + 1}.png"
                pixmap = page.get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72), alpha=False)
                try:
                    page_analysis = _analyze_grid_ocr(
                        pixmap,
                        Path(temp_dir),
                        page_index,
                        total_names,
                        corrections,
                    )
                    if page_analysis["matched_count"] == 0:
                        synthetic_headers = _prepare_ocr_page(pixmap, image_path)
                        tokens = run_tesseract_tsv(image_path) + synthetic_headers
                        page_analysis = _analyze_ocr_tokens(tokens, total_names, corrections)
                except ValueError as exc:
                    last_error = exc
                    continue
                parsed_pages += 1
                for day in WEEKDAYS:
                    schedule[day].extend(page_analysis["schedule"][day])
                for candidate in page_analysis["unknown_names"]:
                    if candidate not in seen_unknown:
                        seen_unknown.add(candidate)
                        unknown_names.append(candidate)
    finally:
        document.close()

    matched_count = sum(len(names) for names in schedule.values())
    if parsed_pages == 0:
        raise ValueError(str(last_error or "OCR 未能定位排班表中的星期列，请使用更清晰的扫描件。"))
    if matched_count == 0:
        raise ValueError("扫描 PDF 完成 OCR 后仍未识别到任何总名单内姓名，请检查扫描清晰度和总名单。")
    return {
        "schedule": schedule,
        "matched_count": matched_count,
        "unknown_names": unknown_names,
    }


def _analyze_grid_ocr(pixmap, temp_dir: Path, page_index: int, total_names: list[str], corrections: dict | None = None) -> dict:
    grid = _detect_table_grid(pixmap)
    schedule = {day: [] for day in WEEKDAYS}
    unknown_names: list[str] = []
    seen_unknown: set[str] = set()
    image = grid["image"]
    vertical_lines = grid["vertical_lines"]
    for row_index, (top, bottom) in enumerate(grid["body_intervals"]):
        for day_index, day in enumerate(("周一", "周二", "周三", "周四", "周五")):
            left = vertical_lines[day_index + 1]
            right = vertical_lines[day_index + 2]
            padding = max(5, round(min(right - left, bottom - top) * 0.025))
            crop = image.crop((left + padding, top + padding, right - padding, bottom - padding))
            stem = f"page-{page_index + 1}-row-{row_index + 1}-day-{day_index + 1}"
            for text in _ocr_cell_lines(crop, temp_dir, stem, total_names, corrections):
                schedule[day].extend(_match_ocr_roster_names(text, total_names, corrections))
                for candidate in _unknown_candidates(text, total_names, corrections):
                    if candidate not in seen_unknown:
                        seen_unknown.add(candidate)
                        unknown_names.append(candidate)
    return {
        "schedule": schedule,
        "matched_count": sum(len(names) for names in schedule.values()),
        "unknown_names": unknown_names,
    }


def _ocr_cell_lines(
    crop: Image.Image,
    temp_dir: Path,
    stem: str,
    total_names: list[str],
    corrections: dict | None = None,
) -> list[str]:
    grayscale = crop.convert("L")
    pixels = np.array(grayscale)
    dark = pixels < 180
    row_indexes = np.where(dark.sum(axis=1) > max(4, round(pixels.shape[1] * 0.01)))[0]
    groups = []
    max_gap = max(2, round(pixels.shape[0] * 0.015))
    for value in (int(item) for item in row_indexes):
        if not groups or value > groups[-1][-1] + max_gap:
            groups.append([value])
        else:
            groups[-1].append(value)
    minimum_height = max(8, round(pixels.shape[0] * 0.04))
    lines = []
    for line_index, group in enumerate(groups):
        if group[-1] - group[0] + 1 < minimum_height:
            continue
        top = max(0, group[0] - 4)
        bottom = min(grayscale.height, group[-1] + 5)
        line_path = temp_dir / f"{stem}-line-{line_index + 1}.png"
        grayscale.crop((0, top, grayscale.width, bottom)).save(line_path)
        alternatives = []
        for psm in (7, 8):
            tokens = run_tesseract_tsv(line_path, language="chi_sim", psm=psm)
            text = "".join(str(token.get("text", "")) for token in tokens)
            if text.strip():
                alternatives.append(text)
        if alternatives:
            lines.append(max(alternatives, key=lambda text: len(_match_ocr_roster_names(text, total_names, corrections))))
    return lines


def _detect_table_grid(pixmap) -> dict:
    image = Image.open(BytesIO(pixmap.tobytes("png"))).convert("L")
    pixels = np.array(image)
    dark = pixels < 160
    horizontal_indexes = np.where(dark.sum(axis=1) > pixels.shape[1] * 0.6)[0]
    vertical_indexes = np.where(dark.sum(axis=0) > pixels.shape[0] * 0.25)[0]
    horizontal_lines = _group_index_centers(horizontal_indexes)
    vertical_lines = _select_seven_grid_lines(_group_index_centers(vertical_indexes))
    if len(vertical_lines) != 7 or len(horizontal_lines) < 3:
        raise ValueError("OCR 未检测到完整的六列表格网格。")

    line_presence = np.zeros(pixels.shape[0], dtype=int)
    for x in vertical_lines:
        line_presence += dark[:, max(0, x - 2):min(dark.shape[1], x + 3)].any(axis=1)
    grid_extent = _longest_index_run(
        np.where(line_presence >= 5)[0],
        max_gap=max(2, round(pixels.shape[0] * 0.05)),
    )
    if grid_extent is None:
        raise ValueError("OCR 未检测到连续的表格列边界。")
    grid_top, grid_bottom = grid_extent
    relevant_horizontal = [
        y for y in horizontal_lines if grid_top - 8 <= y <= grid_bottom + 8
    ]
    if len(relevant_horizontal) < 3:
        raise ValueError("OCR 未检测到完整的表格行边界。")
    header_top = min(relevant_horizontal, key=lambda y: abs(y - grid_top))
    after_header_top = [y for y in relevant_horizontal if y > header_top]
    if not after_header_top:
        raise ValueError("OCR 未检测到星期表头下边界。")
    header_bottom = after_header_top[0]
    row_boundaries = [y for y in relevant_horizontal if y >= header_bottom and y <= grid_bottom + 8]
    minimum_row_height = max(25, round(pixels.shape[0] * 0.025))
    body_intervals = [
        (top, bottom)
        for top, bottom in zip(row_boundaries, row_boundaries[1:])
        if bottom - top >= minimum_row_height and bottom <= grid_bottom + 8
    ]
    if not body_intervals:
        raise ValueError("OCR 未检测到排班数据行。")
    return {
        "image": image,
        "vertical_lines": vertical_lines,
        "body_intervals": body_intervals,
    }


def _longest_index_run(indexes, max_gap: int = 2) -> tuple[int, int] | None:
    values = [int(value) for value in indexes]
    if not values:
        return None
    runs = [[values[0]]]
    for value in values[1:]:
        if value <= runs[-1][-1] + max_gap:
            runs[-1].append(value)
        else:
            runs.append([value])
    longest = max(runs, key=lambda run: run[-1] - run[0])
    return longest[0], longest[-1]


def _match_ocr_roster_names(text: str, total_names: list[str], corrections: dict | None = None) -> list[str]:
    compact = "".join(re.findall(r"[\u4e00-\u9fa5]", str(text or "")))
    candidates = [
        ("".join(re.findall(r"[\u4e00-\u9fa5]", name)), name)
        for name in total_names
    ]
    candidates = [(normalized, name) for normalized, name in candidates if normalized]
    if corrections:
        for typo, correct in corrections.items():
            compact_typo = "".join(re.findall(r"[\u4e00-\u9fa5]", str(typo)))
            if compact_typo:
                candidates.append((compact_typo, correct))
    candidates.sort(key=lambda item: len(item[0]), reverse=True)
    exact_segmentation = _segment_exact_roster(compact, candidates)
    if exact_segmentation:
        return exact_segmentation
    if 2 <= len(compact) <= 4:
        match = _match_short_ocr_name(compact, candidates)
        return [match] if match else []
    matches = []
    index = 0
    while index < len(compact):
        remaining = compact[index:]
        if 2 <= len(remaining) <= 4:
            short_match = _match_short_ocr_name(remaining, candidates)
            if short_match:
                matches.append(short_match)
            break
        exact = [item for item in candidates if compact.startswith(item[0], index)]
        if exact:
            normalized, original = exact[0]
            matches.append(original)
            index += len(normalized)
            continue
        fuzzy = []
        for normalized, original in candidates:
            segment = compact[index:index + len(normalized)]
            if len(segment) == len(normalized) and edit_distance(segment, normalized) == 1:
                fuzzy.append((len(normalized), normalized, original))
        if fuzzy:
            longest = max(item[0] for item in fuzzy)
            best = [item for item in fuzzy if item[0] == longest]
            if len(best) == 1:
                _length, normalized, original = best[0]
                matches.append(original)
                index += len(normalized)
                continue
        index += 1
    return matches


def _match_short_ocr_name(compact: str, candidates: list[tuple[str, str]]) -> str | None:
    same_length = [item for item in candidates if len(item[0]) == len(compact)]
    if not same_length:
        return None
    distances = [
        (edit_distance(compact, normalized), original)
        for normalized, original in same_length
    ]
    best_distance = min(distance for distance, _original in distances)
    threshold = 1 if len(compact) == 2 else 2
    best = [original for distance, original in distances if distance == best_distance]
    if best_distance <= threshold and len(best) == 1:
        return best[0]
    return None


def _segment_exact_roster(compact: str, candidates: list[tuple[str, str]]) -> list[str]:
    memo: dict[int, list[str] | None] = {}

    def solve(index: int) -> list[str] | None:
        if index == len(compact):
            return []
        if index in memo:
            return memo[index]
        for normalized, original in candidates:
            if compact.startswith(normalized, index):
                remainder = solve(index + len(normalized))
                if remainder is not None:
                    memo[index] = [original, *remainder]
                    return memo[index]
        memo[index] = None
        return None

    return solve(0) or []


def _prepare_ocr_page(pixmap, image_path: Path) -> list[dict]:
    """Remove long table lines and synthesize weekday headers from a six-column grid."""
    pixmap.save(image_path)
    image = Image.open(image_path).convert("L")
    pixels = np.array(image)
    dark = pixels < 160
    horizontal_indexes = np.where(dark.sum(axis=1) > pixels.shape[1] * 0.25)[0]
    vertical_indexes = np.where(dark.sum(axis=0) > pixels.shape[0] * 0.25)[0]
    horizontal_lines = _group_index_centers(horizontal_indexes)
    vertical_lines = _select_seven_grid_lines(_group_index_centers(vertical_indexes))

    for index in horizontal_indexes:
        pixels[max(0, index - 3):min(pixels.shape[0], index + 4), :] = 255
    for index in vertical_indexes:
        pixels[:, max(0, index - 3):min(pixels.shape[1], index + 4)] = 255
    Image.fromarray(pixels).save(image_path)

    if len(vertical_lines) != 7 or len(horizontal_lines) < 2:
        return []
    header_top, header_bottom = horizontal_lines[0], horizontal_lines[1]
    header_height = max(12, header_bottom - header_top - 10)
    synthetic = []
    for index, day in enumerate(("周一", "周二", "周三", "周四", "周五")):
        center = (vertical_lines[index + 1] + vertical_lines[index + 2]) / 2
        synthetic.append({
            "text": day,
            "left": int(center - 15),
            "top": int(header_top + 5),
            "width": 30,
            "height": int(header_height),
            "line_num": 1,
        })
    return synthetic


def _group_index_centers(indexes) -> list[int]:
    values = [int(value) for value in indexes]
    if not values:
        return []
    groups = [[values[0]]]
    for value in values[1:]:
        if value <= groups[-1][-1] + 1:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [round(sum(group) / len(group)) for group in groups]


def _select_seven_grid_lines(lines: list[int]) -> list[int]:
    if len(lines) < 7:
        return []
    best = None
    best_score = None
    for start in range(len(lines) - 6):
        candidate = lines[start:start + 7]
        gaps = np.diff(candidate)
        if min(gaps) <= 0:
            continue
        score = float(np.std(gaps) / np.mean(gaps))
        if best_score is None or score < best_score:
            best = candidate
            best_score = score
    return best or []


def _find_weekday_header(table) -> tuple[int | None, dict[str, int]]:
    for row_index, row in enumerate(table or []):
        columns = {}
        for column, value in enumerate(row or []):
            day = _canonical_weekday(value)
            if day is not None:
                columns[day] = column
        if len(columns) >= 2:
            return row_index, columns
    return None, {}


def _canonical_weekday(value) -> str | None:
    text = re.sub(r"\s+", "", str(value or ""))
    for day, aliases in WEEKDAY_PATTERNS.items():
        for alias in aliases:
            alias_compact = re.sub(r"\s+", "", alias)
            if text == alias_compact:
                return day
            if text.startswith(alias_compact):
                rest = text[len(alias_compact):]
                if _is_weekday_annotation(rest):
                    return day
    return None


def _is_notes_row(row) -> bool:
    text = "".join(str(value or "") for value in (row or []))
    return text.strip().startswith("备注")


def _analyze_ocr_tokens(tokens: list[dict], total_names: list[str], corrections: dict | None = None) -> dict:
    """Reconstruct weekday columns from Tesseract-style positioned tokens."""
    schedule = {day: [] for day in WEEKDAYS}
    unknown_names: list[str] = []
    seen_unknown: set[str] = set()

    headers = {}
    for token in tokens:
        day = _canonical_weekday(token.get("text"))
        if day is not None:
            headers[day] = token
    if len(headers) < 2:
        raise ValueError("OCR 未能定位排班表中的星期列，请使用更清晰的扫描件。")

    header_bottom = max(int(token.get("top", 0)) + int(token.get("height", 0)) for token in headers.values())
    note_tops = [
        int(token.get("top", 0))
        for token in tokens
        if _compact_text(token.get("text", "")).startswith("备注")
    ]
    notes_top = min(note_tops) if note_tops else float("inf")
    header_centers = {
        day: float(token.get("left", 0)) + float(token.get("width", 0)) / 2
        for day, token in headers.items()
    }

    day_tokens: dict[str, list[dict]] = {day: [] for day in header_centers}
    for token in tokens:
        top = int(token.get("top", 0))
        if top <= header_bottom or top >= notes_top:
            continue
        center = float(token.get("left", 0)) + float(token.get("width", 0)) / 2
        day = min(header_centers, key=lambda item: abs(header_centers[item] - center))
        day_tokens[day].append(token)

    for day in sorted(header_centers, key=header_centers.get):
        for line_tokens in _cluster_ocr_lines(day_tokens[day]):
            line_text = "".join(
                str(token.get("text", ""))
                for token in sorted(line_tokens, key=lambda item: int(item.get("left", 0)))
            )
            schedule[day].extend(_match_roster_names(line_text, total_names, corrections))
            for candidate in _unknown_candidates(line_text, total_names, corrections):
                if candidate not in seen_unknown:
                    seen_unknown.add(candidate)
                    unknown_names.append(candidate)

    return {
        "schedule": schedule,
        "matched_count": sum(len(names) for names in schedule.values()),
        "unknown_names": unknown_names,
    }


def _cluster_ocr_lines(tokens: list[dict]) -> list[list[dict]]:
    lines: list[list[dict]] = []
    centers: list[float] = []
    for token in sorted(
        tokens,
        key=lambda item: (
            float(item.get("top", 0)) + float(item.get("height", 0)) / 2,
            int(item.get("left", 0)),
        ),
    ):
        center = float(token.get("top", 0)) + float(token.get("height", 0)) / 2
        height = max(float(token.get("height", 0)), 1.0)
        target = None
        for index, line_center in enumerate(centers):
            line_height = max(float(item.get("height", 0)) for item in lines[index])
            if abs(center - line_center) <= max(12.0, max(height, line_height) * 0.65):
                target = index
                break
        if target is None:
            lines.append([token])
            centers.append(center)
        else:
            lines[target].append(token)
            centers[target] = sum(
                float(item.get("top", 0)) + float(item.get("height", 0)) / 2
                for item in lines[target]
            ) / len(lines[target])
    return [line for _, line in sorted(zip(centers, lines), key=lambda item: item[0])]
