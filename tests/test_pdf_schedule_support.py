from pathlib import Path

import fitz
import pytest
from docx import Document

from src.modules.schedule_parser import parse_schedule_file, scan_schedule_file


def make_docx(path: Path, lines: list[str]) -> None:
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)


def make_native_schedule_pdf(path: Path, pages: list[list[list[str]]]) -> None:
    """Create small text-layer PDFs with six-column schedule tables."""
    doc = fitz.open()
    for page_rows in pages:
        page = doc.new_page(width=600, height=260)
        left, top, col_width, row_height = 20, 20, 92, 42
        rows = [["", "周一", "周二", "周三", "周四", "周五"], *page_rows]
        for row_idx in range(len(rows) + 1):
            y = top + row_idx * row_height
            page.draw_line((left, y), (left + 6 * col_width, y), color=(0, 0, 0))
        for col_idx in range(7):
            x = left + col_idx * col_width
            page.draw_line((x, top), (x, top + len(rows) * row_height), color=(0, 0, 0))
        for row_idx, row in enumerate(rows):
            for col_idx, text in enumerate(row):
                page.insert_text(
                    (left + col_idx * col_width + 5, top + row_idx * row_height + 22),
                    text,
                    fontsize=10,
                    fontname="china-s",
                )
    doc.save(path)


def make_scanned_schedule_pdf(path: Path, tmp_path: Path) -> None:
    native_path = tmp_path / "scan-source.pdf"
    make_native_schedule_pdf(
        native_path,
        [[
            ["1-2", "张三 李四", "王五", "", "", ""],
            ["3-4", "张三", "", "赵六", "", ""],
        ]],
    )
    source = fitz.open(native_path)
    scanned = fitz.open()
    for source_page in source:
        pixmap = source_page.get_pixmap(matrix=fitz.Matrix(4, 4), alpha=False)
        page = scanned.new_page(width=source_page.rect.width, height=source_page.rect.height)
        page.insert_image(page.rect, stream=pixmap.tobytes("png"))
    scanned.save(path)


def test_schedule_dispatch_keeps_docx_behavior(tmp_path):
    path = tmp_path / "schedule.docx"
    make_docx(path, ["周一", "张三 李四"])

    schedule = parse_schedule_file(str(path), ["张三", "李四"])
    scan = scan_schedule_file(str(path), ["张三", "李四"])

    assert schedule["周一"] == ["张三", "李四"]
    assert scan == {"matched_count": 2, "unknown_names": []}


def test_schedule_dispatch_rejects_unknown_extension(tmp_path):
    path = tmp_path / "schedule.txt"
    path.write_text("周一 张三", encoding="utf-8")

    with pytest.raises(ValueError, match="PDF 或 Word"):
        parse_schedule_file(str(path), ["张三"])
    with pytest.raises(ValueError, match="PDF 或 Word"):
        scan_schedule_file(str(path), ["张三"])


def test_native_pdf_assigns_weekdays_and_preserves_duplicate_names(tmp_path):
    path = tmp_path / "schedule.pdf"
    make_native_schedule_pdf(
        path,
        [[
            ["1-2", "张 三 李四 张三", "王五", "", "", ""],
            ["3-4", "", "", "赵六", "", ""],
        ]],
    )

    schedule = parse_schedule_file(str(path), ["张三", "李四", "王五", "赵六"])

    assert schedule["周一"] == ["张三", "李四", "张三"]
    assert schedule["周二"] == ["王五"]
    assert schedule["周三"] == ["赵六"]


def test_native_pdf_merges_multiple_pages(tmp_path):
    path = tmp_path / "multi.pdf"
    make_native_schedule_pdf(
        path,
        [
            [["1-2", "张三", "", "", "", ""]],
            [["3-4", "李四", "王五", "", "", ""]],
        ],
    )

    schedule = parse_schedule_file(str(path), ["张三", "李四", "王五"])

    assert schedule["周一"] == ["张三", "李四"]
    assert schedule["周二"] == ["王五"]


def test_corrupt_pdf_raises_clear_error(tmp_path):
    path = tmp_path / "broken.pdf"
    path.write_bytes(b"not a pdf")

    with pytest.raises(ValueError, match="PDF"):
        parse_schedule_file(str(path), ["张三"])


def test_ocr_geometry_assigns_columns_preserves_duplicates_and_skips_notes():
    from src.modules.pdf_schedule_parser import _analyze_ocr_tokens

    tokens = []
    for index, day in enumerate(["周一", "周二", "周三", "周四", "周五"]):
        tokens.append({"text": day, "left": 100 + index * 100, "top": 20, "width": 30, "height": 18, "line_num": 1})
    tokens.extend([
        {"text": "1-2", "left": 20, "top": 70, "width": 30, "height": 16, "line_num": 2},
        {"text": "张", "left": 105, "top": 70, "width": 14, "height": 16, "line_num": 2},
        {"text": "三", "left": 120, "top": 70, "width": 14, "height": 16, "line_num": 2},
        {"text": "李四", "left": 140, "top": 70, "width": 28, "height": 16, "line_num": 2},
        {"text": "张三", "left": 105, "top": 105, "width": 28, "height": 16, "line_num": 3},
        {"text": "王五", "left": 205, "top": 70, "width": 28, "height": 16, "line_num": 2},
        {"text": "备注", "left": 20, "top": 180, "width": 35, "height": 16, "line_num": 8},
        {"text": "赵六", "left": 105, "top": 210, "width": 28, "height": 16, "line_num": 9},
    ])

    analysis = _analyze_ocr_tokens(tokens, ["张三", "李四", "王五", "赵六"])

    assert analysis["schedule"]["周一"] == ["张三", "李四", "张三"]
    assert analysis["schedule"]["周二"] == ["王五"]
    assert "赵六" not in analysis["schedule"]["周一"]
    assert analysis["matched_count"] == 4


def test_ocr_runtime_resolves_engine_and_models_from_supplied_root(tmp_path):
    from src.modules.ocr_runtime import resolve_tesseract_runtime

    runtime = tmp_path / "vendor" / "tesseract"
    tessdata = runtime / "tessdata"
    tessdata.mkdir(parents=True)
    (runtime / "tesseract.exe").write_bytes(b"binary")
    (tessdata / "chi_sim.traineddata").write_bytes(b"model")
    (tessdata / "eng.traineddata").write_bytes(b"model")

    executable, resolved_tessdata = resolve_tesseract_runtime(tmp_path)

    assert executable == runtime / "tesseract.exe"
    assert resolved_tessdata == tessdata


def test_ocr_runtime_rejects_missing_chinese_model(tmp_path):
    from src.modules.ocr_runtime import resolve_tesseract_runtime

    runtime = tmp_path / "vendor" / "tesseract"
    runtime.mkdir(parents=True)
    (runtime / "tesseract.exe").write_bytes(b"binary")

    with pytest.raises(ValueError, match="中文模型"):
        resolve_tesseract_runtime(tmp_path)


def test_scanned_pdf_uses_bundled_ocr_end_to_end(tmp_path):
    path = tmp_path / "scanned.pdf"
    make_scanned_schedule_pdf(path, tmp_path)

    schedule = parse_schedule_file(str(path), ["张三", "李四", "王五", "赵六"])

    assert schedule["周一"] == ["张三", "李四", "张三"]
    assert schedule["周二"] == ["王五"]
    assert schedule["周三"] == ["赵六"]


def test_ocr_name_matching_repairs_unique_single_character_errors():
    from src.modules.pdf_schedule_parser import _match_ocr_roster_names

    assert _match_ocr_roster_names(
        "万张洲吴晓敏柏一敬付雨碗",
        ["万张洲", "吴晓敏", "柏一茗", "付雨鑫"],
    ) == ["万张洲", "吴晓敏", "柏一茗", "付雨鑫"]


def test_ocr_name_matching_prefers_same_length_candidate_over_shorter_name():
    from src.modules.pdf_schedule_parser import _match_ocr_roster_names

    assert _match_ocr_roster_names("王元表", ["王博", "王歆玥"]) == ["王歆玥"]


def test_scanned_grid_detection_finds_weekday_columns_and_body_rows(tmp_path):
    from src.modules.pdf_schedule_parser import _detect_table_grid

    path = tmp_path / "scanned.pdf"
    make_scanned_schedule_pdf(path, tmp_path)
    document = fitz.open(path)
    pixmap = document[0].get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72), alpha=False)

    grid = _detect_table_grid(pixmap)

    assert len(grid["vertical_lines"]) == 7
    assert len(grid["body_intervals"]) == 2


def test_weekly_ui_presents_pdf_first_schedule_picker_and_word_only_actual_picker():
    source = (Path(__file__).parents[1] / "src" / "ui" / "weekly_tab.py").read_text(encoding="utf-8")

    assert "排班文件（PDF 优先，兼容 Word）" in source
    assert '("PDF 排班表", "*.pdf")' in source
    assert '("Word 排班表", "*.docx")' in source
    assert 'title="选择实际周统计名单Word"' in source
    assert 'filetypes=[("Word文件", "*.docx")]' in source


def test_pyinstaller_spec_bundles_ocr_runtime_without_local_json_data():
    spec = (Path(__file__).parents[1] / "build" / "pyinstaller" / "TonBanAPP.spec").read_text(encoding="utf-8")

    assert r"vendor\\tesseract" in spec
    assert "tesseract.exe" in spec
    assert "chi_sim.traineddata" in spec
    assert "data\\config.json" not in spec
