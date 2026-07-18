from pathlib import Path

import fitz
import pytest

from src.modules.final_weeks_parser import _find_week_headings, parse_final_weeks_pdf


def make_final_weeks_pdf(path: Path, blocks: list[tuple[int, dict[str, list[str]]]]) -> None:
    """Create a native-text PDF with vertically stacked final-week blocks."""
    doc = fitz.open()
    page = doc.new_page(width=760, height=max(320, 220 * len(blocks)))
    day_order = ["周一", "周二", "周三", "周四", "周五"]
    left = 28
    label_width = 100
    day_width = 122

    for block_index, (week_number, schedule) in enumerate(blocks):
        top = 30 + block_index * 210
        visible_days = [day for day in day_order if day in schedule]
        right = left + label_width + max(len(visible_days), 1) * day_width
        for row in range(4):
            y = top + row * 48
            page.draw_line((left, y), (right, y), color=(0, 0, 0))
        verticals = [left, left + label_width]
        verticals.extend(left + label_width + index * day_width for index in range(1, len(visible_days) + 1))
        for x in verticals:
            page.draw_line((x, top), (x, top + 3 * 48), color=(0, 0, 0))

        page.insert_text((left + 20, top + 30), f"{week_number}周", fontsize=12, fontname="china-s")
        for day_index, day in enumerate(visible_days):
            x = left + label_width + day_index * day_width
            page.insert_text((x + 36, top + 30), day, fontsize=11, fontname="china-s")
            names = schedule.get(day, [])
            midpoint = (len(names) + 1) // 2
            page.insert_text((x + 8, top + 76), " ".join(names[:midpoint]), fontsize=10, fontname="china-s")
            page.insert_text((x + 8, top + 124), " ".join(names[midpoint:]), fontsize=10, fontname="china-s")
        page.insert_text((left + 25, top + 76), "上午", fontsize=10, fontname="china-s")
        page.insert_text((left + 25, top + 124), "下午", fontsize=10, fontname="china-s")

    doc.save(path)


def test_parse_two_non_consecutive_week_blocks(tmp_path):
    path = tmp_path / "final-weeks.pdf"
    make_final_weeks_pdf(
        path,
        [
            (16, {"周一": ["张三"], "周二": ["李四"]}),
            (18, {"周一": ["王五"], "周二": ["赵六"]}),
        ],
    )

    blocks = parse_final_weeks_pdf(str(path), ["张三", "李四", "王五", "赵六"])

    assert [block.slot for block in blocks] == [1, 2]
    assert [block.week_number for block in blocks] == [16, 18]
    assert blocks[0].schedule["周一"] == ["张三"]
    assert blocks[0].schedule["周二"] == ["李四"]
    assert blocks[1].schedule["周一"] == ["王五"]
    assert blocks[1].schedule["周二"] == ["赵六"]


def test_week_heading_tokens_are_ordered_horizontally_within_the_same_line():
    words = [
        {"text": "周", "x0": 72.84, "x1": 84.84, "top": 122.672, "bottom": 134.672},
        {"text": "18", "x0": 57.6, "x1": 69.804, "top": 123.74, "bottom": 135.74},
        {"text": "周一（6.29）", "x0": 148.94, "x1": 218.54, "top": 122.672, "bottom": 135.74},
    ]

    headings = _find_week_headings(words)

    assert headings == [{"week_number": 18, "top": 122.672, "x0": 57.6}]


def test_weekday_names_do_not_merge_across_week_blocks(tmp_path):
    path = tmp_path / "separated.pdf"
    make_final_weeks_pdf(
        path,
        [
            (9, {"周一": ["张三", "张三"], "周二": ["李四"]}),
            (12, {"周一": ["王五"], "周二": ["赵六"]}),
        ],
    )

    blocks = parse_final_weeks_pdf(str(path), ["张三", "李四", "王五", "赵六"])

    assert blocks[0].schedule["周一"] == ["张三", "张三"]
    assert blocks[1].schedule["周一"] == ["王五"]
    assert "王五" not in blocks[0].schedule["周一"]


def test_corrections_are_applied_inside_the_matching_week_block(tmp_path):
    path = tmp_path / "typo.pdf"
    make_final_weeks_pdf(
        path,
        [
            (7, {"周一": ["张灏琛"]}),
            (11, {"周一": ["张雯捷"]}),
        ],
    )

    blocks = parse_final_weeks_pdf(
        str(path),
        ["张颢琛", "张雯婕"],
        {"张灏琛": "张颢琛", "张雯捷": "张雯婕"},
    )

    assert blocks[0].schedule["周一"] == ["张颢琛"]
    assert blocks[1].schedule["周一"] == ["张雯婕"]


@pytest.mark.parametrize(
    ("weeks", "message"),
    [
        ([18], "恰好包含两个"),
        ([18, 18], "相同周次"),
        ([16, 18, 20], "恰好包含两个"),
    ],
)
def test_requires_exactly_two_distinct_week_blocks(tmp_path, weeks, message):
    path = tmp_path / "invalid.pdf"
    make_final_weeks_pdf(path, [(week, {"周一": ["张三"]}) for week in weeks])

    with pytest.raises(ValueError, match=message):
        parse_final_weeks_pdf(str(path), ["张三"])
