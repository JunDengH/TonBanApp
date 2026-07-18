"""Business orchestration for two ordered final-week statistics."""
from __future__ import annotations

import os
import uuid
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from src.modules.excel_generator import (
    WEEKLY_SHOULD_MAX,
    build_weekly_rows_from_schedules,
    generate_weekly_excel_from_rows,
)
from src.modules.final_weeks_parser import FinalWeekBlock
from src.utils.output_paths import build_unique_output_path
from src.utils.report_titles import (
    build_final_period_report_title,
    build_weekly_report_title,
)


@dataclass
class FinalWeekResult:
    slot: int
    week_number: int
    rows: list[dict]
    warnings: dict = field(default_factory=dict)


@dataclass
class CombinedFinalPeriodResult:
    week_numbers: tuple[int, int]
    rows: list[dict]
    warnings: dict = field(default_factory=dict)


def build_combined_final_period_rows(
    blocks: list[FinalWeekBlock],
    actual_schedules: dict[int, dict[str, list[str]]],
    total_names: list[str],
    senior_assistants: list[str] | None = None,
    senior_should_mode: bool | str = False,
    long_term_leave_assistants: list[str] | None = None,
) -> CombinedFinalPeriodResult:
    """Calculate both final weeks as one period without holiday reduction."""
    if len(blocks) != 2:
        raise ValueError("期末周统计必须包含期末周1和期末周2两个排班区块。")

    for expected_slot, block in enumerate(blocks, start=1):
        if block.slot != expected_slot:
            raise ValueError("期末周排班区块顺序无效，请重新解析 PDF。")
        if actual_schedules.get(expected_slot) is None:
            raise ValueError(f"期末周{expected_slot}缺少已解析的实际值班 Word。")

    total_name_set = set(total_names)
    scheduled_occurrences = [
        name
        for block in blocks
        for names in block.schedule.values()
        for name in names
        if name in total_name_set
    ]
    occurrence_counts = Counter(scheduled_occurrences)
    duplicate_names = {
        name for name, count in occurrence_counts.items() if count > 1
    }
    unique_scheduled_names = [
        name for name in total_names if occurrence_counts[name] > 0
    ]

    merged_actual_names = [
        name
        for slot in (1, 2)
        for names in actual_schedules[slot].values()
        for name in names
    ]
    rows = build_weekly_rows_from_schedules(
        schedule_for_should={"期末周期": unique_scheduled_names},
        schedule_for_actual={"期末周期": merged_actual_names},
        total_names=total_names,
        holidays=[],
        senior_assistants=senior_assistants,
        senior_should_fixed_enabled=senior_should_mode,
        long_term_leave_assistants=long_term_leave_assistants,
        final_week_enabled=True,
    )

    messages = []
    unknown_names = sorted({
        name for block in blocks for name in block.unknown_names
    })
    if duplicate_names:
        messages.append(
            "排班 PDF 中以下姓名重复出现，应值仍只计 2："
            + "、".join(sorted(duplicate_names))
        )
    if unknown_names:
        messages.append("排班 PDF 中存在未识别姓名：" + "、".join(unknown_names))
    over_actual_names = {
        row["姓名"] for row in rows if row["实际班次"] > row["应值班次"]
    }
    if over_actual_names:
        messages.append(
            "以下助理实际班次大于应值班次："
            + "、".join(sorted(over_actual_names))
        )

    return CombinedFinalPeriodResult(
        week_numbers=(blocks[0].week_number, blocks[1].week_number),
        rows=rows,
        warnings={
            "messages": messages,
            "highlight_names": duplicate_names | over_actual_names,
            "duplicate_names": duplicate_names,
        },
    )


def build_final_weeks_rows(
    blocks: list[FinalWeekBlock],
    actual_schedules: dict[int, dict[str, list[str]]],
    total_names: list[str],
    holidays_by_slot: dict[int, list[str]],
    senior_assistants: list[str] | None = None,
    senior_should_mode: bool | str = False,
    long_term_leave_assistants: list[str] | None = None,
) -> list[FinalWeekResult]:
    """Calculate two final-week results without reading or writing files."""
    if len(blocks) != 2:
        raise ValueError("期末周统计必须包含期末周1和期末周2两个排班区块。")

    results = []
    for expected_slot, block in enumerate(blocks, start=1):
        if block.slot != expected_slot:
            raise ValueError("期末周排班区块顺序无效，请重新解析 PDF。")
        actual_schedule = actual_schedules.get(block.slot)
        if actual_schedule is None:
            raise ValueError(f"期末周{block.slot}缺少已解析的实际值班 Word。")
        rows = build_weekly_rows_from_schedules(
            schedule_for_should=block.schedule,
            schedule_for_actual=actual_schedule,
            total_names=total_names,
            holidays=holidays_by_slot.get(block.slot, []),
            senior_assistants=senior_assistants,
            senior_should_fixed_enabled=senior_should_mode,
            long_term_leave_assistants=long_term_leave_assistants,
            final_week_enabled=True,
        )
        highlight_names = {
            row["姓名"] for row in rows if row.get("应值班次", 0) > WEEKLY_SHOULD_MAX
        }
        messages = []
        if block.unknown_names:
            messages.append(
                f"期末周{block.slot}（第{block.week_number}周）排班 PDF 中存在未识别姓名："
                + "、".join(block.unknown_names)
            )
        if highlight_names:
            messages.append(
                f"期末周{block.slot}（第{block.week_number}周）以下助理应值班次大于 {WEEKLY_SHOULD_MAX}："
                + "、".join(sorted(highlight_names))
            )
        results.append(FinalWeekResult(
            slot=block.slot,
            week_number=block.week_number,
            rows=rows,
            warnings={"messages": messages, "highlight_names": highlight_names},
        ))
    return results


def write_combined_final_period_excel(
    result: CombinedFinalPeriodResult,
    output_dir: str | Path,
    year: str,
    season: str,
) -> str:
    """Write one combined final-period workbook through a temporary file."""
    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)
    title = build_final_period_report_title(year, season)
    final_path = build_unique_output_path(output_root, title, ".xlsx")
    temp_path = output_root / f".{final_path.stem}.{uuid.uuid4().hex}.tmp.xlsx"
    try:
        generate_weekly_excel_from_rows(result.rows, str(temp_path))
        os.replace(temp_path, final_path)
        return str(final_path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def write_final_weeks_excels(
    results: list[FinalWeekResult],
    output_dir: str | Path,
    year: str,
    season: str,
) -> list[str]:
    """Generate both workbooks through temporary files, then publish together."""
    if len(results) != 2:
        raise ValueError("必须同时提供期末周1和期末周2的预览数据。")
    if results[0].week_number == results[1].week_number:
        raise ValueError("期末周1和期末周2不能使用相同周次。")

    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)
    temp_paths: list[Path] = []
    final_paths: list[Path] = []
    published_paths: list[Path] = []
    try:
        for result in results:
            title = build_weekly_report_title(year, season, result.week_number)
            final_path = build_unique_output_path(output_root, title, ".xlsx")
            temp_path = output_root / f".{final_path.stem}.{uuid.uuid4().hex}.tmp.xlsx"
            temp_paths.append(temp_path)
            final_paths.append(final_path)
            generate_weekly_excel_from_rows(result.rows, str(temp_path))

        for temp_path, final_path in zip(temp_paths, final_paths):
            os.replace(temp_path, final_path)
            published_paths.append(final_path)
        return [str(path) for path in final_paths]
    except Exception:
        for path in temp_paths:
            path.unlink(missing_ok=True)
        for path in published_paths:
            path.unlink(missing_ok=True)
        raise
