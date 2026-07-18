import inspect
from pathlib import Path

import pytest

from src.modules.final_weeks_parser import FinalWeekBlock
from src.modules import final_weeks_service as service
from src.modules.final_weeks_service import (
    build_combined_final_period_rows,
    build_final_weeks_rows,
    write_combined_final_period_excel,
    write_final_weeks_excels,
)


def block(slot: int, week_number: int, name: str) -> FinalWeekBlock:
    schedule = {"周一": [name], "周二": [], "周三": [], "周四": [], "周五": []}
    return FinalWeekBlock(slot, week_number, schedule, 1, [])


def row_for(result, name):
    return next(row for row in result.rows if row["姓名"] == name)


def combined_row(result, name):
    return next(row for row in result.rows if row["姓名"] == name)


def sample_results():
    return build_final_weeks_rows(
        blocks=[block(1, 16, "张三"), block(2, 18, "李四")],
        actual_schedules={
            1: {"周一": ["张三"]},
            2: {"周一": []},
        },
        total_names=["张三", "李四"],
        holidays_by_slot={1: [], 2: []},
        senior_assistants=[],
        senior_should_mode="normal",
        long_term_leave_assistants=[],
    )


def test_build_final_weeks_rows_keeps_slots_independent():
    results = sample_results()

    assert [result.slot for result in results] == [1, 2]
    assert [result.week_number for result in results] == [16, 18]
    assert row_for(results[0], "张三")["实际班次"] == 2
    assert row_for(results[0], "李四")["应值班次"] == 0
    assert row_for(results[1], "李四")["应值班次"] == 2
    assert row_for(results[1], "李四")["缺班"] == 2


def test_write_final_weeks_excels_uses_actual_week_numbers(tmp_path, monkeypatch):
    def fake_generate(rows, path):
        Path(path).write_text(str(len(rows)), encoding="utf-8")
        return path

    monkeypatch.setattr(service, "generate_weekly_excel_from_rows", fake_generate)

    outputs = write_final_weeks_excels(sample_results(), tmp_path, "2026", "春季")

    assert [Path(path).name for path in outputs] == [
        "2026春季学期第16周助理值班统计.xlsx",
        "2026春季学期第18周助理值班统计.xlsx",
    ]
    assert all(Path(path).exists() for path in outputs)


def test_write_final_weeks_excels_cleans_partial_files(tmp_path, monkeypatch):
    calls = 0

    def fail_second(rows, path):
        nonlocal calls
        calls += 1
        Path(path).write_text("temp", encoding="utf-8")
        if calls == 2:
            raise PermissionError("occupied")
        return path

    monkeypatch.setattr(service, "generate_weekly_excel_from_rows", fail_second)

    with pytest.raises(PermissionError, match="occupied"):
        write_final_weeks_excels(sample_results(), tmp_path, "2026", "春季")

    assert list(tmp_path.iterdir()) == []


def test_build_final_weeks_rows_rejects_missing_slot_schedule():
    with pytest.raises(ValueError, match="期末周2.*实际"):
        build_final_weeks_rows(
            blocks=[block(1, 16, "张三"), block(2, 18, "李四")],
            actual_schedules={1: {"周一": ["张三"]}},
            total_names=["张三", "李四"],
            holidays_by_slot={1: [], 2: []},
        )


def test_combined_period_deduplicates_pdf_names_and_merges_actual_occurrences():
    blocks = [
        FinalWeekBlock(1, 18, {"周一": ["张三", "张三"]}, 2, []),
        FinalWeekBlock(2, 19, {"周一": ["张三", "李四"]}, 2, []),
    ]
    result = build_combined_final_period_rows(
        blocks=blocks,
        actual_schedules={
            1: {"周一": ["张三"]},
            2: {"周二": ["张三", "李四"]},
        },
        total_names=["张三", "李四"],
    )

    assert result.week_numbers == (18, 19)
    assert combined_row(result, "张三") == {
        "姓名": "张三",
        "应值班次": 2,
        "实际班次": 4,
        "缺班": 0,
        "备注": "",
    }
    assert combined_row(result, "李四")["应值班次"] == 2
    assert combined_row(result, "李四")["实际班次"] == 2
    assert result.warnings["duplicate_names"] == {"张三"}


def test_combined_period_applies_long_leave_once_and_senior_override_once():
    result = build_combined_final_period_rows(
        blocks=[block(1, 18, "张三"), block(2, 19, "李四")],
        actual_schedules={1: {"周一": []}, 2: {"周一": []}},
        total_names=["张三", "李四", "王五", "赵六"],
        senior_assistants=["赵六"],
        senior_should_mode="none",
        long_term_leave_assistants=["王五", "赵六"],
    )

    assert combined_row(result, "王五")["应值班次"] == 2
    assert combined_row(result, "王五")["缺班"] == 2
    assert combined_row(result, "赵六")["应值班次"] == 0
    assert combined_row(result, "赵六")["缺班"] == 0


def test_combined_period_rejects_missing_actual_slot_and_reduced_senior_mode():
    blocks = [block(1, 18, "张三"), block(2, 19, "李四")]
    with pytest.raises(ValueError, match="期末周2.*实际"):
        build_combined_final_period_rows(
            blocks,
            {1: {"周一": []}},
            ["张三", "李四"],
        )

    with pytest.raises(ValueError, match="少值班"):
        build_combined_final_period_rows(
            blocks,
            {1: {"周一": []}, 2: {"周一": []}},
            ["张三", "李四"],
            senior_assistants=["张三"],
            senior_should_mode="reduced",
        )


def test_combined_period_api_has_no_holiday_parameter():
    parameters = inspect.signature(build_combined_final_period_rows).parameters
    assert "holidays" not in parameters
    assert "holidays_by_slot" not in parameters


def test_write_combined_final_period_excel_uses_single_period_filename(
    tmp_path, monkeypatch
):
    result = build_combined_final_period_rows(
        [block(1, 18, "张三"), block(2, 19, "李四")],
        {1: {"周一": ["张三"]}, 2: {"周一": ["李四"]}},
        ["张三", "李四"],
    )

    def fake_generate(rows, path):
        Path(path).write_text(str(len(rows)), encoding="utf-8")
        return path

    monkeypatch.setattr(service, "generate_weekly_excel_from_rows", fake_generate)
    output = write_combined_final_period_excel(
        result,
        tmp_path,
        "2026",
        "春季",
    )

    assert Path(output).name == "2026春季学期期末周助理值班统计.xlsx"
    assert Path(output).exists()
    assert not list(tmp_path.glob("*.tmp.xlsx"))


def test_write_combined_final_period_excel_cleans_temp_on_failure(
    tmp_path, monkeypatch
):
    result = build_combined_final_period_rows(
        [block(1, 18, "张三"), block(2, 19, "李四")],
        {1: {"周一": []}, 2: {"周一": []}},
        ["张三", "李四"],
    )

    def fail_generate(rows, path):
        Path(path).write_text("temp", encoding="utf-8")
        raise PermissionError("occupied")

    monkeypatch.setattr(service, "generate_weekly_excel_from_rows", fail_generate)
    with pytest.raises(PermissionError, match="occupied"):
        write_combined_final_period_excel(result, tmp_path, "2026", "春季")
    assert list(tmp_path.iterdir()) == []
