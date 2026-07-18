from pathlib import Path

import pytest

from src.ui.final_weeks_tab import _validate_week_numbers


ROOT = Path(__file__).parents[1]


def test_final_weeks_page_uses_slot_names_and_required_inputs():
    source = (ROOT / "src" / "ui" / "final_weeks_tab.py").read_text(encoding="utf-8")

    assert "期末周1" in source
    assert "期末周2" in source
    assert "期末排班 PDF" in source
    assert source.count("实际值班 Word") >= 2
    assert "18周" not in source
    assert "19周" not in source


def test_final_weeks_page_defaults_to_combined_single_export_without_holidays():
    source = (ROOT / "src" / "ui" / "final_weeks_tab.py").read_text(encoding="utf-8")

    assert "FINAL_WEEKS_SPLIT_EXPORT_ENABLED = False" in source
    assert "预览期末周统计" in source
    assert "生成期末周 Excel" in source
    assert "两周期末周合并为一个统计周期" in source
    assert "if FINAL_WEEKS_SPLIT_EXPORT_ENABLED:" in source
    assert "build_combined_final_period_rows" in source
    assert "write_combined_final_period_excel" in source


def test_final_weeks_split_export_ui_and_service_calls_remain_available():
    source = (ROOT / "src" / "ui" / "final_weeks_tab.py").read_text(encoding="utf-8")

    assert "预览两周统计" in source
    assert "生成两周 Excel" in source
    assert "build_final_weeks_rows" in source
    assert "write_final_weeks_excels" in source
    assert "self._open_split_preview" in source


def test_holiday_controls_are_only_constructed_in_split_mode():
    source = (ROOT / "src" / "ui" / "final_weeks_tab.py").read_text(encoding="utf-8")
    guard = source.index(
        "if FINAL_WEEKS_SPLIT_EXPORT_ENABLED:",
        source.index("def _build_slot_panel"),
    )
    holiday = source.index("holiday = ctk.CTkFrame", guard)
    next_method = source.index("def _build_check_panel", holiday)
    assert guard < holiday < next_method


def test_week_numbers_allow_non_consecutive_positive_values():
    assert _validate_week_numbers("16", "21") == (16, 21)


@pytest.mark.parametrize(
    ("first", "second", "message"),
    [
        ("", "19", "正整数"),
        ("0", "19", "正整数"),
        ("18", "18", "不能相同"),
        ("A", "19", "正整数"),
    ],
)
def test_week_numbers_reject_invalid_or_duplicate_values(first, second, message):
    with pytest.raises(ValueError, match=message):
        _validate_week_numbers(first, second)


def test_navigation_places_final_weeks_between_weekly_and_monthly():
    source = (ROOT / "src" / "ui" / "main_ui.py").read_text(encoding="utf-8")

    weekly = source.index('{"key": "weekly", "icon": "weekly", "label": "周统计"}')
    final_weeks = source.index('{"key": "final_weeks", "icon": "final_weeks", "label": "期末周统计"}')
    monthly = source.index('{"key": "monthly", "icon": "monthly", "label": "月统计"}')
    assert weekly < final_weeks < monthly
    assert "from src.ui.final_weeks_tab import FinalWeeksTab" in source
    assert 'self.pages["final_weeks"] = FinalWeeksTab(content, self.data_mgr)' in source


def test_old_final_week_toggle_is_removed_from_rules_and_regular_weekly_page():
    rules = (ROOT / "src" / "ui" / "rules_tab.py").read_text(encoding="utf-8")
    weekly = (ROOT / "src" / "ui" / "weekly_tab.py").read_text(encoding="utf-8")

    assert "_build_final_week_card" not in rules
    assert "get_final_week_enabled" not in rules
    assert "set_final_week_enabled" not in rules
    assert "get_final_week_enabled" not in weekly
    assert "final_week_enabled=False" in weekly


def test_preview_renders_window_chrome_before_building_large_tables():
    source = (ROOT / "src" / "ui" / "final_weeks_tab.py").read_text(encoding="utf-8")

    combined_start = source.index("def _open_combined_preview")
    combined_render = source.index("dialog.update_idletasks()", combined_start)
    combined_rows = source.index("self._build_preview_table(preview, result)", combined_start)
    assert combined_render < combined_rows

    split_start = source.index("def _open_split_preview")
    split_render = source.index("dialog.update_idletasks()", split_start)
    split_rows = source.index("self._build_preview_table(tab, result)", split_start)
    assert split_render < split_rows


def test_readme_documents_combined_period_without_holiday_reduction():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert "每个唯一姓名整个期末周期应值 2" in readme
    assert "期末周统计不提供或执行放假核减" in readme
    assert "学期期末周助理值班统计.xlsx" in readme
