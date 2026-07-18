"""Report title helpers shared by UI and tests."""


def build_weekly_report_title(year: str, season: str, week: str | int) -> str:
    year_text = str(year).strip()
    season_text = str(season).strip()
    week_text = str(week).strip()
    if not (year_text.isdigit() and len(year_text) == 4):
        raise ValueError("标题年份请填写4位数字，如 2026。")
    if season_text not in ("春季", "秋季"):
        raise ValueError("标题学期只能选择 春季 或 秋季。")
    if not week_text.isdigit() or int(week_text) <= 0:
        raise ValueError("标题周次请填写正整数。")
    return f"{year_text}{season_text}学期第{int(week_text)}周助理值班统计"


def build_final_period_report_title(year: str, season: str) -> str:
    year_text = str(year).strip()
    season_text = str(season).strip()
    if not (year_text.isdigit() and len(year_text) == 4):
        raise ValueError("标题年份请填写4位数字，如 2026。")
    if season_text not in ("春季", "秋季"):
        raise ValueError("标题学期只能选择 春季 或 秋季。")
    return f"{year_text}{season_text}学期期末周助理值班统计"


def build_monthly_report_title(year: str, season: str, start_week: str | int, end_week: str | int) -> str:
    year_text = str(year).strip()
    season_text = str(season).strip()
    start_text = str(start_week).strip()
    end_text = str(end_week).strip()
    if not (year_text.isdigit() and len(year_text) == 4):
        raise ValueError("标题年份请填写4位数字，如 2026。")
    if season_text not in ("春季", "秋季"):
        raise ValueError("标题学期只能选择 春季 或 秋季。")
    if not (start_text.isdigit() and end_text.isdigit()):
        raise ValueError("标题周数请填写正整数。")
    start = int(start_text)
    end = int(end_text)
    if start <= 0 or end <= 0:
        raise ValueError("标题周数必须大于 0。")
    if start > end:
        raise ValueError("标题起始周不能大于结束周。")
    return f"{year_text}{season_text}学期{start}-{end}周助理值班统计"
