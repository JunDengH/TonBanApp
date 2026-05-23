# src/modules/excel_generator.py
"""
Excel生成模块：
A. 周统计Excel（5列）
B. 月统计Excel（6列，含斜线样式）
"""
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.modules.word_parser import parse_weekly_schedule, count_actual_shifts
from src.utils.helpers import pinyin_key


# ============================================================
# 子功能 A：周统计 Excel
# ============================================================
def generate_weekly_excel(
    word_path: str,
    total_names: list,
    holidays: list,
    output_path: str,
    actual_word_path: str | None = None,
    senior_assistants: list | None = None,
    senior_should_fixed_enabled: bool = False,
) -> str:
    """
    生成周统计Excel。
    参数:
        word_path: 排班名单Word文件路径（用于应值班次计算）
        total_names: 总名单（来自 data_manager）
        holidays: 放假的星期列表，如 ["周三", "周五"]
        output_path: 输出文件完整路径（.xlsx）
        actual_word_path: 实际周统计人员名单Word路径（用于实际班次计算）；
                 为空时回退使用 word_path（兼容旧流程）
        senior_assistants: 被标记为大四助理的姓名列表
        senior_should_fixed_enabled: 开启后，大四助理应值班次强制为 1
    返回:
        实际写入的文件路径
    """
    # 1) 解析排班名单：用于应值班次
    weekly_schedule_for_should = parse_weekly_schedule(word_path, total_names)

    # 2) 解析实际名单：用于实际班次（兼容旧流程）
    actual_source_path = actual_word_path or word_path
    weekly_schedule_for_actual = parse_weekly_schedule(actual_source_path, total_names)
    actual = count_actual_shifts(weekly_schedule_for_actual, total_names)

    # 3) 应值班次：优先按排班名单统计，再按放假日期核减
    should = count_actual_shifts(weekly_schedule_for_should, total_names)
    holiday_reduction_counter = {n: 0 for n in total_names}
    for holiday in holidays or []:
        if holiday not in weekly_schedule_for_should:
            continue
        # 统计当日每人排班次数
        day_names = weekly_schedule_for_should[holiday]
        per_person = {}
        for n in day_names:
            per_person[n] = per_person.get(n, 0) + 1
        # 核减
        for person, times in per_person.items():
            if person in holiday_reduction_counter:
                holiday_reduction_counter[person] += times
            should[person] = max(0, should[person] - times)

    # 4) 可选规则：大四助理应值班次固定为 1
    if senior_should_fixed_enabled and senior_assistants:
        senior_set = set(senior_assistants)
        for name in total_names:
            if name in senior_set:
                # 放假优先级更高：先固定为1，再按放假核减
                should[name] = max(0, 1 - holiday_reduction_counter.get(name, 0))

    # 5. 构建DataFrame
    rows = []
    for name in total_names:
        s = should[name]
        a = actual.get(name, 0)
        absence = max(0, s - a)
        rows.append({
            "姓名": name,
            "应值班次": s,
            "实际班次": a,
            "缺班": absence,
            "备注": "",
        })
    df = pd.DataFrame(rows, columns=["姓名", "应值班次", "实际班次", "缺班", "备注"])

    # 6. 按缺班倒序
    df = df.sort_values(by="缺班", ascending=False, kind="mergesort").reset_index(drop=True)

    # 7. 写入Excel（带样式）
    _safe_write_excel(df, output_path, sheet_name="周统计", apply_weekly_style=True)
    return output_path


# ============================================================
# 公共：聚合 4 周 Excel
# ============================================================
def aggregate_weekly_files(weekly_excel_paths: list, total_names: list):
    """
    读取多个周统计 Excel，按总名单聚合"实际班次""应值班次"。
    返回:
        (total_actual: dict[name,int], total_should: dict[name,int])
    说明:
        该函数同时为 excel_generator 与 word_generator 所共用，集中周表解析逻辑。
    """
    if len(weekly_excel_paths) < 1:
        raise ValueError("至少需要提供一个周统计Excel文件")

    total_actual = {n: 0 for n in total_names}
    total_should = {n: 0 for n in total_names}

    for path in weekly_excel_paths:
        df = pd.read_excel(path)
        required = {"姓名", "应值班次", "实际班次"}
        if not required.issubset(df.columns):
            raise ValueError(f"文件缺少必要列：{path}")
        for _, row in df.iterrows():
            name = str(row["姓名"]).strip()
            if name in total_actual:
                total_actual[name] += int(row["实际班次"]) if pd.notna(row["实际班次"]) else 0
                total_should[name] += int(row["应值班次"]) if pd.notna(row["应值班次"]) else 0

    return total_actual, total_should


# ============================================================
# 子功能 B：月统计 Excel（保留作为回退入口，需求3 UI 不再暴露）
# ============================================================
def generate_monthly_excel(
    weekly_excel_paths: list,
    total_names: list,
    output_path: str,
) -> str:
    """
    基于4个周统计Excel生成月统计。
    参数:
        weekly_excel_paths: 4个周Excel的路径列表
        total_names: 总名单
        output_path: 输出文件完整路径
    """
    total_actual, total_should = aggregate_weekly_files(weekly_excel_paths, total_names)

    # 构建DataFrame
    rows = []
    for name in total_names:
        a = total_actual[name]
        s = total_should[name]
        over = a - s          # 多值班次
        absence = s - a       # 缺班数量
        rows.append({
            "姓名": name,
            "总计班次": a,
            "应值班次": s,
            "多值班次": over,
            "缺班数量": absence,
        })
    df = pd.DataFrame(rows)

    # 排序：多值班次降序 -> 缺班数量升序 -> 姓名拼音升序
    df["_pinyin"] = df["姓名"].apply(pinyin_key)
    df = df.sort_values(
        by=["多值班次", "缺班数量", "_pinyin"],
        ascending=[False, True, True],
        kind="mergesort",
    ).drop(columns=["_pinyin"]).reset_index(drop=True)

    # 插入序号列
    df.insert(0, "序号", range(1, len(df) + 1))
    df = df[["序号", "姓名", "总计班次", "应值班次", "多值班次", "缺班数量"]]

    # 写入并应用特殊样式（<=0 单元格画斜线）
    _safe_write_excel(df, output_path, sheet_name="月统计",
                      apply_monthly_style=True)
    return output_path


# ============================================================
# 写入 + 样式
# ============================================================
def _safe_write_excel(df: pd.DataFrame, output_path: str,
                      sheet_name="Sheet1",
                      apply_weekly_style=False,
                      apply_monthly_style=False):
    """写入Excel并捕获文件被占用的情况"""
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    except PermissionError:
        raise PermissionError(
            f"无法写入 {output_path}，请先关闭已打开的该Excel文件后重试。"
        )

    # 应用样式
    wb = load_workbook(output_path)
    ws = wb[sheet_name]
    _apply_base_style(ws)

    if apply_monthly_style:
        _apply_monthly_slash(ws, df)

    wb.save(output_path)


def _apply_base_style(ws):
    """通用样式：居中、边框、列宽自适应、首行加粗"""
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    max_row = ws.max_row
    max_col = ws.max_column

    # 列宽：按表头字符估算
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        max_len = 0
        for row in range(1, max_row + 1):
            v = ws.cell(row=row, column=col).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[letter].width = max(10, max_len * 2 + 2)

    # 单元格格式
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = center
            if row == 1:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="D9E1F2")


def _apply_monthly_slash(ws, df: pd.DataFrame):
    """
    针对月统计：多值班次(E列) / 缺班数量(F列) <= 0 时：
    清空数字，并加 左上→右下 斜线。
    df 顺序：序号,姓名,总计班次,应值班次,多值班次,缺班数量
    """
    thin = Side(border_style="thin", color="000000")
    slash_border = Border(
        left=thin, right=thin, top=thin, bottom=thin,
        diagonal=thin, diagonalDown=True,
    )

    # 多值班次=第5列，缺班数量=第6列（表头占第1行，数据从第2行起）
    col_over = 5
    col_absence = 6
    for i, row_data in df.iterrows():
        excel_row = i + 2
        # 多值班次
        if row_data["多值班次"] <= 0:
            c = ws.cell(row=excel_row, column=col_over)
            c.value = None
            c.border = slash_border
        # 缺班数量
        if row_data["缺班数量"] <= 0:
            c = ws.cell(row=excel_row, column=col_absence)
            c.value = None
            c.border = slash_border
