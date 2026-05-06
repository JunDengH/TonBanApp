# src/modules/word_parser.py
"""
Word解析模块：
1. 解析总名单Word（按空格分隔提取姓名）
2. 解析周统计Word（按周一~周日提取每日值班名单）
"""
import re
from docx import Document
from docx.oxml.ns import qn


# 中文姓名匹配：2-4个汉字
NAME_REGEX = re.compile(r"[\u4e00-\u9fa5]{2,4}")

# 月度Word表头列名（反向解析时用于匹配列索引）
MONTHLY_HEADER_REQUIRED = ("姓名", "多值班次", "缺班数量")
# 识别星期标题：周一、星期一、周1 等
WEEKDAY_PATTERNS = {
    "周一": ["周一", "星期一", "礼拜一", "周1"],
    "周二": ["周二", "星期二", "礼拜二", "周2"],
    "周三": ["周三", "星期三", "礼拜三", "周3"],
    "周四": ["周四", "星期四", "礼拜四", "周4"],
    "周五": ["周五", "星期五", "礼拜五", "周5"],
    "周六": ["周六", "星期六", "礼拜六", "周6"],
    "周日": ["周日", "星期日", "星期天", "礼拜日", "礼拜天", "周7"],
}


def _iter_all_text(doc: Document):
    """遍历文档：正文段落 + 所有表格单元格，返回文本行列表"""
    lines = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if text:
            lines.append(text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    text = p.text.strip()
                    if text:
                        lines.append(text)
    return lines


def parse_total_name_list(docx_path: str) -> list:
    """
    解析总名单：提取所有中文姓名
    要求：Word中姓名以空格分隔
    """
    doc = Document(docx_path)
    lines = _iter_all_text(doc)

    names = []
    seen = set()
    for line in lines:
        # 先按空白分割，再用正则提取中文姓名（防止混入课节等字符）
        for token in re.split(r"\s+", line):
            for name in NAME_REGEX.findall(token):
                if name not in seen:
                    seen.add(name)
                    names.append(name)
    return names


def _match_weekday(text: str) -> str:
    """判断一段文本是否是某一天的标题行；返回 周一~周日 或 None"""
    for key, aliases in WEEKDAY_PATTERNS.items():
        for alias in aliases:
            if alias in text:
                return key
    return None


def parse_weekly_schedule(docx_path: str, total_names: list) -> dict:
    """
    解析周统计Word，返回：
    {
        "周一": ["张三", "李四", "张三", ...],   # 保留重复，表示多次排班
        "周二": [...],
        ...
    }
    说明：
        - 按出现顺序遇到周X标题，其后（直到下一个周X标题）的人名都归入该天；
        - 若文档通过表格呈现，会将所有单元格文本按顺序纳入；
        - 只保留出现在总名单中的姓名，自动过滤"1-2节"等课节干扰。
    """
    doc = Document(docx_path)
    lines = _iter_all_text(doc)

    name_set = set(total_names)
    result = {k: [] for k in WEEKDAY_PATTERNS.keys()}
    current_day = None

    for line in lines:
        weekday = _match_weekday(line)
        if weekday is not None:
            current_day = weekday
            # 标题行也可能同时包含姓名，继续解析该行剩余内容
        if current_day is None:
            continue

        # 提取中文姓名并按总名单过滤
        for name in NAME_REGEX.findall(line):
            if name in name_set:
                result[current_day].append(name)

    return result


def count_actual_shifts(weekly_schedule: dict, total_names: list) -> dict:
    """统计每位助理在一周内的实际出现次数"""
    counter = {n: 0 for n in total_names}
    for day, names in weekly_schedule.items():
        for n in names:
            if n in counter:
                counter[n] += 1
    return counter


# ============================================================
# 上月月度 Word 反向解析（需求3 跨月结转使用）
# ============================================================
def _cell_has_diagonal(cell) -> bool:
    """判断单元格是否含左上→右下斜线（w:tl2br）"""
    tcPr = cell._tc.find(qn("w:tcPr"))
    if tcPr is None:
        return False
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        return False
    return borders.find(qn("w:tl2br")) is not None


def _safe_int(text: str) -> int:
    """将单元格文本稳健转为整数；非数字一律按 0 处理"""
    if text is None:
        return 0
    s = str(text).strip()
    if not s:
        return 0
    # 去掉可能的空白与全角符号
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0


def parse_previous_month_word(docx_path: str) -> dict:
    """
    解析上月月度统计 Word，返回每位助理的"多值班次""缺班数量"。

    反向解析规则（需求3 第4条 防篡改与占位）：
      - 单元格含 w:tl2br 斜线 -> 视为 0
      - 单元格文本为空         -> 视为 0
      - 单元格文本非整数       -> 视为 0

    返回:
        {姓名: {"多值班次": int, "缺班数量": int}, ...}

    异常:
        ValueError: 文档中不存在含"姓名""多值班次""缺班数量"三列的表格时抛出
    """
    doc = Document(docx_path)
    if not doc.tables:
        raise ValueError("上月Word未发现任何表格，无法反向解析。")

    # 找到首个同时含 3 个必需列的表格，并定位其表头行
    target_table = None
    header_row_idx = -1
    col_map = {}
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            cells_text = [c.text.strip() for c in row.cells]
            if all(h in cells_text for h in MONTHLY_HEADER_REQUIRED):
                # 记录列索引（取第一次出现的索引）
                col_map = {h: cells_text.index(h) for h in MONTHLY_HEADER_REQUIRED}
                target_table = table
                header_row_idx = ri
                break
        if target_table is not None:
            break

    if target_table is None:
        raise ValueError(
            "上月Word表头不符合规范，需要至少包含【姓名 / 多值班次 / 缺班数量】三列。"
        )

    name_col = col_map["姓名"]
    over_col = col_map["多值班次"]
    absence_col = col_map["缺班数量"]

    result = {}
    for row in target_table.rows[header_row_idx + 1:]:
        cells = row.cells
        if len(cells) <= max(name_col, over_col, absence_col):
            continue
        name = cells[name_col].text.strip()
        if not name:
            continue
        # 过滤非人名（避免把标题行等也收进来）
        if not NAME_REGEX.fullmatch(name):
            continue

        over_cell = cells[over_col]
        absence_cell = cells[absence_col]

        over_val = 0 if _cell_has_diagonal(over_cell) else _safe_int(over_cell.text)
        absence_val = 0 if _cell_has_diagonal(absence_cell) else _safe_int(absence_cell.text)

        result[name] = {
            "多值班次": over_val,
            "缺班数量": absence_val,
        }
    return result