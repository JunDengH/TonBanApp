# src/modules/word_generator.py
"""
月度 Word 报表生成器。

核心特性：
1. 不再依赖外部 Word 模板；直接在代码中按黄金样本
   `2026春季学期5-8周助理值班统计.docx` 的版式创建文档。
2. 支持跨月结转：
   - 总计班次 = Σ(本月实际) + 上月"多值班次" + 本月加班补录
   - 应值班次 = Σ(本月应值) + 上月"缺班数量"
3. 序号列使用真正的 Word 自动编号，而不是手写数字。
4. ≤0 的"多值班次"/"缺班数量"格子：清空文本并注入 tl2br 斜线占位。
5. 排序：多值降 -> 缺班升 -> 拼音升。
6. 人员流动：以最新总名单为唯一基准；离职者历史丢弃；新进者上月数据=0。
"""
import os

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from src.modules.excel_generator import aggregate_weekly_files
from src.modules.word_parser import parse_previous_month_word
from src.utils.helpers import pinyin_key


# ============================================================
# 版式规格：解析自 2026春季学期5-8周助理值班统计.docx
# ============================================================
COL_NO = 0
COL_NAME = 1
COL_TOTAL = 2
COL_SHOULD = 3
COL_OVER = 4
COL_ABSENCE = 5

HEADERS = ["序号", "姓名", "总计班次", "应值班次", "多值班次", "缺班数量"]

PAGE_SIZE = {"w": "11906", "h": "16838"}
PAGE_MARGINS = {
    "top": "1440",
    "right": "1800",
    "bottom": "1440",
    "left": "1800",
    "header": "851",
    "footer": "992",
    "gutter": "0",
}
TABLE_GRID_WIDTHS = ["1086", "1663", "1708", "1715", "1774", "1778"]
HEADER_CELL_WIDTHS_PCT = ["559", "855", "878", "882", "912", "911"]

TITLE_ROW_HEIGHT = "983"
HEADER_ROW_HEIGHT = "519"
DATA_ROW_HEIGHT = "561"

TITLE_FONT_SIZE = "48"
BODY_FONT_SIZE = "24"
FONT_NAME = "宋体"


# ============================================================
# 对外主函数
# ============================================================
def generate_monthly_word(
    weekly_excel_paths: list,
    total_names: list,
    prev_word_path: str | None,
    output_path: str,
    title_text: str | None = None,
    overtime_shifts: dict | None = None,
) -> str:
    """
    生成月度 Word 报表。

    参数:
        weekly_excel_paths: 本月周统计 Excel 路径列表，至少 1 份
        total_names: 当前总名单（来自 DataManager）
        prev_word_path: 上月月度 Word 路径；None 或 "" 表示首次运行
        output_path: 输出 .docx 路径
        title_text: 可选大标题；None 时使用通用标题
        overtime_shifts: 月度加班补录，如 {"张三": 2}

    返回:
        实际写入的文件路径
    """
    if not total_names:
        raise ValueError("总名单为空，请先在基础设置中导入。")

    cur_actual, cur_should = aggregate_weekly_files(weekly_excel_paths, total_names)

    prev_data = {}
    if prev_word_path:
        prev_data = parse_previous_month_word(prev_word_path)

    rows = []
    for name in total_names:
        prev_over = prev_data.get(name, {}).get("多值班次", 0)
        prev_absence = prev_data.get(name, {}).get("缺班数量", 0)
        try:
            extra = int((overtime_shifts or {}).get(name, 0))
        except (ValueError, TypeError):
            extra = 0
        if extra < 0:
            extra = 0

        total_cnt = cur_actual.get(name, 0) + prev_over + extra
        should_cnt = cur_should.get(name, 0) + prev_absence
        over = total_cnt - should_cnt
        absence = should_cnt - total_cnt
        rows.append({
            "姓名": name,
            "总计班次": total_cnt,
            "应值班次": should_cnt,
            "多值班次": over,
            "缺班数量": absence,
        })

    rows.sort(key=lambda r: (-r["多值班次"], r["缺班数量"], pinyin_key(r["姓名"])))

    doc = _build_monthly_document(rows, title_text or "助理值班统计")

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    try:
        doc.save(output_path)
    except PermissionError:
        raise PermissionError(
            f"无法写入 {output_path}，请先关闭已打开的该 Word 文件后重试。"
        )
    return output_path


# ============================================================
# 文档构建
# ============================================================
def _build_monthly_document(rows: list[dict], title_text: str):
    """从空白文档按内置版式创建月统计 Word。"""
    doc = Document()
    _apply_page_setup(doc)
    _install_numbering_definition(doc)

    table = doc.add_table(rows=2 + len(rows), cols=len(HEADERS))
    _apply_table_properties(table)
    _build_title_row(table.rows[0], title_text)
    _build_header_row(table.rows[1])

    for idx, data in enumerate(rows, start=2):
        _build_data_row(table.rows[idx], data)

    return doc


def _apply_page_setup(doc):
    """设置 A4 页面、页边距和文档网格。"""
    sect_pr = doc.sections[0]._sectPr

    pg_sz = _get_or_add(sect_pr, "w:pgSz")
    pg_sz.set(qn("w:w"), PAGE_SIZE["w"])
    pg_sz.set(qn("w:h"), PAGE_SIZE["h"])

    pg_mar = _get_or_add(sect_pr, "w:pgMar")
    for key, value in PAGE_MARGINS.items():
        pg_mar.set(qn(f"w:{key}"), value)

    pg_borders = _get_or_add(sect_pr, "w:pgBorders")
    _clear_children(pg_borders)
    for side in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "none")
        border.set(qn("w:sz"), "0")
        border.set(qn("w:space"), "0")
        pg_borders.append(border)

    cols = _get_or_add(sect_pr, "w:cols")
    cols.set(qn("w:space"), "425")
    cols.set(qn("w:num"), "1")

    doc_grid = _get_or_add(sect_pr, "w:docGrid")
    doc_grid.set(qn("w:type"), "lines")
    doc_grid.set(qn("w:linePitch"), "312")
    doc_grid.set(qn("w:charSpace"), "0")


def _install_numbering_definition(doc):
    """
    安装与黄金样本一致的 Word 自动编号定义。

    黄金样本使用 numId=1、abstractNumId=0、decimal、lvlText=%1。
    python-docx 默认文档可能自带编号定义，这里清空后重建，避免隐藏模板状态影响输出。
    """
    numbering = doc.part.numbering_part.element
    _clear_children(numbering)

    abstract_num = OxmlElement("w:abstractNum")
    abstract_num.set(qn("w:abstractNumId"), "0")
    _append_leaf(abstract_num, "w:nsid", {"w:val": "25DA516C"})
    _append_leaf(abstract_num, "w:multiLevelType", {"w:val": "multilevel"})
    _append_leaf(abstract_num, "w:tmpl", {"w:val": "25DA516C"})

    level_specs = [
        ("0", "decimal", "%1", "center", {"w:left": "0", "w:firstLine": "175"}, "7"),
        ("1", "lowerLetter", "%2)", "left", {"w:left": "439", "w:hanging": "420"}, None),
        ("2", "lowerRoman", "%3.", "right", {"w:left": "859", "w:hanging": "420"}, None),
        ("3", "decimal", "%4.", "left", {"w:left": "1279", "w:hanging": "420"}, None),
        ("4", "lowerLetter", "%5)", "left", {"w:left": "1699", "w:hanging": "420"}, None),
        ("5", "lowerRoman", "%6.", "right", {"w:left": "2119", "w:hanging": "420"}, None),
        ("6", "decimal", "%7.", "left", {"w:left": "2539", "w:hanging": "420"}, None),
        ("7", "lowerLetter", "%8)", "left", {"w:left": "2959", "w:hanging": "420"}, None),
        ("8", "lowerRoman", "%9.", "right", {"w:left": "3379", "w:hanging": "420"}, None),
    ]
    for ilvl, num_fmt, lvl_text, lvl_jc, indent_attrs, p_style in level_specs:
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), ilvl)
        lvl.set(qn("w:tentative"), "0")
        _append_leaf(lvl, "w:start", {"w:val": "1"})
        _append_leaf(lvl, "w:numFmt", {"w:val": num_fmt})
        if p_style is not None:
            _append_leaf(lvl, "w:pStyle", {"w:val": p_style})
        _append_leaf(lvl, "w:lvlText", {"w:val": lvl_text})
        _append_leaf(lvl, "w:lvlJc", {"w:val": lvl_jc})

        p_pr = OxmlElement("w:pPr")
        _append_leaf(p_pr, "w:ind", indent_attrs)
        lvl.append(p_pr)

        r_pr = OxmlElement("w:rPr")
        _append_leaf(r_pr, "w:rFonts", {"w:hint": "eastAsia"})
        lvl.append(r_pr)
        abstract_num.append(lvl)

    numbering.append(abstract_num)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), "1")
    _append_leaf(num, "w:abstractNumId", {"w:val": "0"})
    numbering.append(num)


def _apply_table_properties(table):
    """设置表格固定布局、宽度、居中、单元格边距和网格。"""
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    _clear_children(tbl_pr)
    _append_leaf(tbl_pr, "w:tblStyle", {"w:val": "3"})
    _append_leaf(tbl_pr, "w:tblW", {"w:w": "5705", "w:type": "pct"})
    _append_leaf(tbl_pr, "w:jc", {"w:val": "center"})
    _append_leaf(tbl_pr, "w:tblLayout", {"w:type": "fixed"})

    cell_mar = OxmlElement("w:tblCellMar")
    for side, width in (("top", "0"), ("left", "108"), ("bottom", "0"), ("right", "108")):
        _append_leaf(cell_mar, f"w:{side}", {"w:w": width, "w:type": "dxa"})
    tbl_pr.append(cell_mar)

    old_grid = tbl.tblGrid
    if old_grid is not None:
        tbl.remove(old_grid)
    grid = OxmlElement("w:tblGrid")
    for width in TABLE_GRID_WIDTHS:
        _append_leaf(grid, "w:gridCol", {"w:w": width})
    tbl.insert(1, grid)


def _build_title_row(row, title_text: str):
    tr = row._tr
    _set_row_properties(tr, TITLE_ROW_HEIGHT)

    tcs = _iter_tcs(tr)
    first_tc = tcs[0]
    for extra_tc in tcs[1:]:
        tr.remove(extra_tc)

    _set_cell_properties(
        first_tc,
        width="5000",
        width_type="pct",
        grid_span="6",
    )
    _set_cell_text(first_tc, title_text, size=TITLE_FONT_SIZE, align="center")


def _build_header_row(row):
    tr = row._tr
    _set_row_properties(tr, HEADER_ROW_HEIGHT)
    for idx, tc in enumerate(_iter_tcs(tr)):
        _set_cell_properties(
            tc,
            width=HEADER_CELL_WIDTHS_PCT[idx],
            width_type="pct",
        )
        _set_cell_text(tc, HEADERS[idx], size=BODY_FONT_SIZE, align="center")


def _build_data_row(row, data: dict):
    tr = row._tr
    _set_row_properties(tr, DATA_ROW_HEIGHT)
    values = [
        "",
        data["姓名"],
        str(data["总计班次"]),
        str(data["应值班次"]),
        "" if data["多值班次"] <= 0 else str(data["多值班次"]),
        "" if data["缺班数量"] <= 0 else str(data["缺班数量"]),
    ]

    for idx, tc in enumerate(_iter_tcs(tr)):
        width = TABLE_GRID_WIDTHS[idx]
        width_type = "dxa"
        if idx in (COL_NO, COL_ABSENCE):
            # 黄金样本的首尾列在单元格层保留 pct，网格仍使用 dxa。
            width = HEADER_CELL_WIDTHS_PCT[idx]
            width_type = "pct"
        _set_cell_properties(
            tc,
            width=width,
            width_type=width_type,
            diagonal=(idx == COL_OVER and data["多值班次"] <= 0)
            or (idx == COL_ABSENCE and data["缺班数量"] <= 0),
        )

        if idx == COL_NO:
            _set_numbered_cell(tc)
        else:
            _set_cell_text(tc, values[idx], size=BODY_FONT_SIZE, align="center")


# ============================================================
# XML 细节工具
# ============================================================
def _iter_tcs(tr_element):
    return tr_element.findall(qn("w:tc"))


def _clear_children(element):
    for child in list(element):
        element.remove(child)


def _get_or_add(parent, tag):
    child = parent.find(qn(tag))
    if child is None:
        child = OxmlElement(tag)
        parent.append(child)
    return child


def _append_leaf(parent, tag, attrs: dict):
    child = OxmlElement(tag)
    for key, value in attrs.items():
        child.set(qn(key), value)
    parent.append(child)
    return child


def _set_row_properties(tr_element, height: str):
    tr_pr = _get_or_add(tr_element, "w:trPr")
    _clear_children(tr_pr)
    _append_leaf(tr_pr, "w:trHeight", {"w:val": height, "w:hRule": "atLeast"})
    _append_leaf(tr_pr, "w:jc", {"w:val": "center"})


def _set_cell_properties(
    tc_element,
    width: str,
    width_type: str,
    grid_span: str | None = None,
    diagonal: bool = False,
):
    tc_pr = _get_or_add(tc_element, "w:tcPr")
    _clear_children(tc_pr)
    _append_leaf(tc_pr, "w:tcW", {"w:w": width, "w:type": width_type})
    if grid_span is not None:
        _append_leaf(tc_pr, "w:gridSpan", {"w:val": grid_span})

    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        _append_leaf(borders, f"w:{side}", {
            "w:val": "single",
            "w:color": "000000",
            "w:sz": "4",
            "w:space": "0",
        })
    if diagonal:
        _append_leaf(borders, "w:tl2br", {
            "w:val": "single",
            "w:sz": "4",
            "w:color": "auto",
        })
    tc_pr.append(borders)
    _append_leaf(tc_pr, "w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": "auto"})
    tc_pr.append(OxmlElement("w:noWrap"))
    _append_leaf(tc_pr, "w:vAlign", {"w:val": "center"})


def _set_cell_text(tc_element, text: str, size: str, align: str = "center"):
    _clear_paragraphs(tc_element)
    p = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    _append_leaf(p_pr, "w:jc", {"w:val": align})
    _append_leaf(p_pr, "w:textAlignment", {"w:val": "center"})
    p_r_pr = OxmlElement("w:rPr")
    _append_run_style(p_r_pr, size)
    p_pr.append(p_r_pr)
    p.append(p_pr)

    r = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    _append_run_style(r_pr, size)
    r.append(r_pr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text if text is not None else ""
    r.append(t)
    p.append(r)
    tc_element.append(p)


def _set_numbered_cell(tc_element):
    _clear_paragraphs(tc_element)
    p = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    p_r_pr = OxmlElement("w:rPr")
    _append_leaf(p_r_pr, "w:lang", {"w:bidi": "ar"})
    _append_leaf(p_r_pr, "w:sz", {"w:val": BODY_FONT_SIZE})
    _append_leaf(p_r_pr, "w:szCs", {"w:val": BODY_FONT_SIZE})
    p_pr.append(p_r_pr)

    num_pr = OxmlElement("w:numPr")
    _append_leaf(num_pr, "w:ilvl", {"w:val": "0"})
    _append_leaf(num_pr, "w:numId", {"w:val": "1"})
    p_pr.append(num_pr)
    p.append(p_pr)
    tc_element.append(p)


def _clear_paragraphs(tc_element):
    for p in list(tc_element.findall(qn("w:p"))):
        tc_element.remove(p)


def _append_run_style(r_pr, size: str):
    _append_leaf(r_pr, "w:rFonts", {
        "w:hint": "eastAsia",
        "w:ascii": FONT_NAME,
        "w:hAnsi": FONT_NAME,
        "w:eastAsia": FONT_NAME,
        "w:cs": FONT_NAME,
    })
    _append_leaf(r_pr, "w:i", {"w:val": "0"})
    _append_leaf(r_pr, "w:iCs", {"w:val": "0"})
    _append_leaf(r_pr, "w:color", {"w:val": "000000"})
    _append_leaf(r_pr, "w:sz", {"w:val": size})
    _append_leaf(r_pr, "w:szCs", {"w:val": size})
    _append_leaf(r_pr, "w:u", {"w:val": "none"})
    _append_leaf(r_pr, "w:lang", {"w:val": "en-US", "w:eastAsia": "zh-CN", "w:bidi": "ar"})
