# src/modules/word_generator.py
"""
月度 Word 报表生成器（需求3 第（三）项）。

核心特性：
1. 采用"模板克隆法"复刻外观：以 assets/monthly_template.docx 为基底，
   只替换标题与数据行，保留原模板的字体 / 字号 / 对齐 / 边框等全部样式。
2. 支持跨月结转：
   - 总计班次 = Σ(本月实际) + 上月"多值班次"
   - 应值班次 = Σ(本月应值) + 上月"缺班数量"
3. ≤0 的"多值班次"/"缺班数量"格子：清空文本并注入 tl2br 斜线占位。
4. 排序：多值降 -> 缺班升 -> 拼音升。
5. 人员流动：以最新总名单为唯一基准；离职者历史丢弃；新进者上月数据=0。
"""
import copy
import os
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from src.modules.excel_generator import aggregate_weekly_files
from src.modules.word_parser import parse_previous_month_word
from src.utils.helpers import pinyin_key


# ============================================================
# 常量
# ============================================================
# 模板中的表格列索引（与样本一致）
COL_NO = 0
COL_NAME = 1
COL_TOTAL = 2
COL_SHOULD = 3
COL_OVER = 4
COL_ABSENCE = 5

# 模板结构约定（与 assets/monthly_template.docx 一致）
TITLE_ROW_INDEX = 0   # 大标题所在的表格行
HEADER_ROW_INDEX = 1  # 表头行
SAMPLE_DATA_ROW_INDEX = 2  # 作为样式样本的数据行（克隆时依赖此行的 rPr/tcPr）


# ============================================================
# 模板路径解析（兼容 PyInstaller 打包）
# ============================================================
def _resolve_template_path() -> Path:
    """
    解析内置模板文件 assets/monthly_template.docx 的绝对路径。
    同时兼容开发环境与 PyInstaller 一体打包 (sys._MEIPASS)。
    """
    candidate_roots = []
    # PyInstaller 运行时
    if hasattr(sys, "_MEIPASS"):
        candidate_roots.append(Path(sys._MEIPASS))
    # 源码运行时：项目根目录
    candidate_roots.append(Path(__file__).resolve().parent.parent.parent)
    for root in candidate_roots:
        p = root / "assets" / "monthly_template.docx"
        if p.exists():
            return p
    raise FileNotFoundError(
        "未找到 assets/monthly_template.docx，请确认模板文件已随程序发布。"
    )


# ============================================================
# 对外主函数
# ============================================================
def generate_monthly_word(
    weekly_excel_paths: list,
    total_names: list,
    prev_word_path: str | None,
    output_path: str,
    title_text: str | None = None,
) -> str:
    """
    生成月度 Word 报表。

    参数:
        weekly_excel_paths: 本月 4 份周统计 Excel 路径
        total_names: 当前总名单（来自 DataManager）
        prev_word_path: 上月月度 Word 路径；None 或 "" 表示首次运行
        output_path: 输出 .docx 路径
        title_text: 可选大标题；None 则沿用模板已有标题

    返回:
        实际写入的文件路径
    """
    if not total_names:
        raise ValueError("总名单为空，请先在基础设置中导入。")

    # 1) 本月周汇总
    cur_actual, cur_should = aggregate_weekly_files(weekly_excel_paths, total_names)

    # 2) 上月历史（可选）
    prev_data = {}
    if prev_word_path:
        prev_data = parse_previous_month_word(prev_word_path)

    # 3) 合并结转
    rows = []
    for name in total_names:
        prev_over = prev_data.get(name, {}).get("多值班次", 0)     # 离职者不会匹配到；新进者默认 0
        prev_absence = prev_data.get(name, {}).get("缺班数量", 0)
        total_cnt = cur_actual.get(name, 0) + prev_over             # 本轮总计班次
        should_cnt = cur_should.get(name, 0) + prev_absence         # 本轮应值班次
        over = total_cnt - should_cnt                                # 多值
        absence = should_cnt - total_cnt                             # 缺班
        rows.append({
            "姓名": name,
            "总计班次": total_cnt,
            "应值班次": should_cnt,
            "多值班次": over,
            "缺班数量": absence,
        })

    # 4) 排序：多值降 -> 缺班升 -> 拼音升
    rows.sort(key=lambda r: (-r["多值班次"], r["缺班数量"], pinyin_key(r["姓名"])))

    # 5) 克隆模板并填充
    template_path = _resolve_template_path()
    doc = Document(str(template_path))
    table = doc.tables[0]

    if title_text is not None:
        _set_title_cell_text(table, TITLE_ROW_INDEX, title_text)

    sample_row = table.rows[SAMPLE_DATA_ROW_INDEX]._tr     # <w:tr>
    sample_row_xml = copy.deepcopy(sample_row)             # 保留样式
    # 清空原有数据行（HEADER_ROW_INDEX+1 及之后）
    _clear_template_data_rows(table, keep_first_n_rows=HEADER_ROW_INDEX + 1)

    # 逐行追加
    tbl_el = table._tbl
    for i, r in enumerate(rows, start=1):
        new_tr = copy.deepcopy(sample_row_xml)
        _fill_row(new_tr, {
            COL_NO: "",
            COL_NAME: r["姓名"],
            COL_TOTAL: str(r["总计班次"]),
            COL_SHOULD: str(r["应值班次"]),
            COL_OVER: "" if r["多值班次"] <= 0 else str(r["多值班次"]),
            COL_ABSENCE: "" if r["缺班数量"] <= 0 else str(r["缺班数量"]),
        })
        _set_word_list_numbering_for_first_col(new_tr)
        # 对 ≤0 的列设置斜线占位，非 ≤0 的列清除可能遗留的斜线
        _set_cell_diagonal(new_tr, COL_OVER, enable=(r["多值班次"] <= 0))
        _set_cell_diagonal(new_tr, COL_ABSENCE, enable=(r["缺班数量"] <= 0))
        tbl_el.append(new_tr)

    # 6) 保存
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    try:
        doc.save(output_path)
    except PermissionError:
        raise PermissionError(
            f"无法写入 {output_path}，请先关闭已打开的该 Word 文件后重试。"
        )
    return output_path


# ============================================================
# 内部：表格操作辅助函数
# ============================================================
def _clear_template_data_rows(table, keep_first_n_rows: int):
    """删除表格中保留数之后的所有行"""
    tbl_el = table._tbl
    trs = tbl_el.findall(qn("w:tr"))
    for tr in trs[keep_first_n_rows:]:
        tbl_el.remove(tr)


def _iter_tcs(tr_element):
    """按顺序返回 <w:tc> 元素列表"""
    return tr_element.findall(qn("w:tc"))


def _set_cell_text_preserving_style(tc_element, text: str, fallback_rPr=None):
    """
    设置单元格文本，保留原 run 的字体/字号样式：
    - 保留第一个 <w:p> 及其中第一个 <w:r> 的 <w:rPr>
    - 清空文本并写入新文本
    - 删除其余多余段落与 run
    """
    # 清除所有 <w:p> 之外可能的内容
    ps = tc_element.findall(qn("w:p"))
    if not ps:
        # 理论上不会发生；兜底
        p = OxmlElement("w:p")
        tc_element.append(p)
        ps = [p]

    # 只保留第一个段落
    for p in ps[1:]:
        tc_element.remove(p)
    p0 = ps[0]

    # 找到第一个 run 的 rPr 作为样式模板
    runs = p0.findall(qn("w:r"))
    template_rPr = None
    if runs:
        first_rPr = runs[0].find(qn("w:rPr"))
        if first_rPr is not None:
            template_rPr = copy.deepcopy(first_rPr)
        # 移除所有 run
        for r in runs:
            p0.remove(r)

    # 若当前格原本无 run 样式（常见于序号列/缺班列空格），使用行级回退样式
    if template_rPr is None and fallback_rPr is not None:
        template_rPr = copy.deepcopy(fallback_rPr)

    # 防止段落样式携带自动编号，导致序号列出现"双序号"
    pPr = p0.find(qn("w:pPr"))
    if pPr is not None:
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            pPr.remove(numPr)
        # 有些模板通过 pStyle 绑定了编号样式（即使没有 numPr 也会显示序号）
        # 数据写入时统一移除该段落样式，保留对齐(jc)等局部格式。
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is not None:
            pPr.remove(pStyle)

    # 新建 run
    new_r = OxmlElement("w:r")
    if template_rPr is not None:
        new_r.append(template_rPr)
    new_t = OxmlElement("w:t")
    new_t.set(qn("xml:space"), "preserve")
    new_t.text = text if text is not None else ""
    new_r.append(new_t)
    p0.append(new_r)


def _find_row_fallback_rPr(tr_element):
    """在一行中寻找一个可复用的 run 样式，供空白单元格回退使用。"""
    for tc in _iter_tcs(tr_element):
        for p in tc.findall(qn("w:p")):
            for r in p.findall(qn("w:r")):
                rPr = r.find(qn("w:rPr"))
                if rPr is not None:
                    return copy.deepcopy(rPr)
    return None


def _fill_row(tr_element, col_to_text: dict):
    """按列索引填充一行单元格文本"""
    fallback_rPr = _find_row_fallback_rPr(tr_element)
    tcs = _iter_tcs(tr_element)
    for col_idx, text in col_to_text.items():
        if col_idx < len(tcs):
            _set_cell_text_preserving_style(
                tcs[col_idx],
                text,
                fallback_rPr=fallback_rPr,
            )


def _set_word_list_numbering_for_first_col(tr_element):
    """将第1列(序号列)设置为Word自动编号，并清空手填文本。"""
    tcs = _iter_tcs(tr_element)
    if not tcs:
        return
    tc = tcs[COL_NO]
    p = tc.find(qn("w:p"))
    if p is None:
        p = OxmlElement("w:p")
        tc.append(p)

    # 清空已有run文本（避免自动编号+手写数字同时出现）
    for r in p.findall(qn("w:r")):
        p.remove(r)

    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)

    # 清理 pStyle，防止样式冲突
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is not None:
        pPr.remove(pStyle)

    # 采用模板中已存在的自动编号定义（numId=1，ilvl=0）
    # 模板的 numId=1 对应 decimal + "%1"（不带点）
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        numPr = OxmlElement("w:numPr")
        pPr.append(numPr)

    ilvl = numPr.find(qn("w:ilvl"))
    if ilvl is None:
        ilvl = OxmlElement("w:ilvl")
        numPr.append(ilvl)
    ilvl.set(qn("w:val"), "0")

    numId = numPr.find(qn("w:numId"))
    if numId is None:
        numId = OxmlElement("w:numId")
        numPr.append(numId)
    numId.set(qn("w:val"), "1")

    # 强制序号显示为小四（12pt，w:sz=24）
    p_rPr = pPr.find(qn("w:rPr"))
    if p_rPr is None:
        p_rPr = OxmlElement("w:rPr")
        pPr.append(p_rPr)
    sz = p_rPr.find(qn("w:sz"))
    if sz is None:
        sz = OxmlElement("w:sz")
        p_rPr.append(sz)
    sz.set(qn("w:val"), "24")
    sz_cs = p_rPr.find(qn("w:szCs"))
    if sz_cs is None:
        sz_cs = OxmlElement("w:szCs")
        p_rPr.append(sz_cs)
    sz_cs.set(qn("w:val"), "24")


def _get_or_create_tcBorders(tc_element):
    """返回/创建 <w:tcPr>/<w:tcBorders>"""
    tcPr = tc_element.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        # tcPr 需要放在 tc 的最前面
        tc_element.insert(0, tcPr)
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    return borders


def _set_cell_diagonal(tr_element, col_idx: int, enable: bool):
    """为某列单元格设置/清除 tl2br 斜线"""
    tcs = _iter_tcs(tr_element)
    if col_idx >= len(tcs):
        return
    tc = tcs[col_idx]
    borders = _get_or_create_tcBorders(tc)
    existing = borders.find(qn("w:tl2br"))
    if enable:
        if existing is None:
            tl2br = OxmlElement("w:tl2br")
            tl2br.set(qn("w:val"), "single")
            tl2br.set(qn("w:sz"), "4")
            tl2br.set(qn("w:color"), "000000")
            borders.append(tl2br)
    else:
        if existing is not None:
            borders.remove(existing)


def _set_title_cell_text(table, title_row_idx: int, text: str):
    """替换标题合并单元格的文本（保留其字体样式）"""
    if title_row_idx >= len(table.rows):
        return
    tr = table.rows[title_row_idx]._tr
    tcs = _iter_tcs(tr)
    if not tcs:
        return
    # 合并标题通常只在首个 tc 写文本，其余 tc 含 <w:vMerge> / 仅做 gridSpan
    _set_cell_text_preserving_style(tcs[0], text)
