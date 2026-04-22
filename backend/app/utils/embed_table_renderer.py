"""
嵌入组件 —— 表格渲染器

支持的 payload 格式
-------------------
{
    "headers": ["姓名", "部门", "职位"],        # 表头列（为空则无表头行）
    "rows": [                                   # 数据行（字符串 or 数字）
        ["张三", "研发部", "工程师"],
        ["李四", "市场部", "经理"]
    ],
    "caption": "表1 人员信息",                  # 可选：表格标题（插入在表格上方）
    "has_header": true,                         # 默认 true；false 时 headers 作普通行
    "col_widths": [3.0, 4.0, 3.0],             # 可选：各列宽度（cm），null = 自适应
    "merge_cells": [                            # 可选：合并单元格 [行, 起始列, 结束列]
        [0, 0, 2]                               # 第 0 行 col0~col2 合并
    ],
    "style": {
        "header_bg":        "4472C4",           # 表头背景色（hex，不含 #）
        "header_color":     "FFFFFF",           # 表头字体颜色
        "border_color":     "BFBFBF",           # 边框颜色
        "font_name":        "仿宋",
        "header_font_size": 10,                 # pt
        "body_font_size":   9,                  # pt
        "alignment":        "center",           # left / center / right（表格整体居中）
        "cell_alignment":   "center"            # 单元格内文字对齐
    }
}
"""

from __future__ import annotations

import html
import logging
import time
from copy import deepcopy
from typing import Any, Dict, List, Optional

from app.utils.embed_marker import (
    EmbedSpec,
    TYPE_TABLE,
    register_docx_renderer,
    register_html_renderer,
)

logger = logging.getLogger(__name__)

# 大表格按 _PROGRESS_ROW_STEP 行打印一次进度
_PROGRESS_ROW_STEP = 200

# ─── 默认样式 ──────────────────────────────────────────────────────────────────

_DEFAULT_STYLE: Dict[str, Any] = {
    "header_bg":        "4472C4",
    "header_color":     "FFFFFF",
    "border_color":     "BFBFBF",
    "font_name":        "仿宋",
    "header_font_size": 10,
    "body_font_size":   9,
    "alignment":        "center",
    "cell_alignment":   "center",
}


def _merge_style(user: Optional[Dict]) -> Dict[str, Any]:
    if not user:
        return dict(_DEFAULT_STYLE)
    return {**_DEFAULT_STYLE, **user}


# ─── DOCX 通用工具（延迟导入 docx，运行时才加载） ─────────────────────────────

def _hex_to_rgb(hex_color: str):
    """延迟导入 RGBColor，避免模块加载时依赖 docx。"""
    from docx.shared import RGBColor
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _align_enum(alignment: str):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    return {
        "left":   WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right":  WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(str(alignment).lower(), WD_ALIGN_PARAGRAPH.CENTER)


def _set_cell_border(cell, color: str) -> None:
    """给单元格四边设置单线边框。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color.upper().lstrip("#"))
        tcBorders.append(el)
    tcPr.append(tcBorders)


def _set_cell_shading(cell, fill_color: str) -> None:
    """设置单元格背景色。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  fill_color.upper().lstrip("#"))
    tcPr.append(shd)


def _write_cell(
    cell,
    text: str,
    font_name: str,
    font_size: int,
    color: str,
    bold: bool,
    align,
    bg: Optional[str] = None,
    border_color: str = "BFBFBF",
) -> None:
    """写入单元格文字 + 样式。（兼容老调用方；内部走慢路径）"""
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.shared import Pt

    _set_cell_border(cell, border_color)
    if bg:
        _set_cell_shading(cell, bg)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    para = cell.paragraphs[0]
    para.alignment = align
    para.clear()
    run = para.add_run(str(text))
    run.font.name   = font_name
    run.font.size   = Pt(font_size)
    run.font.bold   = bold
    run.font.color.rgb = _hex_to_rgb(color)
    # 中文字体需同时设置 eastAsia
    run._element.get_or_add_rPr().get_or_add_rFonts().set(
        qn("w:eastAsia"), font_name
    )


def _build_tc_borders_template(border_color: str):
    """预先构造一个 w:tcBorders 元素，循环里 deepcopy 即可，避免每个单元格重复创建 5 个 OxmlElement。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    color = border_color.upper().lstrip("#")
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        tcBorders.append(el)
    return tcBorders


def _build_shd_template(fill_color: str):
    """预先构造一个 w:shd 背景元素。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  fill_color.upper().lstrip("#"))
    return shd


def _fast_write_cell(
    cell,
    text: str,
    font_name: str,
    pt_size,
    rgb,
    bold: bool,
    align,
    tc_borders_tpl,
    shd_tpl=None,
    east_asia_qn=None,
    vcenter=None,
) -> None:
    """
    高频路径下使用的快速单元格写入：
      - 不重复创建 OxmlElement，统一用 deepcopy(模板)
      - 不重复解析 hex / 构造 Pt / RGBColor，全部由调用方预计算
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcPr.append(deepcopy(tc_borders_tpl))
    if shd_tpl is not None:
        tcPr.append(deepcopy(shd_tpl))
    if vcenter is not None:
        cell.vertical_alignment = vcenter

    para = cell.paragraphs[0]
    para.alignment = align
    # 新建 cell 时段落里只有一个空 run，省去 clear() 调用（其内部会遍历删除）
    run = para.add_run(text)
    font = run.font
    font.name = font_name
    font.size = pt_size
    font.bold = bold
    font.color.rgb = rgb
    # 中文字体：eastAsia
    if east_asia_qn is not None:
        run._element.get_or_add_rPr().get_or_add_rFonts().set(east_asia_qn, font_name)


# ─── DOCX 渲染器 ───────────────────────────────────────────────────────────────

def render_table_to_docx(spec: EmbedSpec, paragraph, document) -> None:
    """
    用 spec.payload 中的表格数据，在 paragraph 所在位置插入真实 Word 表格。
    插入完成后移除占位符段落。

    性能优化要点：
        - 所有样式相关对象（RGBColor/Pt/border tcPr 模板/shd 模板）在循环外
          预计算一次，循环内仅做 deepcopy + 赋值
        - 整行取 cells 后顺序遍历，避免反复 table.cell(ri, ci) 索引
        - 使用 lxml getnext/getparent，定位占位符位置不再 O(N)
        - 大表格按 _PROGRESS_ROW_STEP 行打印进度日志，便于观察
    """
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt

    t_total = time.perf_counter()

    payload = spec.payload or {}
    style   = _merge_style(payload.get("style"))

    headers: List[str]       = [str(h) for h in (payload.get("headers") or [])]
    rows:    List[List[str]] = [[str(c) for c in r] for r in (payload.get("rows") or [])]
    caption: str             = payload.get("caption") or ""
    has_header: bool         = bool(payload.get("has_header", True)) and bool(headers)
    col_widths: Optional[List[float]] = payload.get("col_widths")
    merge_instructions: List = payload.get("merge_cells") or []

    # ── 列数 ─────────────────────────────────────────────────────────────────
    col_count = len(headers) if headers else (max((len(r) for r in rows), default=1))
    if col_count == 0:
        logger.warning("render_table_to_docx: 列数为 0，跳过 embed_id=%s", spec.embed_id)
        return

    # ── 构建全部行 ────────────────────────────────────────────────────────────
    all_rows: List[List[str]] = []
    is_header_row: List[bool] = []
    if has_header:
        all_rows.append(headers)
        is_header_row.append(True)
    for r in rows:
        padded = r + [""] * max(0, col_count - len(r))
        all_rows.append(padded[:col_count])
        is_header_row.append(False)

    if not all_rows:
        logger.warning("render_table_to_docx: 无任何行数据，跳过 embed_id=%s", spec.embed_id)
        return

    total_rows = len(all_rows)
    align_para = _align_enum(style["cell_alignment"])
    font_name  = style["font_name"]

    # ── 预计算样式对象（循环内零分配） ────────────────────────────────────────
    header_pt   = Pt(int(style["header_font_size"]))
    body_pt     = Pt(int(style["body_font_size"]))
    header_rgb  = _hex_to_rgb(style["header_color"])
    body_rgb    = _hex_to_rgb("000000")
    border_tpl  = _build_tc_borders_template(style["border_color"])
    header_shd  = _build_shd_template(style["header_bg"]) if style.get("header_bg") else None
    east_asia   = qn("w:eastAsia")
    vcenter     = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # ── 占位符父节点定位（lxml O(1)）─────────────────────────────────────────
    p_elem = paragraph._element
    parent = p_elem.getparent()
    if parent is None:
        logger.warning("render_table_to_docx: 占位段落已脱离父节点 embed_id=%s", spec.embed_id)
        return
    insert_idx = parent.index(p_elem)  # lxml 原生 index，比 list(parent).index 快

    # ── 可选：在占位段落前插入标题段落 ───────────────────────────────────────
    if caption:
        cap_p = OxmlElement("w:p")
        cap_r = OxmlElement("w:r")
        cap_rPr = OxmlElement("w:rPr")
        cap_rPr.append(OxmlElement("w:b"))
        cap_r.append(cap_rPr)
        cap_t = OxmlElement("w:t")
        cap_t.text = caption
        cap_r.append(cap_t)
        cap_p.append(cap_r)
        parent.insert(insert_idx, cap_p)
        insert_idx += 1   # 占位符已后移一位

    # ── 创建 Word 表格 ────────────────────────────────────────────────────────
    t_create = time.perf_counter()
    table = document.add_table(rows=total_rows, cols=col_count)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_elem = table._tbl
    tbl_elem.getparent().remove(tbl_elem)    # 从 body 末尾摘除
    parent.insert(insert_idx, tbl_elem)       # 插到占位符所在位置
    create_cost = time.perf_counter() - t_create

    # ── 列宽 ─────────────────────────────────────────────────────────────────
    if col_widths:
        widths_cm = [Cm(float(w)) for w in col_widths[:col_count]]
        # 仅对第一行设置即可（python-docx 会通过 gridCol 影响整列），
        # 但为兼容老逻辑仍逐行设置；改用整行 cells 取出，避免 .cell(r, c) 重复索引
        for row_obj in table.rows:
            row_cells = row_obj.cells
            for ci, w in enumerate(widths_cm):
                if ci < len(row_cells):
                    row_cells[ci].width = w
        table.autofit = False
    else:
        table.autofit = True

    # ── 填充单元格（核心循环，热路径） ───────────────────────────────────────
    t_fill = time.perf_counter()
    table_rows = table.rows  # 缓存属性引用
    for ri in range(total_rows):
        row_data = all_rows[ri]
        is_hdr   = is_header_row[ri]
        if is_hdr:
            pt_size, rgb, shd = header_pt, header_rgb, header_shd
        else:
            pt_size, rgb, shd = body_pt, body_rgb, None
        row_cells = table_rows[ri].cells   # 整行一次性取出
        for ci in range(col_count):
            _fast_write_cell(
                cell           = row_cells[ci],
                text           = row_data[ci],
                font_name      = font_name,
                pt_size        = pt_size,
                rgb            = rgb,
                bold           = is_hdr,
                align          = align_para,
                tc_borders_tpl = border_tpl,
                shd_tpl        = shd,
                east_asia_qn   = east_asia,
                vcenter        = vcenter,
            )
        # 大表格分阶段进度日志
        if total_rows >= _PROGRESS_ROW_STEP and (
            (ri + 1) % _PROGRESS_ROW_STEP == 0 or ri == total_rows - 1
        ):
            elapsed = time.perf_counter() - t_fill
            done = ri + 1
            speed = done / elapsed if elapsed > 0 else 0.0
            eta = (total_rows - done) / speed if speed > 0 else 0.0
            logger.info(
                "render_table_to_docx 写入进度 embed_id=%s %d/%d 行"
                " 已耗时=%.2fs 速度=%.0f行/s ETA≈%.2fs",
                spec.embed_id, done, total_rows, elapsed, speed, eta,
            )
    fill_cost = time.perf_counter() - t_fill

    # ── 合并单元格（[行, 起始列, 结束列]） ────────────────────────────────────
    for instr in merge_instructions:
        if len(instr) != 3:
            continue
        ri, c_start, c_end = int(instr[0]), int(instr[1]), int(instr[2])
        if ri >= total_rows or c_end >= col_count:
            continue
        try:
            table.cell(ri, c_start).merge(table.cell(ri, c_end))
        except Exception as e:
            logger.warning(f"合并单元格失败 [{ri},{c_start},{c_end}]：{e}")

    # ── 移除占位符段落（此时它已被推到 insert_idx+1 位置） ────────────────────
    parent.remove(p_elem)

    total_cost = time.perf_counter() - t_total
    logger.info(
        "render_table_to_docx 完成 embed_id=%s rows=%d cols=%d"
        " 创建=%.2fs 填充=%.2fs 总计=%.2fs",
        spec.embed_id, total_rows, col_count, create_cost, fill_cost, total_cost,
    )


# ─── HTML 预览渲染器（无外部依赖，可单独测试） ─────────────────────────────────

def render_table_to_html(spec: EmbedSpec) -> str:
    """生成在富文本编辑器中可见的 HTML 预览表格。"""
    payload = spec.payload or {}
    style   = _merge_style(payload.get("style"))

    headers: List[str] = [str(h) for h in (payload.get("headers") or [])]
    rows: List[List]   = payload.get("rows") or []
    caption: str       = payload.get("caption") or ""
    has_header: bool   = bool(payload.get("has_header", True)) and bool(headers)

    header_bg    = f"#{style['header_bg']}"
    header_color = f"#{style['header_color']}"
    border_color = f"#{style['border_color']}"
    font_name    = style["font_name"]
    h_size       = style["header_font_size"]
    b_size       = style["body_font_size"]
    cell_align   = style["cell_alignment"]

    td_base = (
        f"padding:4px 8px; border:1px solid {border_color}; "
        f"font-family:{html.escape(font_name)},serif; text-align:{cell_align};"
    )

    lines: List[str] = []
    if caption:
        lines.append(
            f'<p class="excel-swk" style="font-weight:bold; margin:4px 0; '
            f'font-family:{html.escape(font_name)},serif;">'
            f'{html.escape(caption)}</p>'
        )

    lines.append(
        f'<table class="excel-swk" style="border-collapse:collapse; width:100%; '
        f'font-size:{b_size}pt; margin:4px 0;" data-embed-preview="table">'
    )

    if has_header:
        lines.append("<thead><tr>")
        for h in headers:
            lines.append(
                f'<th style="{td_base} background:{header_bg}; '
                f'color:{header_color}; font-size:{h_size}pt; '
                f'font-weight:bold;">{html.escape(str(h))}</th>'
            )
        lines.append("</tr></thead>")

    if rows:
        col_count = len(headers) if headers else (len(rows[0]) if rows else 0)
        lines.append("<tbody>")
        for row in rows:
            lines.append("<tr>")
            for ci in range(col_count):
                val = row[ci] if ci < len(row) else ""
                lines.append(f'<td style="{td_base}">{html.escape(str(val))}</td>')
            lines.append("</tr>")
        lines.append("</tbody>")

    lines.append("</table>")
    return "\n".join(lines)


# ─── 注册 ─────────────────────────────────────────────────────────────────────

register_html_renderer(TYPE_TABLE, render_table_to_html)
register_docx_renderer(TYPE_TABLE, render_table_to_docx)

logger.debug("embed_table_renderer: HTML 和 DOCX 渲染器已注册")
