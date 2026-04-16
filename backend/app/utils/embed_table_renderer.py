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
from typing import Any, Dict, List, Optional

from app.utils.embed_marker import (
    EmbedSpec,
    TYPE_TABLE,
    register_docx_renderer,
    register_html_renderer,
)

logger = logging.getLogger(__name__)

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
    """写入单元格文字 + 样式。"""
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


# ─── DOCX 渲染器 ───────────────────────────────────────────────────────────────

def render_table_to_docx(spec: EmbedSpec, paragraph, document) -> None:
    """
    用 spec.payload 中的表格数据，在 paragraph 所在位置插入真实 Word 表格。
    插入完成后移除占位符段落。
    """
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.shared import Cm

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
        logger.warning("render_table_to_docx: 列数为 0，跳过")
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
        logger.warning("render_table_to_docx: 无任何行数据，跳过")
        return

    align_para   = _align_enum(style["cell_alignment"])
    font_name    = style["font_name"]
    p_elem       = paragraph._element
    parent       = p_elem.getparent()

    # 记录占位符在父节点中的当前索引（后续所有插入基于此值）
    insert_idx = list(parent).index(p_elem)

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
    # add_table 追加到 body 末尾；先把 tbl_elem 从末尾取出，再插到正确位置
    table = document.add_table(rows=len(all_rows), cols=col_count)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_elem = table._tbl
    tbl_elem.getparent().remove(tbl_elem)   # 从末尾摘除
    parent.insert(insert_idx, tbl_elem)      # 插到占位符所在位置

    # ── 列宽 ─────────────────────────────────────────────────────────────────
    if col_widths:
        for row_obj in table.rows:
            for ci, cell in enumerate(row_obj.cells):
                if ci < len(col_widths):
                    cell.width = Cm(float(col_widths[ci]))
        table.autofit = False
    else:
        table.autofit = True

    # ── 填充单元格 ────────────────────────────────────────────────────────────
    for ri, (row_data, is_hdr) in enumerate(zip(all_rows, is_header_row)):
        for ci, text in enumerate(row_data):
            cell = table.cell(ri, ci)
            _write_cell(
                cell         = cell,
                text         = text,
                font_name    = font_name,
                font_size    = style["header_font_size"] if is_hdr else style["body_font_size"],
                color        = style["header_color"] if is_hdr else "000000",
                bold         = is_hdr,
                align        = align_para,
                bg           = style["header_bg"] if is_hdr else None,
                border_color = style["border_color"],
            )

    # ── 合并单元格（[行, 起始列, 结束列]） ────────────────────────────────────
    for instr in merge_instructions:
        if len(instr) != 3:
            continue
        ri, c_start, c_end = int(instr[0]), int(instr[1]), int(instr[2])
        if ri >= len(all_rows) or c_end >= col_count:
            continue
        try:
            table.cell(ri, c_start).merge(table.cell(ri, c_end))
        except Exception as e:
            logger.warning(f"合并单元格失败 [{ri},{c_start},{c_end}]：{e}")

    # ── 移除占位符段落（此时它已被推到 insert_idx+1 位置） ────────────────────
    p_elem.getparent().remove(p_elem)

    logger.info(
        f"render_table_to_docx: embed_id={spec.embed_id} "
        f"rows={len(all_rows)} cols={col_count}"
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
            f'<p style="font-weight:bold; margin:4px 0; '
            f'font-family:{html.escape(font_name)},serif;">'
            f'{html.escape(caption)}</p>'
        )

    lines.append(
        f'<table style="border-collapse:collapse; width:100%; '
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


