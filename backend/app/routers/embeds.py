"""
嵌入组件（Embed）路由

- POST /doc_editor/embeds                 创建组件并返回可插入 HTML 的标记
- GET  /doc_editor/embeds/{embed_id}      获取组件详情（payload）
- PUT  /doc_editor/embeds/{embed_id}      更新组件
- DELETE /doc_editor/embeds/{embed_id}    软删除组件
- GET  /doc_editor/embeds/{embed_id}/go   URL 跳转（302 → url）
- POST /doc_editor/embeds/scan            从 HTML 扫描所有 embed 标记
- GET  /doc_editor/embeds/by_record/{record_id}  列出某个文档下所有组件
"""
import logging
from typing import Any, Dict, List, Optional

import io
import os
import tempfile

from fastapi import APIRouter, Body, Path
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse

from app.db.database import (
    delete_embed_component,
    get_embed_component,
    insert_embed_component,
    list_embed_components_by_record,
    update_embed_component,
)
from app.models.schemas import unified_response
from app.utils.embed_marker import (
    DISPLAY_BLOCK,
    DISPLAY_INLINE,
    EmbedSpec,
    build_embed_marker,
    parse_embed_markers_from_html,
    spec_from_db_row,
    spec_to_db_row,
)

# 导入即完成渲染器注册
import app.utils.embed_table_renderer  # noqa: F401

router = APIRouter()
logger = logging.getLogger(__name__)


# ─── 创建 ─────────────────────────────────────────────────────────────────────

@router.post("/embeds", summary="创建嵌入组件并返回可插入HTML的标记")
async def create_embed(
    data: Dict[str, Any] = Body(..., description="组件业务数据(JSON)"),
    embed_type: str = Body("custom", description="组件类型：text/table/image/chart/reference/custom"),
    title: str = Body("", description="标题（显示用）"),
    url: Optional[str] = Body(None, description="跳转URL；不传则自动生成 /doc_editor/embeds/{id}/go"),
    display: Optional[str] = Body(None, description="inline / block；不传按类型自动判定"),
    version: int = Body(1, description="payload结构版本"),
    record_id: Optional[int] = Body(None, description="关联文档记录ID"),
    node_id: Optional[int] = Body(None, description="关联节点ID"),
):
    try:
        spec, snippet = build_embed_marker(
            data=data,
            embed_type=embed_type,
            title=title,
            url=url,
            display=display,
            version=version,
            record_id=record_id,
            node_id=node_id,
            url_builder=lambda eid: f"/doc_editor/embeds/{eid}/go",
        )
    except ValueError as e:
        return unified_response(400, str(e))

    try:
        insert_embed_component(spec_to_db_row(spec))
    except Exception as e:
        logger.error(f"create_embed: 入库失败 err={e}")
        return unified_response(500, f"入库失败：{str(e)}")

    return unified_response(200, "创建成功", {
        "embed_id": spec.embed_id,
        "url":      spec.url,
        "display":  spec.display,
        "html":     snippet,
        "spec":     spec.to_dict(),
    })


# ─── 查询 ─────────────────────────────────────────────────────────────────────

@router.get("/embeds/{embed_id}", summary="查询嵌入组件详情")
async def get_embed(embed_id: str = Path(..., description="embed_id，如 EMB_xxxxxxxxxxxx")):
    row = get_embed_component(embed_id)
    if not row:
        return unified_response(404, f"未找到 embed_id={embed_id}")
    spec = spec_from_db_row(row)
    return unified_response(200, "查询成功", spec.to_dict())


# ─── 更新 ─────────────────────────────────────────────────────────────────────

@router.put("/embeds/{embed_id}", summary="更新嵌入组件")
async def update_embed(
    embed_id: str = Path(...),
    data: Optional[Dict[str, Any]] = Body(None, description="新的业务数据；为空时不更新payload"),
    title: Optional[str] = Body(None),
    url: Optional[str] = Body(None),
    display: Optional[str] = Body(None),
    version: Optional[int] = Body(None),
):
    row = get_embed_component(embed_id)
    if not row:
        return unified_response(404, f"未找到 embed_id={embed_id}")

    spec = spec_from_db_row(row)
    if data is not None:
        spec.payload = data
    if title is not None:
        spec.title = title
    if url is not None:
        spec.url = url
    if display in (DISPLAY_INLINE, DISPLAY_BLOCK):
        spec.display = display
    if version is not None:
        spec.version = int(version)

    ok = update_embed_component(embed_id, spec_to_db_row(spec))
    if not ok:
        return unified_response(500, "更新失败")
    return unified_response(200, "更新成功", spec.to_dict())


# ─── 删除 ─────────────────────────────────────────────────────────────────────

@router.delete("/embeds/{embed_id}", summary="删除嵌入组件（软删除）")
async def delete_embed(embed_id: str = Path(...)):
    ok = delete_embed_component(embed_id)
    if not ok:
        return unified_response(404, f"未找到或已删除：{embed_id}")
    return unified_response(200, "删除成功", {"embed_id": embed_id})


# ─── URL 跳转 ─────────────────────────────────────────────────────────────────

@router.get("/embeds/{embed_id}/go", summary="根据embed_id跳转到组件URL")
async def go_embed(embed_id: str = Path(...)):
    row = get_embed_component(embed_id)
    if not row:
        return unified_response(404, f"未找到 embed_id={embed_id}")

    target = row.get("url")

    # 修复：如果 target 为空，或者 target 指向了接口自身，则不进行跳转，直接返回数据
    if not target or target.endswith(f"/embeds/{embed_id}/go"):
        spec = spec_from_db_row(row)
        return unified_response(200, "该组件无跳转URL", spec.to_dict())

    return RedirectResponse(url=target, status_code=302)


# ─── 扫描 HTML ────────────────────────────────────────────────────────────────

@router.post("/embeds/scan", summary="从HTML文本中扫描所有嵌入组件标记")
async def scan_embeds(html_text: str = Body(..., embed=True, description="HTML文本")):
    markers = parse_embed_markers_from_html(html_text)
    return unified_response(200, "扫描成功", {
        "count":   len(markers),
        "markers": markers,
    })


# ─── 列表 ─────────────────────────────────────────────────────────────────────

@router.get("/embeds/by_record/{record_id}", summary="列出某文档下所有嵌入组件")
async def list_by_record(record_id: int = Path(..., gt=0)):
    rows = list_embed_components_by_record(record_id)
    return unified_response(200, "查询成功", {
        "record_id": record_id,
        "count":     len(rows),
        "items":     rows,
    })


# ─── 测试接口 ─────────────────────────────────────────────────────────────────

_TEST_PAYLOAD = {
    "caption": "表1 项目人员信息",
    "headers": ["姓名", "部门", "职位", "入职年份"],
    "rows": [
        ["张三", "研发部", "高级工程师", "2019"],
        ["李四", "市场部", "产品经理",  "2021"],
        ["王五", "财务部", "会计",      "2020"],
        ["赵六", "法务部", "法务专员",  "2022"],
    ],
    "has_header": True,
    "col_widths": [3.0, 3.0, 4.0, 3.0],
    "merge_cells": [],
    "style": {
        "header_bg":        "4472C4",
        "header_color":     "FFFFFF",
        "border_color":     "BFBFBF",
        "font_name":        "仿宋",
        "header_font_size": 10,
        "body_font_size":   9,
        "cell_alignment":   "center",
    },
}


@router.get(
    "/embeds/test/html",
    summary="【测试】HTML表格预览渲染，直接返回表格页面",
    response_class=HTMLResponse,
)
async def test_html_preview(
    caption: str = "表1 项目人员信息",
):
    from app.utils.embed_marker import build_embed_marker, TYPE_TABLE
    from app.utils.embed_table_renderer import render_table_to_html

    payload = {**_TEST_PAYLOAD, "caption": caption}
    spec, _ = build_embed_marker(data=payload, embed_type=TYPE_TABLE, title=caption)
    table_html = render_table_to_html(spec)

    page = f"""<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="UTF-8">
  <title>表格预览</title>
  <style>
    body {{ font-family: "仿宋", serif; padding: 40px; background: #fff; }}
  </style>
</head>
<body>
  {table_html}
</body>
</html>"""
    return HTMLResponse(content=page)


@router.post(
    "/embeds/test/docx",
    summary="【测试】DOCX表格替换渲染，返回可下载的 .docx 文件",
)
async def test_docx_replace(
        payload: Optional[Dict[str, Any]] = Body(
            None,
            description="表格 payload；不传则使用内置示例数据",
        ),
):
    try:
        from docx import Document
    except ImportError:
        return unified_response(500, "python-docx 未安装")

    from app.utils.embed_marker import TYPE_TABLE, VISIBLE_PLACEHOLDER_FMT, build_embed_marker
    from app.utils.embed_table_renderer import render_table_to_docx

    # 👇 关键修复点：如果传进来的是空字典 {}，也会使用默认测试数据
    # 只有当传来的 payload 里面真正包含 "rows" 或 "headers" 时，才使用传入的数据，否则一律使用测试数据
    tbl_payload = payload if (payload and ("rows" in payload or "headers" in payload)) else _TEST_PAYLOAD

    spec, _ = build_embed_marker(
        data=tbl_payload,
        embed_type=TYPE_TABLE,
        title=tbl_payload.get("caption", "测试表格"),
    )

    doc = Document()
    placeholder_text = VISIBLE_PLACEHOLDER_FMT.format(embed_id=spec.embed_id)
    placeholder_para = doc.add_paragraph(placeholder_text)

    try:
        render_table_to_docx(spec, placeholder_para, doc)
    except Exception as e:
        logger.exception("test_docx_replace: 渲染失败")
        return unified_response(500, f"表格渲染失败：{e}")

    import tempfile
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", prefix="embed_test_", delete=False)
    tmp.close()
    doc.save(tmp.name)

    from fastapi.responses import FileResponse
    return FileResponse(
        path=tmp.name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"embed_table_test_{spec.embed_id}.docx",
    )