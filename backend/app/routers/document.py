import datetime
import json
import logging
import os
from typing import Any, List, Optional

import psycopg2
from fastapi import APIRouter, Body, Form, HTTPException, Request, UploadFile, File
from fastapi.responses import JSONResponse
from psycopg2.extras import RealDictCursor

from app.converters.docx_converter import convert_html_to_docx, docx_to_html
from app.core.config import UPLOAD_DIR, STATIC_WEB_FRONT_PREFIX
from app.db.database import (
    build_simplified_tree,
    get_db_connection,
    get_next_batch_count,
    process_single_tree_node,
    process_split_tree_nodes,
    process_split_tree_nodes_with_select,
    query_and_build_tree,
)
from app.models.schemas import unified_response
from app.utils.file_utils import generate_unique_file_id
from app.utils.html_utils import (
    get_html_heading_levels,
    html_base64_images_to_urls,
    html_img_url_to_base64,
)
from app.utils.path_utils import save_html_and_get_url
from mergfile import MergeRequest, TreeItem, call_docx_merge, call_docx_split, call_set_table_width

router = APIRouter()
logger = logging.getLogger(__name__)

# ─── 格式配置（合并接口使用） ────────────────────────────────────────────────

_MERGE_FORMAT_CONFIG = {
    "Heading": {f"Heading{i}": {"use": i == 1, "style": {
        "alignment": "left", "line_spacing": "single", "line_spacing_value": 1,
        "left_indent": 0, "right_indent": 0, "space_before": 0, "space_after": 0,
        "first_line_indent": 0, "font_name": "仿宋",
        "font_size": {"1": "初号", "2": "三号", "3": "三号", "4": "四号",
                      "5": "四号", "6": "小四", "7": "小四", "8": "小四", "9": "五号"}[str(i)],
        "bold": True, "italic": i == 1, "underline": None, "color": None,
    }} for i in range(1, 10)},
    "Text": {"use": True, "style": {
        "alignment": "left", "line_spacing": "single", "line_spacing_value": 1,
        "left_indent": 0, "right_indent": 0, "space_before": 0, "space_after": 0,
        "first_line_indent": 0, "font_name": "仿宋", "font_size": "小四",
        "bold": None, "italic": None, "underline": True, "color": None,
    }},
    "Table": {"use": False, "style": {
        "repeat_header": False, "line_break": False, "alignment": "left",
        "font_name": "仿宋", "font_size": "五号", "left_indent": 0, "right_indent": 0,
        "first_line_indent": 0, "bold": None, "italic": None, "underline": None, "color": None,
    }},
    "Header": {"use": True, "show_logo": True, "logo": "", "show_name": False, "name": "123"},
    "Footer": {"use": False, "style": {"alignment": "left"}},
    "other": {"numbering": True, "use": False},
    "Margin": {"use": False, "top": 2.54, "bottom": 2.54, "left": 3.18, "right": 3.18},
}

_MERGE_FORMAT_ARGS = {
    "config_dict": _MERGE_FORMAT_CONFIG,
    "token": "984f5b0a2793eeafeeddfd2cd095ad31",
    "key": "984f5b0a2793eeafeeddfd2cd095ad31-1772598822992",
}


# ─── 路由 ────────────────────────────────────────────────────────────────────

@router.get("/get_html_by_node/{node_id}", summary="根据节点ID查询HTML文本")
async def get_html_by_node(request: Request, node_id: int):
    """根据标题树节点 ID 查询存储的 HTML 文本"""
    select_sql = """
    SELECT t.html_content, t.title_text, t.create_time, t.update_time, t.level, t.eid, t.idx,
           t.node_type, t.origin_file_path, t.is_conversion_completion,
           r.original_filename, r.upload_time, r.update_time as file_update_time,
           r.split_file_id, r.process_mode
    FROM "yxdl_docx_title_trees" t
    LEFT JOIN "yxdl_docx_upload_records" r ON t.record_id = r.id
    WHERE t.id = %s
    """
    try:
        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cursor:
                cursor.execute(select_sql, (node_id,))
                result = cursor.fetchone()
        if not result:
            return unified_response(404, f"未找到ID为{node_id}的标题树节点")

        if result["is_conversion_completion"] == 0:
            html_content, temp_file_docx_ = docx_to_html(result["origin_file_path"])
            eid = os.path.splitext(os.path.basename(temp_file_docx_))[0]
            with get_db_connection() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cursor:
                    cursor.execute(
                        """UPDATE "yxdl_docx_title_trees"
                           SET html_content = %s, update_time = NOW(),
                               is_conversion_completion = 1, update_file_path = %s, eid = %s
                           WHERE id = %s""",
                        (html_content, temp_file_docx_, eid, node_id),
                    )
                    conn.commit()
                    cursor.execute(select_sql, (node_id,))
                    updated_result = cursor.fetchone()
            return unified_response(200, "查询HTML文本成功", {
                "node_id": node_id,
                "title_text": updated_result["title_text"],
                "level": updated_result["level"],
                "http_path": save_html_and_get_url(updated_result["html_content"] or ""),
                "temp_file_docx_": temp_file_docx_,
            })

        return unified_response(200, "查询HTML文本成功", {
            "node_id": node_id,
            "title_text": result["title_text"],
            "level": result["level"],
            "http_path": save_html_and_get_url(result["html_content"] or ""),
        })

    except Exception as e:
        return unified_response(500, f"查询HTML文本失败：{str(e)}")


@router.post("/update_html_by_node", summary="更新节点HTML文本")
async def update_html_by_node(
    request: Request,
    node_id: int = Body(..., description="要更新的节点ID"),
    html_content: str = Body(..., description="更新后的HTML文本"),
    title_text: Optional[str] = Body(None, description="可选：更新节点标题文本"),
):
    """更新指定节点 ID 的 HTML 文本"""
    try:
        if node_id <= 0:
            return unified_response(400, "节点ID必须为正整数")
        if not html_content.strip():
            return unified_response(400, "HTML内容不能为空")

        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute('SELECT id FROM "yxdl_docx_title_trees" WHERE id = %s', (node_id,))
                if not cursor.fetchone():
                    return unified_response(404, f"未找到ID为{node_id}的标题树节点")

        html_content, _ = html_img_url_to_base64(html_content)
        html_content, _ = html_base64_images_to_urls(html_content, UPLOAD_DIR, STATIC_WEB_FRONT_PREFIX)

        success, result, temp_docx_path_1 = convert_html_to_docx(html_content)
        eid = os.path.splitext(os.path.basename(temp_docx_path_1))[0]

        update_fields = ["html_content = %s", "update_time = %s", "update_file_path = %s", "eid = %s"]
        current_time = datetime.datetime.now()
        update_values = [html_content, current_time, temp_docx_path_1, eid]

        if title_text is not None and title_text.strip():
            update_fields.append("title_text = %s")
            update_values.append(title_text.strip())

        update_sql = f"""
        UPDATE "yxdl_docx_title_trees"
        SET {", ".join(update_fields)}
        WHERE id = %s
        """
        update_values.append(node_id)

        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(update_sql, tuple(update_values))
                conn.commit()

        return unified_response(200, "节点HTML内容更新成功", {
            "node_id": node_id,
            "updated_title": "标题更新为" + title_text.strip() if title_text else "标题未更新",
            "update_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
        })

    except Exception as e:
        return unified_response(500, f"更新节点HTML失败：{str(e)}")


@router.post("/update_html_by_node_new", summary="更新节点HTML文本（新版，支持文件上传）")
async def update_html_by_node_new(
    request: Request,
    node_id: int = Form(..., description="要更新的节点ID"),
    file: UploadFile = File(..., description="HTML文件（.html）"),
    title_text: Optional[str] = Form(None, description="可选：更新节点标题文本"),
):
    """更新指定节点 ID 的 HTML 文本（通过文件上传，支持有标题时自动拆分重排）"""
    MAX_LEVEL_NODE = 9
    try:
        if node_id <= 0:
            return unified_response(400, "节点ID必须为正整数")

        html_content = (await file.read()).decode("utf-8")
        if not html_content.strip():
            return unified_response(400, "HTML内容不能为空")

        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    'SELECT id, record_id, level FROM "yxdl_docx_title_trees" WHERE id = %s',
                    (node_id,),
                )
                row = cursor.fetchone()
                if not row:
                    return unified_response(404, f"未找到ID为{node_id}的标题树节点")
                record_id = row[1]
                now_level = row[2]

        logger.info(f"update_html_by_node_new: node_id={node_id} record_id={record_id} level={now_level}")

        html_content, _ = html_img_url_to_base64(html_content)
        html_content, _ = html_base64_images_to_urls(html_content, UPLOAD_DIR, STATIC_WEB_FRONT_PREFIX)

        existing_levels, max_level = get_html_heading_levels(html_content)
        current_time = datetime.datetime.now()

        # ── 无标题：直接更新当前节点 ────────────────────────────────────────
        if max_level == 0:
            success, result, temp_docx_path_1 = convert_html_to_docx(html_content)
            eid = os.path.splitext(os.path.basename(temp_docx_path_1))[0]

            update_fields = ["html_content = %s", "update_time = %s",
                             "update_file_path = %s", "eid = %s", "is_conversion_completion = %s"]
            update_values = [html_content, current_time, temp_docx_path_1, eid, 1]

            if title_text is not None and title_text.strip():
                update_fields.append("title_text = %s")
                update_values.append(title_text.strip())

            update_sql = f"""
                UPDATE "yxdl_docx_title_trees"
                SET {", ".join(update_fields)}
                WHERE id = %s
            """
            update_values.append(node_id)

            with get_db_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(update_sql, tuple(update_values))
                    conn.commit()

            node_ids = query_and_build_tree(record_id, current_time)
            return unified_response(200, "节点HTML内容更新成功", {
                "node_id": node_id,
                "node_ids": node_ids,
                "updated_title": ("标题更新为" + title_text.strip()) if title_text else "标题未更新",
                "update_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
            })

        # ── 有标题：转换 → 拆分 → 重新入库 ──────────────────────────────────
        success, result, temp_docx_path_1 = convert_html_to_docx(html_content)
        if not success:
            return unified_response(500, f"HTML转DOCX失败：{result}")

        original_filename = os.path.abspath(temp_docx_path_1)
        split_file_id = generate_unique_file_id()

        insert_sql = """
            INSERT INTO "yxdl_docx_upload_records"
            (original_filename, new_filename, save_path,
             upload_time, update_time, split_file_id, process_mode)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            RETURNING id;
        """
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(insert_sql, (
                    original_filename, original_filename, temp_docx_path_1,
                    current_time, current_time, split_file_id, "split",
                ))
                new_record_id = cursor.fetchone()[0]
                conn.commit()

        new_file_path = call_set_table_width(temp_docx_path_1)
        with open(new_file_path, "rb") as _f:
            file_bytes = _f.read()

        split_result = call_docx_split(
            file_stream=file_bytes,
            file_name=original_filename,
            file_id=str(node_id),
            had_title=1,
            rm_outline_in_doc=1,
        )
        if split_result.status == 1:
            return unified_response(500, split_result.msg)

        from app.db.database import assign_file_path_to_tree, build_eid_path_mapping
        tree_nodes = [TreeItem(**item) for item in split_result.data.get("tree", [])]
        eid_path_map = build_eid_path_mapping(split_result.data.get("files", []))
        for node in tree_nodes:
            assign_file_path_to_tree(node, eid_path_map)

        if not tree_nodes:
            return unified_response(500, "拆分结果为空")

        batch_count = get_next_batch_count(record_id)
        first_node = tree_nodes.pop(0)
        first_result = process_single_tree_node(first_node, record_id, node_id, current_time, convert_html=False)

        remaining_nodes = (first_node.children or []) + tree_nodes
        process_split_tree_nodes(
            nodes=remaining_nodes,
            record_id=record_id,
            current_time=current_time,
            file_base_path=temp_docx_path_1,
            convert_html=False,
            parent_id=first_result.get("node_id"),
            batch_count=batch_count,
        )

        node_ids = query_and_build_tree(record_id, current_time)
        return unified_response(200, "更新成功", {
            "record_id": record_id,
            "node_ids": node_ids,
            "node_type": "branch",
            "split_file_id": split_file_id,
        })

    except Exception as e:
        return unified_response(500, f"更新节点HTML失败：{str(e)}")


@router.post("/merge_docx_office_server", summary="合并拆分的DOCX节点")
async def merge_docx_office_server(
    request: Request,
    node_id: int = Body(..., description="要更新的节点ID"),
    html_content: str = Body(..., description="需要转换的HTML文本"),
    filename: Optional[str] = Body("output.docx", description="下载的DOCX文件名"),
    title_text: Optional[str] = Body(None, description="可选：更新节点标题文本"),
):
    """调用合并接口生成合并后的 DOCX 文件"""
    if node_id <= 0:
        return unified_response(400, "节点ID必须为正整数")
    if not html_content.strip():
        return unified_response(400, "HTML内容不能为空")

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute('SELECT id, record_id FROM "yxdl_docx_title_trees" WHERE id = %s', (node_id,))
            row = cursor.fetchone()
            if not row:
                return unified_response(404, f"未找到ID为{node_id}的标题树节点")
            result_record_id = row[1]

    logger.info(f"merge_docx_office_server: node_id={node_id} record_id={result_record_id}")
    current_time = datetime.datetime.now()

    select_sql = """
        SELECT
            id, title_text, level, eid, idx, parent_id, batch_count,
            origin_file_path, update_file_path, is_conversion_completion, split_id
        FROM "yxdl_docx_title_trees"
        WHERE record_id = %s
        ORDER BY level ASC, idx ASC;
    """
    try:
        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cursor:
                cursor.execute(select_sql, (result_record_id,))
                node_records = cursor.fetchall()
    except Exception as e:
        return unified_response(500, f"查询数据库失败：{str(e)}")

    if not node_records:
        return unified_response(404, "该记录下无任何节点")

    tree_nodes_org = [
        TreeItem(**{
            "id":                       item.get("split_id"),
            "text":                     item.get("title_text"),
            "level":                    item.get("level"),
            "eid":                      item.get("eid"),
            "idx":                      item.get("idx"),
            "parent_id":                item.get("parent_id"),
            "file_path":                item.get("origin_file_path"),
            "update_file_path":         item.get("update_file_path", ""),
            "is_conversion_completion": item.get("is_conversion_completion", 0),
            "children": [], "file_name": None, "file_info": None, "node_type": "",
        })
        for item in node_records
    ]

    nested_dicts = build_simplified_tree(node_records)

    def _refresh_level_idx(nodes: list, parent_level: int = 0, counter: List[int] = None) -> list:
        if counter is None:
            counter = [0]
        for node in nodes:
            node["level"] = parent_level + 1
            node["idx"] = counter[0]
            counter[0] += 1
            if node.get("children"):
                _refresh_level_idx(node["children"], parent_level + 1, counter)
        return nodes

    nested_dicts = _refresh_level_idx(nested_dicts)

    def _collect_updates(nodes: list) -> list:
        result = []
        for node in nodes:
            result.append({"id": node["id"], "level": node["level"], "idx": node["idx"]})
            if node.get("children"):
                result.extend(_collect_updates(node["children"]))
        return result

    updates = _collect_updates(nested_dicts)
    if updates:
        try:
            with get_db_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.executemany(
                        """UPDATE "yxdl_docx_title_trees"
                           SET level = %s, idx = %s, update_time = %s
                           WHERE id = %s""",
                        [(u["level"], u["idx"], current_time, u["id"]) for u in updates],
                    )
                    conn.commit()
        except Exception as e:
            logger.error(f"merge_docx_office_server: 刷新 level/idx 失败 err={e}")
            return unified_response(500, f"刷新节点层级失败：{str(e)}")

    def _dicts_to_tree_items(nodes_dict: list) -> list:
        result = []
        for d in nodes_dict:
            item = TreeItem(
                eid=d.get("eid", ""),
                level=d.get("level", 1),
                idx=d.get("idx", 0),
                text=d.get("text", "") or d.get("title_text", ""),
                children=_dicts_to_tree_items(d.get("children", [])),
            )
            matched = next((n for n in tree_nodes_org if n.eid == item.eid), None)
            if matched:
                item.id                       = matched.id
                item.parent_id                = matched.parent_id
                item.file_path                = matched.file_path
                item.update_file_path         = matched.update_file_path
                item.is_conversion_completion = matched.is_conversion_completion
            result.append(item)
        return result

    nested_tree_items = _dicts_to_tree_items(nested_dicts)

    def _collect_files(nodes: list) -> list:
        paths = []
        seen = set()
        def _dfs(node_list):
            for node in node_list:
                fp = node.update_file_path if node.is_conversion_completion == 1 and node.update_file_path else (node.file_path or "")
                if fp and fp not in seen:
                    seen.add(fp)
                    paths.append(fp)
                _dfs(node.children or [])
        _dfs(nodes)
        return paths

    files_ = _collect_files(nested_tree_items)

    try:
        merge_request = MergeRequest(tree=nested_tree_items, files=files_, format_args=_MERGE_FORMAT_ARGS)
        merged_file_message = call_docx_merge(merge_request, add_title=0, add_heading_num=1)
        old_filepath = merged_file_message.data.get("filepath", "")
        if old_filepath:
            new_filepath = call_set_table_width(old_filepath)
            merged_file_message.data["filepath"] = new_filepath
        return merged_file_message
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"文件合并失败：{str(e)}")


@router.post("/generator_query_by_type", summary="查询生成器格式")
async def query_format_storage_by_type(
    request: Request,
    formant_type: Any = Body(..., description="type类型，format_storage_id"),
    table_title: str = Body("", description="站位数据"),
):
    """通过 type 查询配置格式存储数据"""
    try:
        request_body = await request.json()
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="请求体格式错误，必须是有效的 JSON")

    if "formant_type" not in request_body:
        raise HTTPException(status_code=400, detail="缺少必填参数：formant_type")

    type_value = formant_type
    if not isinstance(type_value, int):
        try:
            type_value = int(type_value)
        except (ValueError, TypeError):
            raise HTTPException(status_code=400, detail="参数 type 必须是整数")

    query_sql = """
        SELECT id, format_name, base64_img
        FROM cfg_format_storage
        WHERE type = %s AND status = 1;
    """
    try:
        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute(query_sql, (type_value,))
                results = [dict(row) for row in cur.fetchall()]
        return unified_response(200, "查询成功", results)
    except psycopg2.Error as e:
        raise HTTPException(status_code=500, detail=f"数据库查询失败：{str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"查询失败：{str(e)}")
