import datetime
import json
import logging
import os
import time
from typing import Any, List, Optional, Dict

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
from app.models.schemas import unified_response, UpdateTreeStructureRequest, TreeNodeUpdate
from app.utils.file_utils import generate_unique_file_id
from app.utils.html_utils import (
    get_html_heading_levels,
    html_base64_images_to_urls,
    html_img_url_to_base64,
    replace_first_heading_text,
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
    # 一键排版参数
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


@router.post("/modify_title_and_inner_cylinder_title_by_node", summary="修改节点标题及HTML内标题")
async def modify_title_and_inner_cylinder_title_by_node(
        request: Request,
        node_id: int = Body(..., description="要修改的节点ID"),
        new_title: str = Body(..., description="新标题文本"),
):
    """修改文件树节点的标题（title_text），同步将 HTML 内第一个标题标签的文本替换为新标题，
    重新转换为 DOCX 并更新 yxdl_docx_title_trees 表。"""
    try:
        if node_id <= 0:
            return unified_response(400, "节点ID必须为正整数")
        new_title = new_title.strip()
        if not new_title:
            return unified_response(400, "新标题不能为空")

        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cursor:
                cursor.execute(
                    """SELECT id, html_content, title_text,
                              is_conversion_completion, origin_file_path
                       FROM "yxdl_docx_title_trees" WHERE id = %s""",
                    (node_id,),
                )
                row = cursor.fetchone()
        if not row:
            return unified_response(404, f"未找到ID为{node_id}的标题树节点")

        html_content = row["html_content"] or ""

        # 若 html_content 尚未生成（is_conversion_completion=0），先从 origin_file_path 转换
        if not html_content.strip() or row["is_conversion_completion"] == 0:
            origin_path = row.get("origin_file_path") or ""
            if origin_path.strip():
                html_content, _ = docx_to_html(origin_path)
            # 转换后仍为空则降级为仅改 title_text
            if not html_content.strip():
                current_time = datetime.datetime.now()
                with get_db_connection() as conn:
                    with conn.cursor() as cursor:
                        cursor.execute(
                            """UPDATE "yxdl_docx_title_trees"
                               SET title_text = %s, update_time = %s
                               WHERE id = %s""",
                            (new_title, current_time, node_id),
                        )
                        conn.commit()
                return unified_response(200, "节点标题修改成功（仅标题，无HTML内容）", {
                    "node_id": node_id,
                    "old_title": row["title_text"],
                    "new_title": new_title,
                    "update_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                })

        html_content = replace_first_heading_text(html_content, new_title)
        success, result, temp_docx_path = convert_html_to_docx(html_content)
        if not success:
            return unified_response(500, f"HTML转DOCX失败：{result}")
        eid = os.path.splitext(os.path.basename(temp_docx_path))[0]
        current_time = datetime.datetime.now()
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    """UPDATE "yxdl_docx_title_trees"
                       SET title_text = %s, html_content = %s,
                           update_file_path = %s, eid = %s,
                           update_time = %s, is_conversion_completion = 1
                       WHERE id = %s""",
                    (new_title, html_content, temp_docx_path, eid, current_time, node_id),
                )
                conn.commit()

        return unified_response(200, "节点标题修改成功", {
            "node_id": node_id,
            "old_title": row["title_text"],
            "new_title": new_title,
            "update_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
        })

    except Exception as e:
        return unified_response(500, f"修改节点标题失败：{str(e)}")


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

        # 调用表格宽度适配接口
        # new_file_path = call_set_table_width(temp_docx_path_1)
        new_file_path = temp_docx_path_1
        with open(new_file_path, "rb") as _f:
            file_bytes = _f.read()

        split_result = call_docx_split(
            file_stream=file_bytes,
            file_name=original_filename,
            file_id=str(node_id),
            had_title=1,
            rm_outline_in_doc=1,
            del_page_break=0,
        )
        if split_result.status == 1:
            return unified_response(500, split_result.msg)

        title_font_dict = split_result.data.get("title_font_dict") or {}
        if title_font_dict:
            import json
            with get_db_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(
                        'UPDATE "yxdl_docx_upload_records" SET title_font_dict = %s WHERE id = %s',
                        (json.dumps(title_font_dict), new_record_id),
                    )
                    conn.commit()

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
        config_dict: Dict[str, Any] = Body(
            None,
            description="要更新的节点配置字典",
            example={"id": 123, "name": "test_node"}  # 建议加上示例
        ),
        token: str = Body(None, description="要更新的节点ID"),
        key: str = Body(None, description="要更新的节点ID"),
):
    """调用合并接口生成合并后的 DOCX 文件"""

    if node_id <= 0:
        return unified_response(400, "节点ID必须为正整数")
    if not html_content.strip():
        return unified_response(400, "HTML内容不能为空")

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(
                '''SELECT t.id, t.record_id, r.title_font_dict
                   FROM "yxdl_docx_title_trees" t
                   LEFT JOIN "yxdl_docx_upload_records" r ON r.id = t.record_id
                   WHERE t.id = %s''',
                (node_id,),
            )
            row = cursor.fetchone()
            if not row:
                return unified_response(404, f"未找到ID为{node_id}的标题树节点")
            result_record_id = row[1]
            title_font_dict = row[2] or {}

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
            "id": item.get("split_id"),
            "text": item.get("title_text"),
            "level": item.get("level"),
            "eid": item.get("eid"),
            "idx": item.get("idx"),
            "parent_id": item.get("parent_id"),
            "file_path": item.get("origin_file_path"),
            "update_file_path": item.get("update_file_path", ""),
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
                item.id = matched.id
                item.parent_id = matched.parent_id
                item.file_path = matched.file_path
                item.update_file_path = matched.update_file_path
                item.is_conversion_completion = matched.is_conversion_completion
            result.append(item)
        return result

    nested_tree_items = _dicts_to_tree_items(nested_dicts)

    def _collect_files(nodes: list) -> list:
        paths = []
        seen = set()

        def _dfs(node_list):
            for node in node_list:
                fp = node.update_file_path if node.is_conversion_completion == 1 and node.update_file_path else (
                            node.file_path or "")
                if fp and fp not in seen:
                    seen.add(fp)
                    paths.append(fp)
                _dfs(node.children or [])

        _dfs(nodes)
        return paths

    files_ = _collect_files(nested_tree_items)

    # ── 合并 DOCX ─────────────────────────────────────────────────────────────
    try:
        if config_dict:
            megre_docx_config = {"config_dict": config_dict,
                                 # 一键排版参数
                                 "token": token,
                                 "key": key,
                                 "title_font_dict": title_font_dict, }
            merge_request = MergeRequest(tree=nested_tree_items, files=files_, format_args=megre_docx_config)
        else:
            megre_docx_config = {"title_font_dict": title_font_dict} if title_font_dict else {}
            if megre_docx_config:
                merge_request = MergeRequest(tree=nested_tree_items, files=files_, format_args=megre_docx_config)
            else:
                merge_request = MergeRequest(tree=nested_tree_items, files=files_)
        logger.info("一键排版参数")
        logger.info(megre_docx_config)
        merged_file_message = call_docx_merge(merge_request, add_title=0, add_heading_num=1, update_title=1)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"文件合并失败：{str(e)}")

    # ── 表格宽度适配（独立 try，失败时降级使用原路径，不影响后续 embed 替换） ────
    old_filepath = merged_file_message.data.get("out_path", "")
    new_filepath = ""
    if old_filepath:
        try:
            widened = call_set_table_width(old_filepath)
            if widened:
                merged_file_message.data["filepath"] = widened
                new_filepath = widened
            else:
                logger.warning("merge_docx_office_server: call_set_table_width 返回空，使用原路径")
        except Exception as e:
            logger.warning(f"merge_docx_office_server: call_set_table_width 失败，降级使用原路径 err={e}")

    # embed 替换使用的有效路径：优先 new_filepath，兜底 old_filepath
    effective_filepath = new_filepath or old_filepath

    # ── 替换合并后 DOCX 中的 embed 占位符为真实表格 ──────────────────────────
    # 处理逻辑：
    #   1. 扫描 DOCX，找出所有包含 【EMB_xxx】 的占位符段落
    #   2. 记录每个占位符段落紧跟的下一个元素：若为表格（w:tbl），说明是
    #      HTML embed 标记被 DOCX 转换服务渲染时遗留的预览表格（前 N 行 + 全部数据链接），
    #      需要在 DOCX 渲染器插入完整表格后将其移除，避免文档中出现重复表格。
    #   3. 调用 render_docx_replace_plan 将占位符替换为完整 Word 表格。
    #   4. 删除步骤 2 中标记的遗留表格。
    #
    # 性能/可观测性：每个阶段都有独立的耗时日志，方便定位慢点。
    if effective_filepath:
        try:
            from docx import Document
            from app.utils.embed_marker import (
                build_docx_replace_plan_from_map,
                collect_docx_embed_paragraphs,
                render_docx_replace_plan_parallel,
                spec_from_db_row,
            )
            from app.db.database import get_embed_components_by_ids

            t_pipeline = time.perf_counter()

            # ── 第一步：打开 DOCX，单次扫描同时收集 embed_id 和对应段落 ──────
            # collect_docx_embed_paragraphs 一次遍历返回 {embed_id: paragraph}，
            # 后续用映射直接构建替换计划，省去 build_docx_replace_plan 的第二次段落扫描。
            t_stage = time.perf_counter()
            doc = Document(effective_filepath)
            para_map = collect_docx_embed_paragraphs(doc)
            logger.info(
                f"merge_docx_office_server[embed]: 扫描 DOCX 占位符完成"
                f" record_id={result_record_id} found={len(para_map)}"
                f" 耗时={time.perf_counter() - t_stage:.2f}s"
            )

            if not para_map:
                logger.info(
                    f"merge_docx_office_server[embed]: DOCX 无占位符，跳过 DB 查询"
                    f" record_id={result_record_id}"
                )
            else:
                # ── 第二步：按实际 embed_id 精准查询 DB，不加载无关组件 ──────────
                t_stage = time.perf_counter()
                rows_by_id = get_embed_components_by_ids(list(para_map.keys()))
                logger.info(
                    f"merge_docx_office_server[embed]: 加载 embed 组件"
                    f" record_id={result_record_id}"
                    f" matched={len(rows_by_id)}"
                    f" 耗时={time.perf_counter() - t_stage:.2f}s"
                )

                if rows_by_id:
                    t_stage = time.perf_counter()
                    specs_by_id = {
                        eid: spec_from_db_row(row)
                        for eid, row in rows_by_id.items()
                    }
                    logger.info(
                        f"merge_docx_office_server[embed]: 反序列化 specs 完成"
                        f" specs={len(specs_by_id)}"
                        f" 耗时={time.perf_counter() - t_stage:.2f}s"
                    )

                    # ── 第三步：从映射直接组装替换计划，无需重新扫描文档 ──────
                    t_stage = time.perf_counter()
                    plan = build_docx_replace_plan_from_map(para_map, specs_by_id)
                    logger.info(
                        f"merge_docx_office_server[embed]: 构建替换计划完成 plan={len(plan)}"
                        f" 耗时={time.perf_counter() - t_stage:.2f}s"
                    )

                    if plan:
                        # ── 收集遗留元素（预览表格及 tip 段落） ──────────────────
                        # _render_preview_page 将【EMB_xxx】注入 caption <p> 的首部，
                        # DOCX 转换后结构为：
                        #   段落:【EMB_xxx】caption  ← 占位符（caption <p> 转来的，含隐藏 span）
                        #   表格: 预览表格           ← 遗留预览表格（直接紧跟，无中间段落）
                        #   段落: tip（可选）        ← "仅显示前N行…" 提示段
                        # 扫描策略：直接找 w:tbl，再收集 tbl 后紧跟的 w:p（tip）。
                        t_stage = time.perf_counter()
                        stale_elems: list = []
                        for item in plan:
                            para_elem = item["paragraph"]._element
                            intermediate: list = []
                            sibling = para_elem.getnext()
                            found_tbl = None
                            for _ in range(5):  # 最多向后扫 5 个兄弟节点
                                if sibling is None:
                                    break
                                tag = sibling.tag
                                if tag.endswith("}tbl") or tag == "w:tbl":
                                    found_tbl = sibling
                                    break
                                elif tag.endswith("}p") or tag == "w:p":
                                    intermediate.append(sibling)
                                    sibling = sibling.getnext()
                                else:
                                    break
                            if found_tbl is not None:
                                stale_elems.extend(intermediate)
                                stale_elems.append(found_tbl)
                                # 收集 tbl 后紧跟的 tip 段落（如有）
                                after_tbl = found_tbl.getnext()
                                if after_tbl is not None and (
                                    after_tbl.tag.endswith("}p") or after_tbl.tag == "w:p"
                                ):
                                    stale_elems.append(after_tbl)
                        logger.info(
                            f"merge_docx_office_server[embed]: 定位遗留元素完成"
                            f" stale={len(stale_elems)}"
                            f" 耗时={time.perf_counter() - t_stage:.2f}s"
                        )

                        # ── 执行 embed 占位符 → 完整 Word 表格的替换 ─────────────
                        # 使用进程池并行构建各表格的 <w:tbl> XML，主进程串行 splice；
                        # plan 项 ≤ 1 时会自动降级为串行，避免进程启动开销倒赔。
                        t_stage = time.perf_counter()
                        replaced = render_docx_replace_plan_parallel(plan, doc, max_workers=4)
                        logger.info(
                            f"merge_docx_office_server[embed]: 渲染替换完成"
                            f" replaced={replaced}/{len(plan)}"
                            f" 耗时={time.perf_counter() - t_stage:.2f}s"
                        )

                        # ── 删除遗留元素（中间 caption 段落 + 预览表格） ─────────
                        t_stage = time.perf_counter()
                        removed = 0
                        for elem in stale_elems:
                            parent = elem.getparent()
                            if parent is not None:
                                try:
                                    parent.remove(elem)
                                    removed += 1
                                except Exception as rm_err:
                                    logger.warning(
                                        f"merge_docx_office_server: 删除遗留元素失败 err={rm_err}"
                                    )
                        logger.info(
                            f"merge_docx_office_server[embed]: 删除遗留元素完成"
                            f" removed={removed}/{len(stale_elems)}"
                            f" 耗时={time.perf_counter() - t_stage:.2f}s"
                        )

                        t_stage = time.perf_counter()
                        doc.save(effective_filepath)
                        logger.info(
                            f"merge_docx_office_server[embed]: 保存 DOCX 完成"
                            f" 耗时={time.perf_counter() - t_stage:.2f}s"
                        )

                        logger.info(
                            f"merge_docx_office_server[embed]: 全流程完成"
                            f" record_id={result_record_id} filepath={effective_filepath}"
                            f" replaced={replaced} removed_stale={removed}"
                            f" 总耗时={time.perf_counter() - t_pipeline:.2f}s"
                        )
                    else:
                        logger.info(
                            f"merge_docx_office_server[embed]: 无可替换占位符 record_id={result_record_id}"
                        )
                else:
                    logger.info(
                        f"merge_docx_office_server[embed]: DB 中无匹配 embed 组件"
                        f" record_id={result_record_id}"
                    )
        except Exception as e:
            logger.error(f"merge_docx_office_server: embed 替换失败 err={e}")

    # ── 修正 out_map_path：使其指向最终文件（filepath），而非原始合并文件（out_path） ──
    # 背景：call_set_table_width 会生成带 settable- 前缀的新文件并写入 data["filepath"]，
    # 但外部合并服务只返回了 out_path 对应的 out_map_path URL，
    # 此处将 URL 里的文件名替换为 filepath 的文件名，使前端拿到的下载链接指向正确文件。
    if effective_filepath:
        old_out_path = merged_file_message.data.get("out_path", "")
        old_map_url  = merged_file_message.data.get("out_map_path", "")
        if old_out_path and old_map_url:
            old_basename = os.path.basename(old_out_path)
            new_basename = os.path.basename(effective_filepath)
            if old_basename and old_basename != new_basename and old_basename in old_map_url:
                new_map_url = old_map_url.replace(old_basename, new_basename, 1)
                merged_file_message.data["out_map_path"] = new_map_url
                logger.info(
                    f"merge_docx_office_server: out_map_path 已更新"
                    f" {old_map_url} -> {new_map_url}"
                )

    return merged_file_message


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


@router.post("/update_tree_structure_based_on_record_id", summary="根据record_id更新树结构层级")
async def update_tree_structure_based_on_record_id(body: UpdateTreeStructureRequest):
    """
    根据前端传入的树形结构更新 yxdl_docx_title_trees 中节点的
    level、parent_id、idx 字段。

    - idx 按 DFS 顺序全局递增（与现有逻辑保持一致）
    - 根节点的 parent_id 置为 NULL
    """
    record_id = body.record_id

    # ── 1. 校验 record_id 存在 ─────────────────────────────────────────────
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    'SELECT id FROM "yxdl_docx_upload_records" WHERE id = %s',
                    (record_id,),
                )
                if not cursor.fetchone():
                    return unified_response(404, f"未找到record_id={record_id}的上传记录")
    except Exception as e:
        return unified_response(500, f"数据库查询失败：{str(e)}")

    # ── 2. DFS 展开树，收集 (node_id, level, parent_id, idx) ──────────────
    updates: list[tuple] = []  # (level, parent_id, idx, node_id)
    counter = [0]

    def _dfs(nodes: list[TreeNodeUpdate], parent_id: int | None):
        for node in nodes:
            idx = counter[0]
            counter[0] += 1
            updates.append((node.level, parent_id, idx, node.node_id))
            if node.children:
                _dfs(node.children, node.node_id)

    _dfs(body.node_ids, None)

    if not updates:
        return unified_response(400, "node_ids 为空，无需更新")

    # ── 3. 校验所有 node_id 均属于该 record_id ────────────────────────────
    node_id_list = [u[3] for u in updates]
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    f'SELECT id FROM "yxdl_docx_title_trees" '
                    f'WHERE record_id = %s AND id = ANY(%s)',
                    (record_id, node_id_list),
                )
                valid_ids = {row[0] for row in cursor.fetchall()}
    except Exception as e:
        return unified_response(500, f"校验节点归属失败：{str(e)}")

    invalid = [nid for nid in node_id_list if nid not in valid_ids]
    if invalid:
        return unified_response(400, f"以下节点不属于 record_id={record_id}：{invalid}")

    # ── 4. 批量更新 ────────────────────────────────────────────────────────
    current_time = datetime.datetime.now()
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.executemany(
                    """UPDATE "yxdl_docx_title_trees"
                       SET level = %s, parent_id = %s, idx = %s, update_time = %s
                       WHERE id = %s""",
                    [(level, parent_id, idx, current_time, node_id)
                     for level, parent_id, idx, node_id in updates],
                )
                conn.commit()
    except Exception as e:
        return unified_response(500, f"更新树结构失败：{str(e)}")

    return unified_response(200, "树结构更新成功", {
        "record_id": record_id,
        "updated_count": len(updates),
        "update_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
    })