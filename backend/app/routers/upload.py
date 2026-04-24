import datetime
import logging
import os
import re

import aiohttp
from fastapi import APIRouter, Body, File, Request, UploadFile

from app.core.config import UPLOAD_DIR
from app.core.constants import DEFAULT_MAIN_NODE, ProcessMode
from app.converters.docx_converter import docx_to_html
from app.db.database import (
    assign_file_path_to_tree,
    build_eid_path_mapping,
    create_single_main_node,
    get_db_connection,
    get_next_batch_count,
    process_split_tree_nodes,
)
from app.models.schemas import unified_response
from app.utils.file_utils import generate_unique_file_id, generate_unique_filename
from app.utils.html_utils import merge_html_texts
from app.utils.path_utils import local_upload_path_to_web_path, save_html_and_get_url
from file_resp import SplitUpload
from mergfile import TreeItem, call_docx_split, call_set_table_width

router = APIRouter()
logger = logging.getLogger(__name__)

# 前端分片上传时会在文件名前拼接 "{hex}-{timestamp}" 标识，如：
#   de7bc437130692b5c8287b787baa26bb-1776828981534测试_3000段表格数据.xlsx
# 入库前需将该前缀剥离，只保留用户真实文件名。
_FRONTEND_FILENAME_PREFIX_RE = re.compile(r'^[0-9a-f]+-\d+')


def _strip_filename_prefix(file_name: str) -> str:
    """去除前端拼接的 hex-timestamp 前缀，返回用户原始文件名。"""
    return _FRONTEND_FILENAME_PREFIX_RE.sub("", file_name)


# ─── 公共逻辑：保存文件 + 创建上传记录 ─────────────────────────────────────

async def _save_file_and_record(file_content: bytes, original_filename: str, process_mode: str):
    """保存文件到 UPLOAD_DIR，创建数据库记录，返回 (record_id, abs_file_path)"""
    new_filename = generate_unique_filename(original_filename)
    abs_file_path = os.path.abspath(os.path.join(UPLOAD_DIR, new_filename))

    with open(abs_file_path, "wb") as f:
        f.write(file_content)

    current_time = datetime.datetime.now()
    insert_sql = """
    INSERT INTO "yxdl_docx_upload_records"
    (original_filename, new_filename, save_path, upload_time, update_time, split_file_id, process_mode)
    VALUES (%s, %s, %s, %s, %s, %s, %s)
    RETURNING id;
    """
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(insert_sql, (
                original_filename, new_filename, abs_file_path,
                current_time, current_time, "", process_mode,
            ))
            record_id = cursor.fetchone()[0]
            conn.commit()
    return record_id, abs_file_path, current_time


async def _get_file_content_from_source(file_source_type: str, file_source: str):
    """根据来源类型（url / static）获取文件内容和文件名"""
    if file_source_type == "url":
        async with aiohttp.ClientSession() as session:
            async with session.get(file_source) as response:
                if response.status != 200:
                    return None, None, f"下载文件失败，HTTP状态码：{response.status}"
                content_disposition = response.headers.get("Content-Disposition", "")
                if "filename=" in content_disposition:
                    original_filename = content_disposition.split("filename=")[-1].strip("'\"")
                else:
                    original_filename = file_source.split("/")[-1]
                return await response.read(), original_filename, None

    elif file_source_type == "static":
        static_file_path = os.path.abspath(file_source)
        if not os.path.exists(static_file_path):
            return None, None, f"静态文件不存在：{static_file_path}"
        if not static_file_path.lower().endswith(".docx"):
            ext = static_file_path.split(".")[-1] if "." in static_file_path else "无后缀"
            return None, None, f"仅支持docx格式文件，当前文件格式：{ext}"
        with open(static_file_path, "rb") as f:
            return f.read(), os.path.basename(static_file_path), None

    return None, None, f"不支持的文件来源类型：{file_source_type}，仅支持url/static"


async def _split_mode(
    file_content: bytes,
    original_filename: str,
    abs_file_path: str,
    record_id: int,
    current_time: datetime.datetime,
):
    """执行 split 模式：调用拆分接口、入库节点"""
    split_file_id = generate_unique_file_id()

    update_sql = """
    UPDATE "yxdl_docx_upload_records"
    SET split_file_id = %s, update_time = %s
    WHERE id = %s;
    """
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(update_sql, (split_file_id, current_time, record_id))
            conn.commit()

    # 调用表格宽度适配接口
    new_file_path = call_set_table_width(abs_file_path)
    # new_file_path = abs_file_path
    with open(new_file_path, "rb") as _f:
        file_bytes = _f.read()

    split_result = call_docx_split(
        file_stream=file_bytes,
        file_name=original_filename,
        file_id=split_file_id,
        had_title=1,
        rm_outline_in_doc=1,
        del_page_break=0,
    )

    title_font_dict = split_result.data.get("title_font_dict") or {}
    if title_font_dict:
        import json
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    'UPDATE "yxdl_docx_upload_records" SET title_font_dict = %s WHERE id = %s',
                    (json.dumps(title_font_dict), record_id),
                )
                conn.commit()

    tree_nodes = [TreeItem(**item) for item in split_result.data.get("tree", [])]
    eid_path_map = build_eid_path_mapping(split_result.data.get("files", []))
    for node in tree_nodes:
        assign_file_path_to_tree(node, eid_path_map)

    batch_count = get_next_batch_count(record_id)
    node_ids = process_split_tree_nodes(
        nodes=tree_nodes,
        record_id=record_id,
        current_time=current_time,
        file_base_path=abs_file_path,
        batch_count=batch_count,
    )
    return split_file_id, node_ids


# ─── 路由 ────────────────────────────────────────────────────────────────────

@router.post("/upload_and_generate_tree", summary="上传文件并生成标题树节点")
async def upload_and_generate_tree(
    file: UploadFile = File(..., description="需要上传的DOCX格式文件"),
    process_mode: ProcessMode = Body("split", description="处理模式：single-单个主节点，split-接口拆分多节点"),
):
    """上传 DOCX 文件并生成标题树节点"""
    file_path = ""
    split_file_id = ""
    try:
        filename = file.filename or ""
        if not filename.lower().endswith(".docx"):
            ext = filename.split(".")[-1] if "." in filename else "无后缀"
            return unified_response(400, f"仅支持docx格式文件，当前文件格式：{ext}")

        file_content = await file.read()
        record_id, abs_file_path, current_time = await _save_file_and_record(
            file_content, filename, process_mode
        )
        file_path = abs_file_path

        if process_mode == "single":
            node_id = create_single_main_node(record_id, current_time, abs_file_path)
            return unified_response(200, "文件上传成功，生成单个主节点", {
                "record_id": record_id,
                "process_mode": process_mode,
                "original_filename": filename,
                "file_path": abs_file_path,
                "create_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                "node_count": 1,
                "node_ids": node_id,
                "node_type": "main",
                "split_file_id": "",
                "split_files": [],
                "node_level": DEFAULT_MAIN_NODE["level"],
                "node_eid": DEFAULT_MAIN_NODE["eid"],
                "node_idx": DEFAULT_MAIN_NODE["idx"],
                "tips": "可使用node_id调用查询接口获取HTML文本",
            })

        split_file_id, node_ids = await _split_mode(
            file_content, filename, abs_file_path, record_id, current_time
        )
        return unified_response(200, f"文件上传拆分成功，共生成{len(node_ids)}个分支节点", {
            "record_id": record_id,
            "node_ids": node_ids,
            "node_type": "branch",
            "split_file_id": split_file_id,
        })

    except Exception as e:
        try:
            if file_path and os.path.exists(file_path):
                os.remove(file_path)
        except Exception:
            pass
        return unified_response(500, f"上传处理失败（模式：{process_mode}）：{str(e)}", {
            "process_mode": process_mode, "split_file_id": split_file_id
        })


@router.post("/route_generate_tree", summary="文件路径生成标题树节点")
async def route_generate_tree(
    file_source_type: str = Body("url", description="文件来源类型：url-从URL下载，static-从静态路径读取"),
    file_source: str = Body(..., description="文件来源：URL地址 或 服务器静态文件路径"),
    process_mode: ProcessMode = Body("split", description="处理模式：single-单个主节点，split-接口拆分多节点"),
):
    """获取文件（URL下载/静态路径读取）并生成标题树节点"""
    file_path = ""
    split_file_id = ""
    try:
        file_content, original_filename, error = await _get_file_content_from_source(
            file_source_type, file_source
        )
        if error:
            return unified_response(400, error)

        if not original_filename.lower().endswith(".docx"):
            ext = original_filename.split(".")[-1] if "." in original_filename else "无后缀"
            return unified_response(400, f"仅支持docx格式文件，当前文件格式：{ext}")

        record_id, abs_file_path, current_time = await _save_file_and_record(
            file_content, original_filename, process_mode
        )
        file_path = abs_file_path

        if process_mode == "single":
            node_id = create_single_main_node(record_id, current_time, abs_file_path)
            return unified_response(200, "文件获取成功，生成单个主节点", {
                "record_id": record_id,
                "process_mode": process_mode,
                "original_filename": original_filename,
                "file_path": abs_file_path,
                "create_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                "node_count": 1,
                "node_ids": node_id,
                "node_type": "main",
                "split_file_id": "",
                "split_files": [],
                "node_level": DEFAULT_MAIN_NODE["level"],
                "node_eid": DEFAULT_MAIN_NODE["eid"],
                "node_idx": DEFAULT_MAIN_NODE["idx"],
                "tips": "可使用node_id调用查询接口获取HTML文本",
            })

        split_file_id, node_ids = await _split_mode(
            file_content, original_filename, abs_file_path, record_id, current_time
        )
        return unified_response(200, f"文件获取拆分成功，共生成{len(node_ids)}个分支节点", {
            "record_id": record_id,
            "node_ids": node_ids,
            "node_type": "branch",
            "split_file_id": split_file_id,
        })

    except Exception as e:
        try:
            if file_path and os.path.exists(file_path):
                os.remove(file_path)
        except Exception:
            pass
        return unified_response(500, f"文件处理失败（模式：{process_mode}）：{str(e)}", {
            "process_mode": process_mode, "split_file_id": split_file_id
        })


@router.post("/route_docx2html_marge", summary="文件路径docx转化html")
async def route_docx2html_marge(
    file_source_type: str = Body("url", description="文件来源类型：url-从URL下载，static-从静态路径读取"),
    file_source: str = Body(..., description="文件来源：URL地址 或 服务器静态文件路径"),
):
    """获取 DOCX 文件并转换为合并 HTML"""
    file_path = ""
    split_file_id = ""

    try:
        file_content, original_filename, error = await _get_file_content_from_source(
            file_source_type, file_source
        )
        if error:
            return unified_response(400, error)

        if not original_filename.lower().endswith(".docx"):
            ext = original_filename.split(".")[-1] if "." in original_filename else "无后缀"
            return unified_response(400, f"仅支持docx格式文件，当前文件格式：{ext}")

        record_id, abs_file_path, current_time = await _save_file_and_record(
            file_content, original_filename, "split"
        )
        file_path = abs_file_path

        split_file_id = generate_unique_file_id()
        update_sql = """
        UPDATE "yxdl_docx_upload_records"
        SET split_file_id = %s, update_time = %s
        WHERE id = %s;
        """
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(update_sql, (split_file_id, current_time, record_id))
                conn.commit()
        new_file_path = call_set_table_width(abs_file_path)

        with open(new_file_path, "rb") as _f:
            file_bytes = _f.read()

        split_result = call_docx_split(
            file_stream=file_bytes,
            file_name=original_filename,
            file_id=split_file_id,
            had_title=1,
            rm_outline_in_doc=1,
            del_page_break=0,
        )

        html_list = []
        for file__ in split_result.data.get("files", []):
            html_content, _ = docx_to_html(file__)
            html_list.append(html_content)
        total_html_content = merge_html_texts(html_list)

        return unified_response(200, "文件html转换成功", {
            "http_path": save_html_and_get_url(total_html_content),
        })

    except Exception as e:
        try:
            if file_path and os.path.exists(file_path):
                os.remove(file_path)
        except Exception:
            pass
        return unified_response(500, f"文件处理失败：{str(e)}", {"split_file_id": split_file_id})


@router.post("/split_uploads", summary="分片上传文件")
async def split_upload_and_generate_tree(
    request: Request,
    file: UploadFile = File(..., description="需要上传的分片文件"),
    file_no: str = Body(..., description="分片编号"),
    file_sign: str = Body(..., description="文件唯一标识"),
    file_name: str = Body(..., description="原始文件名"),
    files_total_count: str = Body(..., description="分片总数"),
):
    """文件分片上传接口"""
    try:
        full_file_name = f"{file_sign}_{file_name}"
        abs_file_path = os.path.abspath(UPLOAD_DIR)
        path_ = local_upload_path_to_web_path(abs_file_path, request)
        sep = os.path.sep
        real_file_path = (
            abs_file_path + sep + full_file_name
            if not abs_file_path.endswith(sep)
            else abs_file_path + full_file_name
        )

        file_content = await file.read()
        status, result, msg = SplitUpload(
            UPLOAD_DIR, file_no, full_file_name, files_total_count, file_sign, file_content
        )

        if status == 0:
            logger.debug(f"finish upload. sign:{file_sign} | file_no:{file_no} | total:{files_total_count}")
        else:
            logger.warning(f"upload fail. error:{msg} | sign:{file_sign} | file_no:{file_no}")

        if result == 1:
            file_path = f"{path_}{full_file_name}"
            # 分片合并完成，写入 xlsx 上传记录表
            try:
                from app.db.database import insert_xlsx_upload_record
                insert_xlsx_upload_record(
                    original_filename=_strip_filename_prefix(file_name),
                    new_filename=full_file_name,
                    file_sign=file_sign,
                    save_path=real_file_path,
                )
            except Exception as db_err:
                logger.warning(f"split_uploads: 写入 xlsx 上传记录失败 err={db_err}")
        else:
            file_path = ""
            real_file_path = ""

        from fastapi.responses import JSONResponse
        return JSONResponse(content={
            "status": status,
            "is_finish": result,
            "msg": msg,
            "data": {"file_path": file_path, "real_file_path": real_file_path},
        })

    except Exception as e:
        logger.debug(f"TemplateUploadFile-失败：{e}")
        from fastapi.responses import JSONResponse
        return JSONResponse(status_code=500, content={"status": 1, "is_finish": 0, "msg": "接口异常", "data": ""})