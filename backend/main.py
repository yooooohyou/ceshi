from fastapi import FastAPI, UploadFile, File, Body, Request, HTTPException
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.exceptions import RequestValidationError
import os
import datetime
import random
import string
import psycopg2
from psycopg2.extras import RealDictCursor
from contextlib import contextmanager
from typing import Optional, Tuple, Union, List, Dict, Any, Literal
import io
from docxautogenerator import generate_fully_centered_patent_doc
from mergfile import call_docx_split,call_docx_merge, TreeItem,MergeRequest
import json

from docxhtmlcoverter import DocxHtmlConverter

# ====================== 配置项 ======================
app = FastAPI(title="DOCX文件上传&HTML转换接口", version="1.0")

# 基础路径配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
# 确保目录存在（增加权限检查）
try:
    os.makedirs(UPLOAD_DIR, exist_ok=True)
except PermissionError:
    UPLOAD_DIR = os.path.join(os.gettempdir(), "docx_uploads")
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    print(f"警告：无法在当前目录创建uploads文件夹，已切换到系统临时目录：{UPLOAD_DIR}")

# PostgreSQL数据库配置（请替换为你的实际配置）
POSTGRES_CONFIG = {
    "host": "10.13.6.59",
    "port": 15400,
    "user": "dev_scxx",
    "password": "scxx7233Cc",
    "database": "yxdl_zhtb_dev",
    "options": "-c client_encoding=utf8"
}

# 默认主节点配置
DEFAULT_MAIN_NODE = {
    "title": "文档内容",
    "level": 1,
    "eid": "main_node",
    "idx": 0
}

# 处理模式定义（默认改为split）
ProcessMode = Literal["single", "split"]



class HTMLSafeJSONEncoder(json.JSONEncoder):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.ensure_ascii = False  # 保留中文等非ASCII字符
        self.escape_forward_slashes = False  # 禁用/的转义

# ====================== 统一返回格式工具函数 ======================
def unified_response(code: int, message: str, data: dict = None) -> JSONResponse:
    """生成统一格式的响应体"""
    return JSONResponse(
        status_code=200,
        content={
            "code": code,
            "message": message,
            "data": data or {}
        }
    )


# ====================== 全局异常处理 ======================
@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    """处理参数校验异常"""
    return unified_response(
        code=400,
        message="参数校验失败",
        data={"errors": exc.errors()}
    )


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    """处理所有未捕获的异常"""
    if hasattr(exc, "status_code") and hasattr(exc, "detail"):
        return unified_response(
            code=exc.status_code,
            message=exc.detail,
            data={}
        )
    else:
        return unified_response(
            code=500,
            message=f"服务器内部错误：{str(exc)}",
            data={}
        )


# ====================== 核心工具函数 ======================
def generate_unique_filename(original_filename: str) -> str:
    """生成唯一文件名，避免冲突"""
    if "." in original_filename:
        file_ext = os.path.splitext(original_filename)[-1].lower()
    else:
        file_ext = ".docx"

    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    random_str = ''.join(random.choices(string.ascii_letters + string.digits, k=6))
    return f"{timestamp}_{random_str}{file_ext}"


def generate_unique_file_id() -> str:
    """生成唯一的file_id，用于拆分接口调用"""
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
    random_str = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
    return f"docx_{timestamp}_{random_str}"


def docx_to_html(file_path: str) -> str:
    """Word转HTML的实现函数"""
    try:
        # abs_file_path = os.path.abspath(file_path)
        # if not os.path.exists(abs_file_path):
        #     return f"<p>转换失败：文件不存在: {abs_file_path}</p>"

        # 文件大小检查
        file_size = os.path.getsize(file_path)
        if file_size > 10 * 1024 * 1024:
            print(f"警告：文件过大（{file_size / 1024 / 1024:.2f}MB），可能转换失败")

        converter = DocxHtmlConverter()
        temp_html_filename = generate_unique_filename("temp.html")
        temp_html_path = os.path.join(UPLOAD_DIR, temp_html_filename)

        # 执行DOCX转HTML
        converter.docx_to_single_html(file_path, temp_html_path)

        # 读取并返回HTML内容
        html_content = ""
        if os.path.exists(temp_html_path):
            try:
                with open(temp_html_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
            except UnicodeDecodeError:
                with open(temp_html_path, 'r', encoding='gbk') as f:
                    html_content = f.read()
            finally:
                try:
                    # os.remove(temp_html_path)
                    pass
                except Exception as e:
                    print(f"警告：无法删除临时文件 {temp_html_path} - {e}")

        return html_content or ""
    except Exception as e:
        print(f"Word转HTML失败: {str(e)}")
        return f"<p>转换失败：{str(e)}</p>"


def convert_html_to_docx(html_content: str) -> Tuple[bool, Union[io.BytesIO, str]]:
    """HTML转DOCX的实现函数"""
    try:
        if not html_content.strip():
            return False, "HTML内容不能为空"

        converter = DocxHtmlConverter()
        temp_docx_filename = generate_unique_filename("html2docx.docx")
        temp_docx_path = os.path.join(UPLOAD_DIR, temp_docx_filename)

        converter.html_text_to_docx(html_content, temp_docx_path)

        if not os.path.exists(temp_docx_path):
            return False, f"转换失败：未生成文件 {temp_docx_path}"

        # 读取文件到内存流
        docx_stream = io.BytesIO()
        with open(temp_docx_path, 'rb') as f:
            docx_stream.write(f.read())
        docx_stream.seek(0)

        # 删除临时文件
        try:
            if os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)
        except Exception as e:
            print(f"警告：无法删除临时文件 {temp_docx_path} - {e}")

        return True, docx_stream
    except PermissionError:
        return False, "权限错误：无法创建/读取临时文件"
    except Exception as e:
        return False, f"HTML转DOCX失败：{str(e)}"


# ====================== 数据库工具函数 ======================
@contextmanager
def get_db_connection():
    """PostgreSQL数据库连接上下文管理器"""
    conn = None
    try:
        conn = psycopg2.connect(**POSTGRES_CONFIG)
        conn.autocommit = False
        yield conn
    except Exception as e:
        if conn:
            conn.rollback()
        raise Exception(f"数据库操作异常：{str(e)}")
    finally:
        if conn:
            conn.close()


def init_db_tables():
    """初始化PostgreSQL数据表"""
    # 1. 文件上传记录表
    create_file_table_sql = """
    DROP TABLE IF EXISTS "yxdl_docx_upload_records";
    CREATE TABLE "yxdl_docx_upload_records" (
      "id" SERIAL PRIMARY KEY,
      "original_filename" varchar(255) COLLATE "pg_catalog"."default" NOT NULL,
      "new_filename" varchar(255) COLLATE "pg_catalog"."default" NOT NULL,
      "save_path" varchar(512) COLLATE "pg_catalog"."default" NOT NULL,
      "upload_time" TIMESTAMP NOT NULL,
      "update_time" TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
      "split_file_id" varchar(128) COLLATE "pg_catalog"."default" COMMENT '拆分接口使用的file_id',
      "process_mode" varchar(16) COLLATE "pg_catalog"."default" DEFAULT 'split' COMMENT '处理模式：single/split'
    );
    COMMENT ON COLUMN "yxdl_docx_upload_records"."id" IS '记录ID';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."original_filename" IS '原始文件名';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."new_filename" IS '新文件名';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."save_path" IS '文件保存路径';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."upload_time" IS '上传时间';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."update_time" IS '更新时间';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."split_file_id" IS '拆分接口使用的file_id';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."process_mode" IS '处理模式：single/split';
    COMMENT ON TABLE "yxdl_docx_upload_records" IS 'DOCX文件上传记录';
    """

    # 2. 标题树节点表
    create_title_tree_table_sql = """
    DROP TABLE IF EXISTS "yxdl_docx_title_trees";
    CREATE TABLE "yxdl_docx_title_trees" (
      "id" SERIAL PRIMARY KEY,
      "record_id" int4,
      "title_text" varchar(512) COLLATE "pg_catalog"."default" NOT NULL,
      "html_content" text COLLATE "pg_catalog"."default",
      "create_time" TIMESTAMP NOT NULL,
      "update_time" TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
      "level" int4 COMMENT '节点层级',
      "eid" varchar(128) COLLATE "pg_catalog"."default" COMMENT '拆分接口返回的eid',
      "idx" int4 COMMENT '拆分接口返回的idx',
      "node_type" varchar(16) COLLATE "pg_catalog"."default" DEFAULT 'main' COMMENT '节点类型：main/branch'
    );
    COMMENT ON COLUMN "yxdl_docx_title_trees"."id" IS '节点ID';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."record_id" IS '关联文件记录ID';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."title_text" IS '标题文本';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."html_content" IS 'Word转换后的HTML文本';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."create_time" IS '创建时间';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."update_time" IS '更新时间';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."level" IS '节点层级';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."eid" IS '拆分接口返回的eid';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."idx" IS '拆分接口返回的idx';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."node_type" IS '节点类型：main/branch';
    COMMENT ON TABLE "yxdl_docx_title_trees" IS '标题树节点表';
    """

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(create_file_table_sql)
                cursor.execute(create_title_tree_table_sql)
                conn.commit()
        print("PostgreSQL数据表初始化成功")
    except Exception as e:
        print(f"PostgreSQL数据表初始化失败：{str(e)}")
        print("警告：数据库初始化失败，部分功能将不可用")


# 初始化数据表（首次运行取消注释）
# init_db_tables()

# ====================== 工具函数：处理拆分节点 ======================

def build_eid_path_mapping(files: List[str]) -> Dict[str, str]:
    """
    构建 eid 到文件路径的映射关系
    :param files: 文件路径列表
    :return: {eid: 文件路径} 的字典
    """
    eid_path_map = {}
    for file_path in files:
        # 从路径中提取文件名（不含后缀）作为 eid
        file_name = file_path.split("/")[-1]  # 取最后一段路径（文件名）
        eid = file_name.rsplit(".", 1)[0]     # 去掉后缀（如 .docx）
        eid_path_map[eid] = file_path
    return eid_path_map

def assign_file_path_to_tree(node: TreeItem, eid_path_map: Dict[str, str]):
    """
    递归为树节点分配对应的文件路径
    :param node: 树节点对象
    :param eid_path_map: eid-路径映射字典
    """
    # 为当前节点赋值文件路径
    if node.eid in eid_path_map:
        node.file_path = eid_path_map[node.eid]
    # 递归处理子节点
    if node.children:
        for child in node.children:
            assign_file_path_to_tree(child, eid_path_map)


def process_split_tree_nodes(
        nodes: List[TreeItem],
        record_id: int,
        current_time: datetime.datetime,
        file_base_path: str
) -> List[Dict[str, Any]]:
    """
    递归处理拆分后的树节点，转换HTML并入库，返回带层级结构的节点信息

    Args:
        nodes: 树节点列表
        record_id: 记录ID
        current_time: 当前时间
        file_base_path: 文件基础路径

    Returns:
        嵌套结构的节点列表，格式：
        [
            {
                "name": "节点标题",
                "node_id": 数据库ID,
                "children": [
                    {
                        "name": "子节点标题",
                        "node_id": 子节点数据库ID,
                        "children": [...]
                    }
                ]
            }
        ]
    """
    # 参数校验
    if not isinstance(nodes, list) or not nodes:
        # logger.warning("传入的节点列表为空或不是列表类型")
        return []

    if not isinstance(record_id, int) or record_id <= 0:
        # logger.error(f"无效的record_id: {record_id}")
        raise ValueError("record_id必须是正整数")

    result_nodes = []

    for node in nodes:
        if not isinstance(node, TreeItem):
            # logger.warning(f"无效的节点类型: {type(node)}，跳过处理")
            continue
        try:
            # 1. 生成节点标题
            node_title = node.text.strip() if (node.text and isinstance(node.text,
                                                                        str)) else f"节点_{node.eid or '未知'}_{node.level}_{node.idx}"


            # 2. 转换该节点为HTML
            # node_file_path = f"{file_base_path}_{node.eid or node.idx}" if node.eid else file_base_path
            node_file_path = node.file_path
            file_name = node_file_path
            level = node.level
            # html_content = docx_to_html(node_file_path, node.text or "")
            html_content = ""

            # 3. 插入数据库
            insert_tree_sql = """
            INSERT INTO "yxdl_docx_title_trees" 
            (record_id, title_text, html_content, create_time, update_time, level, eid, idx, node_type, origin_file_path, is_conversion_completion)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id;
            """

            with get_db_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(insert_tree_sql, (
                        record_id,
                        node_title,
                        html_content,
                        current_time,
                        current_time,
                        node.level,
                        node.eid,
                        node.idx,
                        "branch",
                        node_file_path,
                        0
                    ))
                    node_id = cursor.fetchone()[0]
                    conn.commit()

            # 4. 构建当前节点的返回字典
            current_node = {
                "name": node_title,
                "node_id": node_id,
                "level": level,
                "file_name": file_name,
                "children": []  # 初始化子节点列表
            }
            # logger.info(f"成功创建分支节点：ID={node_id}, EID={node.eid}, 标题={node_title}")

            # 5. 递归处理子节点并赋值给children
            if node.children:
                child_nodes = process_split_tree_nodes(
                    node.children,
                    record_id,
                    current_time,
                    file_base_path
                )
                current_node["children"] = child_nodes

            # 6. 将当前节点添加到结果列表
            result_nodes.append(current_node)

        except ValueError as ve:
            # logger.error(f"节点参数错误（eid={node.eid}）：{str(ve)}", exc_info=True)
            continue
        except Exception as e:
            # logger.error(f"处理节点失败（eid={node.eid}）：{str(e)}", exc_info=True)
            continue

    return result_nodes


# ====================== 工具函数：创建单个主节点 ======================
def create_single_main_node(
        record_id: int,
        current_time: datetime.datetime,
        file_path: str
) -> int:
    """
    创建单个主节点（原逻辑）- 修正SQL适配最新表结构
    :return: 主节点ID
    """
    # 1. 转换整个文档为HTML
    html_content = docx_to_html(file_path)

    # 2. 插入主节点到数据库（完整字段）
    insert_tree_sql = """
    INSERT INTO "yxdl_docx_title_trees" 
    (
        record_id, 
        title_text, 
        html_content, 
        create_time, 
        update_time, 
        level, 
        eid, 
        idx, 
        node_type
    )
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
    RETURNING id;
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(insert_tree_sql, (
                record_id,  # record_id
                DEFAULT_MAIN_NODE["title"],  # title_text
                html_content,  # html_content
                current_time,  # create_time
                current_time,  # update_time
                DEFAULT_MAIN_NODE["level"],  # level
                DEFAULT_MAIN_NODE["eid"],  # eid
                DEFAULT_MAIN_NODE["idx"],  # idx
                "main"  # node_type
            ))
            node_id = cursor.fetchone()[0]
            conn.commit()

    print(f"成功创建主节点：ID={node_id}, 标题={DEFAULT_MAIN_NODE['title']}")
    return node_id


# ====================== 核心接口：上传文件（默认split，返回结构统一） ======================
@app.post("/doc_editor/upload_and_generate_tree", summary="上传文件并生成标题树节点")
async def upload_and_generate_tree(
        file: UploadFile = File(..., description="需要上传的DOCX格式文件"),
        process_mode: ProcessMode = Body("split", description="处理模式：single-单个主节点，split-接口拆分多节点（默认）")
) -> JSONResponse:
    """
    上传DOCX文件并生成标题树节点
    - 默认模式（不传参数）：split，调用拆分接口生成多个分支节点
    - single模式：生成单个主节点
    """
    file_path = ""
    split_file_id = ""
    try:
        # 1. 校验文件格式
        filename = file.filename or ""
        if not filename.lower().endswith(".docx"):
            ext = filename.split('.')[-1] if '.' in filename else '无后缀'
            return unified_response(
                code=400,
                message=f"仅支持docx格式文件，当前文件格式：{ext}",
                data={}
            )

        # 2. 读取并保存文件
        file_content = await file.read()
        original_filename = filename
        new_filename = generate_unique_filename(original_filename)
        file_path = os.path.join(UPLOAD_DIR, new_filename)
        abs_file_path = os.path.abspath(file_path)

        with open(abs_file_path, "wb") as f:
            f.write(file_content)

        # 3. 生成文件记录（默认process_mode改为split）
        current_time = datetime.datetime.now()
        insert_file_sql = """
        INSERT INTO "yxdl_docx_upload_records" 
        (
            original_filename, 
            new_filename, 
            save_path, 
            upload_time, 
            update_time, 
            split_file_id, 
            process_mode
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        RETURNING id;
        """

        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(insert_file_sql, (
                    original_filename,  # original_filename
                    new_filename,  # new_filename
                    abs_file_path,  # save_path
                    current_time,  # upload_time
                    current_time,  # update_time
                    split_file_id,  # split_file_id（初始为空）
                    process_mode  # process_mode（默认split）
                ))
                record_id = cursor.fetchone()[0]
                conn.commit()

        # ====================== 分支1：single模式（返回结构对齐split） ======================
        if process_mode == "single":
            node_id = create_single_main_node(
                record_id=record_id,
                current_time=current_time,
                file_path=abs_file_path
            )

            # 统一返回结构：node_ids为列表，node_count为1
            return unified_response(
                code=200,
                message="文件上传成功，生成单个主节点",
                data={
                    "record_id": record_id,
                    "process_mode": process_mode,
                    "original_filename": original_filename,
                    "file_path": abs_file_path,
                    "create_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                    "node_count": 1,  # 对齐split模式
                    "node_ids": node_id,  # 改为列表格式
                    "node_type": "main",
                    "split_file_id": "",  # 对齐split模式（空值）
                    "split_files": [],  # 对齐split模式（空列表）
                    "node_level": DEFAULT_MAIN_NODE["level"],
                    "node_eid": DEFAULT_MAIN_NODE["eid"],
                    "node_idx": DEFAULT_MAIN_NODE["idx"],
                    "tips": "可使用node_id调用查询接口获取HTML文本"
                }
            )

        # ====================== 分支2：split模式（默认模式） ======================
        elif process_mode == "split":
            # 生成唯一file_id
            split_file_id = generate_unique_file_id()

            # 更新文件记录的split_file_id
            update_file_sql = """
            UPDATE "yxdl_docx_upload_records" 
            SET split_file_id = %s, update_time = %s
            WHERE id = %s;
            """
            with get_db_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(update_file_sql, (split_file_id, current_time, record_id))
                    conn.commit()

            # 调用拆分接口
            split_result = call_docx_split(
                file_stream=file_content,
                file_name=original_filename,
                file_id=split_file_id
            )

            # # 校验拆分结果
            # if split_result.status != 200:
            #     return unified_response(
            #         code=split_result.status,
            #         message=f"文件拆分：{split_result.msg}",
            #         data={"split_data": split_result.dict(), "process_mode": process_mode}
            #     )
            # 处理拆分后的树节点
            # tree_nodes = [TreeItem(**item) for item in split_result.data.get("tree", [])]
            tree_nodes = [TreeItem(**item) for item in split_result.data.get("tree", [])]

            # 2. 构建 eid-文件路径 映射
            eid_path_map = build_eid_path_mapping(split_result.data.get("files", []))

            # 3. 为每个树节点分配文件路径
            for node in tree_nodes:
                assign_file_path_to_tree(node, eid_path_map)
            print(tree_nodes)
            node_ids = process_split_tree_nodes(
                nodes=tree_nodes,
                record_id=record_id,
                current_time=current_time,
                file_base_path=abs_file_path
            )

            # 返回结果（保持原有结构）
            return unified_response(
                code=200,
                message=f"文件上传拆分成功，共生成{len(node_ids)}个分支节点",
                data={
                    "record_id": record_id,
                    "node_ids": node_ids,
                    "node_type": "branch",
                    "split_file_id": split_file_id
                }
            )

    except Exception as e:
        # 清理临时文件
        try:
            pass
            # if file_path and os.path.exists(file_path):
            #     os.remove(file_path)
            # # 清理拆分接口的远端文件
            # if split_file_id and process_mode == "split":
            #     delete_request = DeleteRequest(id=split_file_id)
            #     call_docx_delete(delete_request)
        except Exception as cleanup_e:
            print(f"清理临时文件失败：{cleanup_e}")

        return unified_response(
            code=500,
            message=f"上传处理失败（模式：{process_mode}）：{str(e)}",
            data={"process_mode": process_mode, "split_file_id": split_file_id}
        )


# ====================== 保留原有接口 ======================
@app.get("/doc_editor/get_html_by_node/{node_id}", summary="根据节点ID查询HTML文本")
async def get_html_by_node(node_id: int) -> JSONResponse:
    """根据标题树节点ID查询存储的HTML文本"""
    try:
        select_sql = """
        SELECT t.html_content, t.title_text, t.create_time, t.update_time, t.level, t.eid, t.idx, t.node_type, t.origin_file_path, t.is_conversion_completion,
               r.original_filename, r.upload_time, r.update_time as file_update_time, r.split_file_id, r.process_mode
        FROM "yxdl_docx_title_trees" t
        LEFT JOIN "yxdl_docx_upload_records" r ON t.record_id = r.id
        WHERE t.id = %s
        """

        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cursor:
                cursor.execute(select_sql, (node_id,))
                result = cursor.fetchone()

                if not result:
                    return unified_response(
                        code=404,
                        message=f"未找到ID为{node_id}的标题树节点",
                        data={}
                    )

        # 格式化时间字段
        def format_time(time_obj):
            return time_obj.strftime("%Y-%m-%d %H:%M:%S") if time_obj else ""

        if result["is_conversion_completion"] == 0:
            html_content = docx_to_html(result["origin_file_path"])
            with get_db_connection() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cursor:
                    update_sql = """
                                        UPDATE "yxdl_docx_title_trees"
                                        SET html_content = %s, update_time = NOW()
                                        WHERE id = %s
                                    """
                    cursor.execute(update_sql, (html_content, node_id))
                    conn.commit()  # 提交事务

                    # 3. 重新查询更新后的完整数据（可选，用于返回最新状态）
                    cursor.execute(select_sql, (node_id,))
                    updated_result = cursor.fetchone()
            return unified_response(
                code=200,
                message="查询HTML文本成功",
                data={
                    "node_id": node_id,
                    "title_text": updated_result["title_text"],
                    "level": updated_result["level"],
                    "html_content": updated_result["html_content"]
                }
            )
        else:
            return unified_response(
                code=200,
                message="查询HTML文本成功",
                data={
                    "node_id": node_id,
                    "title_text": result["title_text"],
                    "level": result["level"],
                    "html_content": result["html_content"]
                }
            )

    except Exception as e:
        return unified_response(
            code=500,
            message=f"查询HTML文本失败：{str(e)}",
            data={}
        )


@app.post("/doc_editor/update_html_by_node", summary="更新节点HTML文本")
async def update_html_by_node(
        node_id: int = Body(..., description="要更新的节点ID"),
        html_content: str = Body(..., description="更新后的HTML文本"),
        title_text: Optional[str] = Body(None, description="可选：更新节点标题文本")
) -> JSONResponse:
    """更新指定节点ID的HTML文本"""
    try:
        if node_id <= 0:
            return unified_response(
                code=400,
                message="节点ID必须为正整数",
                data={}
            )
        if not html_content.strip():
            return unified_response(
                code=400,
                message="HTML内容不能为空",
                data={}
            )

        update_fields = ["html_content = %s", "update_time = %s"]
        update_values = [html_content, datetime.datetime.now()]
        current_time = update_values[1]

        if title_text is not None and title_text.strip():
            update_fields.append("title_text = %s")
            update_values.append(title_text.strip())

        update_sql = f"""
        UPDATE "yxdl_docx_title_trees" 
        SET {', '.join(update_fields)}
        WHERE id = %s
        """
        update_values.append(node_id)

        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute("SELECT id FROM \"yxdl_docx_title_trees\" WHERE id = %s", (node_id,))
                if not cursor.fetchone():
                    return unified_response(
                        code=404,
                        message=f"未找到ID为{node_id}的标题树节点",
                        data={}
                    )

                cursor.execute(update_sql, tuple(update_values))
                affected_rows = cursor.rowcount
                conn.commit()

        return unified_response(
            code=200,
            message="节点HTML内容更新成功",
            data={
                "node_id": node_id,
                "updated_title": "标题更新为" + title_text.strip() if title_text else "标题未更新",
                "affected_rows": affected_rows,
                "update_time": current_time.strftime("%Y-%m-%d %H:%M:%S")
            }
        )

    except Exception as e:
        return unified_response(
            code=500,
            message=f"更新节点HTML失败：{str(e)}",
            data={}
        )


@app.post("/doc_editor/html_to_docx", summary="HTML转DOCX文件流", response_model=None)
async def html_to_docx_api(
        node_id: int = Body(..., description="要更新的节点ID"),
        html_content: str = Body(..., description="需要转换的HTML文本"),
        filename: Optional[str] = Body("output.docx", description="下载的DOCX文件名（默认output.docx）"),
        title_text: Optional[str] = Body(None, description="可选：更新节点标题文本")
) -> Union[JSONResponse, StreamingResponse]:
    """接收HTML文本生成DOCX文件流"""
    try:
        if node_id <= 0:
            return unified_response(
                code=400,
                message="节点ID必须为正整数",
                data={}
            )
        if not html_content.strip():
            return unified_response(
                code=400,
                message="HTML内容不能为空",
                data={}
            )

        update_fields = ["html_content = %s", "update_time = %s"]
        update_values = [html_content, datetime.datetime.now()]
        current_time = update_values[1]

        if title_text is not None and title_text.strip():
            update_fields.append("title_text = %s")
            update_values.append(title_text.strip())

        update_sql = f"""
        UPDATE "yxdl_docx_title_trees" 
        SET {', '.join(update_fields)}
        WHERE id = %s
        """
        update_values.append(node_id)

        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute("SELECT id FROM \"yxdl_docx_title_trees\" WHERE id = %s", (node_id,))
                if not cursor.fetchone():
                    return unified_response(
                        code=404,
                        message=f"未找到ID为{node_id}的标题树节点",
                        data={}
                    )

                cursor.execute(update_sql, tuple(update_values))
                conn.commit()

        # 校验文件名格式
        if not filename.lower().endswith(".docx"):
            filename += ".docx"
        filename = os.path.basename(filename).replace('/', '_').replace('\\', '_').replace(':', '_')

        # 调用HTML转DOCX函数
        success, result = convert_html_to_docx(html_content)
        if not success:
            return unified_response(
                code=500,
                message=result,
                data={}
            )

        # 构造响应头
        headers = {
            "Content-Disposition": f"attachment; filename={filename}",
            "Access-Control-Expose-Headers": "Content-Disposition",
            "X-Update-Time": current_time.strftime("%Y-%m-%d %H:%M:%S")
        }

        # 返回文件流
        response = StreamingResponse(
            result,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers
        )
        response.status_code = 200
        return response

    except Exception as e:
        return unified_response(
            code=500,
            message=f"HTML转DOCX失败：{str(e)}",
            data={}
        )


# ====================== 新增：合并接口 ======================
@app.post("/doc_editor/merge_docx", summary="合并拆分的DOCX节点")
async def merge_docx(
        tree: List[Dict[str, Any]] = Body(..., description="节点树结构"),
        files: List[str] = Body(..., description="文件列表")
) -> StreamingResponse:
    """调用合并接口生成合并后的DOCX文件流"""
    try:
        # 构造合并请求
        merge_request = MergeRequest(tree=tree, files=files)

        # 调用合并接口
        merged_file_stream = call_docx_merge(merge_request)

        # 构造响应
        headers = {
            "Content-Disposition": "attachment; filename=merged_docx.docx",
            "Access-Control-Expose-Headers": "Content-Disposition"
        }

        return StreamingResponse(
            io.BytesIO(merged_file_stream),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers
        )
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"文件合并失败：{str(e)}"
        )


@app.post("/doc_editor/generate_patent_doc/default", summary="生成器示例接口")
async def generate_default_patent_doc():
    """使用默认的专利数据和图片路径生成文档并返回下载"""
    try:
        save_path = os.path.join(UPLOAD_DIR, "专利报告.docx")
        DEFAULT_PATENT_DATA = [
            ['1', '发明专利', '一种核岛用防爆电话电缆', 'ZL201210300077.1', '远程电缆股份有限公司', '2015-07-29'],
            ['2', '发明专利', '一种电缆保护夹座安装方法', 'ZL201110454443.4', '远程电缆股份有限公司', '2014-06-25'],
            ['3', '发明专利', '硅烷交联聚乙烯绝缘电缆料及其制造方法', 'ZL200710022494.3', '远程电缆股份有限公司',
             '2010-12-08'],
            ['4', '发明专利', '多色架空绝缘电缆及其制造方法', 'ZL200710019537.2', '远程电缆股份有限公司', '2010-01-13'],
            ['5', '发明专利', '一种用于5G传输技术的配电专用线及其生产工艺', 'ZL201911360899.7', '远程电缆股份有限公司',
             '2020-09-25'],
            ['6', '发明专利', '一种核岛用防爆电话电缆', 'ZL201210300077.1', '远程电缆股份有限公司', '2015-07-29'],
            ['7', '发明专利', '一种电缆保护夹座安装方法', 'ZL201110454443.4', '远程电缆股份有限公司', '2014-06-25'],
            ['8', '发明专利', '硅烷交联聚乙烯绝缘电缆料及其制造方法', 'ZL200710022494.3', '远程电缆股份有限公司',
             '2010-12-08'],
            ['9', '发明专利', '多色架空绝缘电缆及其制造方法', 'ZL200710019537.2', '远程电缆股份有限公司', '2010-01-13'],
            ['10', '发明专利', '一种用于5G传输技术的配电专用线及其生产工艺', 'ZL201911360899.7', '远程电缆股份有限公司',
             '2020-09-25']
        ]

        DEFAULT_CERT_IMG_PATHS = [
            "./pic/图片2.jpg", "./pic/图片3.jpg", "./pic/图片4.jpg",
            "./pic/图片5.jpg", "./pic/图片6.jpg", "./pic/图片7.jpg",
            "./pic/图片8.jpg", "./pic/图片3.jpg", "./pic/图片4.jpg",
            "./pic/图片5.jpg", "./pic/图片6.jpg", "./pic/图片7.jpg",
            "./pic/图片8.jpg"
        ]
        # 生成文档
        generate_fully_centered_patent_doc(
            DEFAULT_PATENT_DATA,
            DEFAULT_CERT_IMG_PATHS,
            save_path,
            last_img_display_mode=1
        )

        # 检查文件是否生成成功
        if not os.path.exists(save_path):
            raise HTTPException(status_code=404, detail="文档生成失败，文件不存在")
        html_content = docx_to_html(save_path)
        print(save_path)
        save_path2 = os.path.join(UPLOAD_DIR, "专利报告.html")
        with open(save_path2, 'w', encoding='utf-8') as f:
            f.write(html_content)
        # try:
        #     if os.path.exists(save_path):
        #         os.remove(save_path)
        # except Exception as e:
        #     print(f"警告：无法删除临时文件 {save_path} - {e}")
        return unified_response(
            code=200,
            message="节点HTML内容更新成功",
            data={
                "html_content": html_content
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成文档时出错: {str(e)}")

# ====================== 启动服务 ======================
if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        app=__name__ + ":app",
        host="0.0.0.0",
        port=8080,
        reload=True
    )