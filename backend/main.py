from fastapi import FastAPI, UploadFile, File, Body, Request
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.exceptions import RequestValidationError
from fastapi.encoders import jsonable_encoder
import os
import sys
import datetime
import random
import string
import psycopg2
from psycopg2 import OperationalError, ProgrammingError
from psycopg2.extras import RealDictCursor  # 用于返回字典格式结果
from contextlib import contextmanager
from typing import Optional
import io  # 用于生成内存文件流

from docxhtmlcoverter import DocxHtmlConverter

# ====================== 配置项 ======================
# FastAPI配置
app = FastAPI(title="DOCX文件上传&HTML互转接口", version="1.0")

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
    "port": 15400,  # PostgreSQL默认端口
    "user": "dev_scxx",  # 默认用户
    "password": "scxx7233Cc",  # 替换为你的PostgreSQL密码
    "database": "yxdl_zhtb_dev",  # 替换为你的数据库名
    "options": "-c client_encoding=utf8"  # 编码设置
}

# 默认主节点配置
DEFAULT_MAIN_NODE = {
    "title": "文档内容",
    "level": 1
}


# ====================== 统一返回格式工具函数 ======================
def unified_response(code: int, message: str, data: dict = None):
    """生成统一格式的响应体"""
    return JSONResponse(
        status_code=200,  # 所有场景HTTP状态码统一为200，业务状态码在body中
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
    # 区分自定义业务异常和系统异常
    if hasattr(exc, "status_code") and hasattr(exc, "detail"):
        # FastAPI原生HTTPException
        return unified_response(
            code=exc.status_code,
            message=exc.detail,
            data={}
        )
    else:
        # 其他系统异常
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

def word_to_html(file_path: str) -> str:
    """Word转HTML的实现函数"""
    try:
        abs_file_path = os.path.abspath(file_path)
        if not os.path.exists(abs_file_path):
            return f"<p>转换失败：文件不存在: {abs_file_path}</p>"

        # 文件大小检查
        file_size = os.path.getsize(abs_file_path)
        if file_size > 10 * 1024 * 1024:  # 10MB
            print(f"警告：文件过大（{file_size / 1024 / 1024:.2f}MB），可能转换失败")

        converter = DocxHtmlConverter()
        temp_html_filename = generate_unique_filename("temp.html")
        temp_html_path = os.path.join(UPLOAD_DIR, temp_html_filename)

        # 执行DOCX转HTML
        converter.docx_to_single_html(abs_file_path, temp_html_path)

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
                    os.remove(temp_html_path)
                except Exception as e:
                    print(f"警告：无法删除临时文件 {temp_html_path} - {e}")

        return html_content or ""
    except Exception as e:
        print(f"Word转HTML失败: {str(e)}")
        return f"<p>转换失败：{str(e)}</p>"

def html_to_docx(html_content: str) -> tuple[bool, io.BytesIO | str]:
    """HTML转DOCX的实现函数"""
    try:
        if not html_content.strip():
            return False, "HTML内容不能为空"

        converter = DocxHtmlConverter()
        temp_docx_filename = generate_unique_filename("html2docx.docx")
        temp_docx_path = os.path.join(UPLOAD_DIR, temp_docx_filename)

        # 执行HTML转DOCX
        converter.html_text_to_docx(html_content, temp_docx_path)

        # 检查文件是否生成
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
        yield conn
    except Exception as e:
        raise Exception(f"数据库操作异常：{str(e)}")
    finally:
        if conn:
            conn.close()

def init_db_tables():
    """初始化PostgreSQL数据表（使用指定的建表语句）"""
    # 1. 文件上传记录表（先删后建）
    create_file_table_sql = """
    DROP TABLE IF EXISTS "yxdl_docx_upload_records";
    CREATE TABLE "yxdl_docx_upload_records" (
      "id" SERIAL PRIMARY KEY,
      "original_filename" varchar(255) COLLATE "pg_catalog"."default" NOT NULL,
      "new_filename" varchar(255) COLLATE "pg_catalog"."default" NOT NULL,
      "save_path" varchar(512) COLLATE "pg_catalog"."default" NOT NULL,
      "upload_time" TIMESTAMP NOT NULL,
      "update_time" TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
    COMMENT ON COLUMN "yxdl_docx_upload_records"."id" IS '记录ID';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."original_filename" IS '原始文件名';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."new_filename" IS '新文件名';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."save_path" IS '文件保存路径';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."upload_time" IS '上传时间';
    COMMENT ON COLUMN "yxdl_docx_upload_records"."update_time" IS '更新时间';
    COMMENT ON TABLE "yxdl_docx_upload_records" IS 'DOCX文件上传记录';
    """

    # 2. 标题树节点表（先删后建）
    create_title_tree_table_sql = """
    DROP TABLE IF EXISTS "yxdl_docx_title_trees";
    CREATE TABLE "yxdl_docx_title_trees" (
      "id" SERIAL PRIMARY KEY,
      "record_id" int4,
      "title_text" varchar(512) COLLATE "pg_catalog"."default" NOT NULL,
      "html_content" text COLLATE "pg_catalog"."default",
      "create_time" TIMESTAMP NOT NULL,
      "update_time" TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
    COMMENT ON COLUMN "yxdl_docx_title_trees"."id" IS '节点ID';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."record_id" IS '关联文件记录ID';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."title_text" IS '标题文本';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."html_content" IS 'Word转换后的HTML文本';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."create_time" IS '创建时间';
    COMMENT ON COLUMN "yxdl_docx_title_trees"."update_time" IS '更新时间';
    COMMENT ON TABLE "yxdl_docx_title_trees" IS '标题树节点表';
    """

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                # 执行建表语句（先删后建）
                cursor.execute(create_file_table_sql)
                cursor.execute(create_title_tree_table_sql)
                conn.commit()
        print("PostgreSQL数据表初始化成功（使用指定的建表语句）")
    except Exception as e:
        print(f"PostgreSQL数据表初始化失败：{str(e)}")
        print("警告：数据库初始化失败，部分功能将不可用")

# 初始化数据表
# init_db_tables()

# ====================== 核心接口 ======================
@app.post("/doc_editor/upload_and_generate_tree", summary="上传文件并生成标题树节点")
async def upload_and_generate_tree(
        file: UploadFile = File(..., description="需要上传的DOCX格式文件")
):
    """上传DOCX文件后自动生成标题树主节点，返回节点ID"""
    file_path = ""  # 初始化文件路径变量
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

        # 3. 插入文件记录到数据库（更新表名）
        current_time = datetime.datetime.now()
        insert_file_sql = """
        INSERT INTO "yxdl_docx_upload_records" 
        (original_filename, new_filename, save_path, upload_time, update_time)
        VALUES (%s, %s, %s, %s, %s)
        RETURNING id;  -- PostgreSQL获取自增ID的方式
        """

        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(insert_file_sql, (
                    original_filename,
                    new_filename,
                    abs_file_path,
                    current_time,
                    current_time
                ))
                record_id = cursor.fetchone()[0]  # 获取返回的ID

                # 4. 调用Word转HTML函数
                html_content = word_to_html(abs_file_path)

                # 5. 插入标题树节点（主节点，更新表名）
                insert_tree_sql = """
                INSERT INTO "yxdl_docx_title_trees" 
                (record_id, title_text, html_content, create_time, update_time)
                VALUES (%s, %s, %s, %s, %s)
                RETURNING id;
                """
                cursor.execute(insert_tree_sql, (
                    record_id,
                    DEFAULT_MAIN_NODE["title"],
                    html_content,
                    current_time,
                    current_time
                ))
                node_id = cursor.fetchone()[0]  # 获取返回的ID
                conn.commit()

        # 6. 返回结果
        return unified_response(
            code=200,
            message="文件上传成功并生成标题树节点",
            data={
                "node_id": node_id,
                "record_id": record_id,
                "original_filename": original_filename,
                "file_path": abs_file_path,
                "create_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                "update_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                "tips": "可使用node_id调用查询接口获取HTML文本"
            }
        )

    except Exception as e:
        # 清理临时文件
        try:
            if file_path and os.path.exists(file_path):
                os.remove(file_path)
        except Exception as cleanup_e:
            print(f"清理临时文件失败：{cleanup_e}")

        return unified_response(
            code=500,
            message=f"上传并生成节点失败：{str(e)}",
            data={}
        )

@app.get("/doc_editor/get_html_by_node/{node_id}", summary="根据节点ID查询HTML文本")
async def get_html_by_node(node_id: int):
    """根据标题树节点ID查询存储的HTML文本（返回更新时间）"""
    try:
        # 查询节点的HTML内容（更新表名）
        select_sql = """
        SELECT t.html_content, t.title_text, t.create_time, t.update_time, 
               r.original_filename, r.upload_time, r.update_time as file_update_time
        FROM "yxdl_docx_title_trees" t
        LEFT JOIN "yxdl_docx_upload_records" r ON t.record_id = r.id
        WHERE t.id = %s
        """

        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cursor:  # 返回字典格式
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

        # 返回HTML文本
        return unified_response(
            code=200,
            message="查询HTML文本成功",
            data={
                "node_id": node_id,
                "title_text": result["title_text"],
                "original_filename": result["original_filename"],
                "html_content": result["html_content"],
                "node_create_time": format_time(result["create_time"]),
                "node_update_time": format_time(result["update_time"]),
                "file_upload_time": format_time(result["upload_time"]),
                "file_update_time": format_time(result["file_update_time"])
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
):
    """更新指定节点ID的HTML文本（更新update_time字段）"""
    try:
        # 1. 基础参数校验
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

        # 2. 构建动态更新SQL（更新表名）
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

        # 3. 数据库操作
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                # 校验节点是否存在（更新表名）
                cursor.execute("SELECT id FROM \"yxdl_docx_title_trees\" WHERE id = %s", (node_id,))
                if not cursor.fetchone():
                    return unified_response(
                        code=404,
                        message=f"未找到ID为{node_id}的标题树节点",
                        data={}
                    )

                # 执行更新
                affected_rows = cursor.execute(update_sql, tuple(update_values))
                conn.commit()

        # 4. 返回结果
        return unified_response(
            code=200,
            message="节点HTML内容更新成功",
            data={
                "node_id": node_id,
                "updated_title": "标题更新为"+title_text.strip() if title_text else "标题未更新",
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

@app.post("/doc_editor/html_to_docx", summary="HTML转DOCX文件流")
async def html_to_docx_api(
        node_id: int = Body(..., description="要更新的节点ID"),
        html_content: str = Body(..., description="需要转换的HTML文本"),
        filename: Optional[str] = Body("output.docx", description="下载的DOCX文件名（默认output.docx）"),
        title_text: Optional[str] = Body(None, description="可选：更新节点标题文本")
):
    """接收HTML文本生成DOCX文件流（更新节点update_time）"""
    try:
        # 1. 基础参数校验
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

        # 2. 构建动态更新SQL（更新表名）
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

        # 3. 数据库操作
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                # 校验节点是否存在（更新表名）
                cursor.execute("SELECT id FROM \"yxdl_docx_title_trees\" WHERE id = %s", (node_id,))
                if not cursor.fetchone():
                    return unified_response(
                        code=404,
                        message=f"未找到ID为{node_id}的标题树节点",
                        data={}
                    )

                # 执行更新
                affected_rows = cursor.execute(update_sql, tuple(update_values))
                conn.commit()

        # 4. 校验文件名格式
        if not filename.lower().endswith(".docx"):
            filename += ".docx"
        filename = filename.replace('/', '_').replace('\\', '_').replace(':', '_')

        # 5. 调用HTML转DOCX函数
        success, result = html_to_docx(html_content)
        if not success:
            return unified_response(
                code=500,
                message=result,
                data={}
            )

        # 6. 构造响应头（包含更新时间）
        headers = {
            "Content-Disposition": f"attachment; filename={filename}",
            "Access-Control-Expose-Headers": "Content-Disposition",
            "X-Update-Time": current_time.strftime("%Y-%m-%d %H:%M:%S")
        }

        # 7. 返回文件流
        response = StreamingResponse(
            result,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers
        )
        # 为文件流响应补充统一格式的提示（前端可通过响应头或返回体解析）
        response.status_code = 200
        return response

    except Exception as e:
        return unified_response(
            code=500,
            message=f"HTML转DOCX失败：{str(e)}",
            data={}
        )

# ====================== 启动服务 ======================
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        app=__name__ + ":app",
        host="0.0.0.0",
        port=8080,
        reload=True
    )