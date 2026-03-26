import base64
import pathlib
import shutil
import time
import urllib

import aiohttp
from fastapi import FastAPI, UploadFile, File, Body, Request, HTTPException, Query
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.exceptions import RequestValidationError
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import datetime
import random
import string
import psycopg2
import configparser
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
import subprocess
import shutil
from psycopg2.extras import RealDictCursor
from contextlib import contextmanager
from typing import Optional, Tuple, Union, List, Dict, Any, Literal
import io
import uuid
import asyncio
from docxautogenerator import generate_fully_centered_patent_doc, generate_report_doc, generate_car_info_doc
from mergfile import call_docx_split,call_docx_merge, TreeItem,MergeRequest
import json
import requests
import os
import tempfile
from urllib.parse import urlparse
import file_resp
import platform
import re
from bs4 import BeautifulSoup
from PIL import Image
from docxhtmlcoverter import DocxHtmlConverter
import logging
from logging.handlers import RotatingFileHandler
import sys
# ====================== 配置项 ======================
def setup_logging():
    log_format = "%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s"
    log_level = logging.INFO
    log_file = "app.log"

    # 配置根日志器
    logging.basicConfig(
        level=log_level,
        format=log_format,
        handlers=[
            logging.StreamHandler(sys.stdout),  # 输出到终端
            RotatingFileHandler(  # 输出到文件（按大小分割）
                log_file,
                maxBytes=10*1024*1024,  # 10MB/文件
                backupCount=5,  # 保留5个备份
                encoding="utf-8"
            )
        ]
    )

    # 调整第三方库日志级别
    logging.getLogger("uvicorn.access").setLevel(logging.WARNING)
    return log_file

# 初始化日志和日志文件路径
log_file_path = setup_logging()
logger = logging.getLogger(__name__)
app = FastAPI(title="DOCX文件上传&HTML转换接口", version="1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 生产环境请替换为具体的允许域名，如["http://localhost:3000", "http://192.168.1.100"]
    allow_credentials=True,
    allow_methods=["*"],  # 允许所有HTTP方法
    allow_headers=["*"],  # 允许所有请求头
)

def read_sc_web_config(config_filename: str = "sc_web.conf") -> configparser.ConfigParser:
    """
    读取conf文件夹下的sc_web.conf配置文件

    Args:
        config_filename: 配置文件名，默认为sc_web.conf

    Returns:
        解析后的配置对象

    Raises:
        FileNotFoundError: 配置文件不存在
        PermissionError: 没有读取文件的权限
        Exception: 其他读取或解析错误
    """
    # 获取配置文件路径（相对当前脚本的路径）
    config_dir = Path(__file__).parent / "conf"
    config_file = config_dir / config_filename

    # 检查配置文件是否存在
    if not config_file.exists():
        raise FileNotFoundError(f"配置文件不存在: {config_file}")

    # 检查是否有读取权限
    if not os.access(config_file, os.R_OK):
        raise PermissionError(f"没有读取权限: {config_file}")

    # 创建配置解析器并读取配置文件
    config = configparser.ConfigParser()
    try:
        config.read(config_file, encoding="utf-8")
    except Exception as e:
        raise Exception(f"解析配置文件失败: {str(e)}")

    return config


# 新增：专门读取server_uploads配置的快捷函数
def get_server_uploads_config() -> dict:
    """
    快捷获取[server_uploads]节的所有配置，简化调用
    """
    try:
        config = read_sc_web_config()
        # 检查server_uploads节是否存在
        if "server_uploads" not in config.sections():
            raise Exception("配置文件中未找到[server_uploads]节")

        # 提取配置项并返回字典
        uploads_config = {
            "user_local_path": config.get("server_uploads", "user_local_path", fallback=""),
            "web_path": config.get("server_uploads", "web_path", fallback="")
        }
        return uploads_config
    except Exception as e:
        raise Exception(f"获取server_uploads配置失败: {str(e)}")




system_path = platform.system()

def read_logs(
    file_path: str,
    level: Optional[str] = None,
    keyword: Optional[str] = None,
    limit: int = 100
) -> list:
    """读取日志文件，支持筛选和限制条数"""
    if not os.path.exists(file_path):
        return ["日志文件不存在"]

    # 按行读取日志（从末尾开始，取最新的）
    logs = []
    with open(file_path, "r", encoding="utf-8") as f:
        # 先获取所有行，再反转（最新的在前）
        all_lines = f.readlines()
        reversed_lines = reversed(all_lines)  # 从最后一行开始读

        for line in reversed_lines:
            line = line.strip()
            if not line:
                continue

            # 按级别筛选
            if level and f" - {level.upper()} - " not in line:
                continue

            # 按关键词筛选
            if keyword and keyword not in line:
                continue

            logs.append(line)
            if len(logs) >= limit:
                break

    # 再次反转，让最新的在最后（符合阅读习惯）
    return logs[::-1]

# 4. 日志查询接口
@app.get("/api/logs")
async def get_logs(
    level: Optional[str] = Query(None, description="日志级别（INFO/WARNING/ERROR）"),
    keyword: Optional[str] = Query(None, description="日志关键词"),
    limit: int = Query(100, ge=1, le=1000, description="返回日志条数（1-1000）")
):
    logs = read_logs(log_file_path, level, keyword, limit)
    return {
        "log_file": log_file_path,
        "total": len(logs),
        "limit": limit,
        "logs": logs
    }

# 简化的系统判断
if system_path == "Windows":
    # 基础路径配置
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    # # 静态目录的本地绝对路径（与mount的directory保持一致）
    UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
    # # 静态文件的Web访问前缀（与mount的第一个参数保持一致）
    STATIC_WEB_PREFIX = "/uploads"
    WEB_File_Path = False
    try:
        os.makedirs(UPLOAD_DIR, exist_ok=True)
    except PermissionError:
        UPLOAD_DIR = os.path.join(os.gettempdir(), "docx_uploads")
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        print(f"警告：无法在当前目录创建uploads文件夹，已切换到系统临时目录：{UPLOAD_DIR}")
    app.mount(STATIC_WEB_PREFIX, StaticFiles(directory=UPLOAD_DIR), name="uploads")
    pass
else:
    uploads_config = get_server_uploads_config()
    UPLOAD_DIR = uploads_config["user_local_path"]
    STATIC_WEB_PREFIX = uploads_config["web_path"]
    WEB_File_Path = True


# 确保目录存在（增加权限检查）


# PostgreSQL数据库配置（从 conf/sc_web.conf 读取）
_pg = read_sc_web_config()["postgres"]
POSTGRES_CONFIG = {
    "host":     _pg.get("host"),
    "port":     int(_pg.get("port")),
    "user":     _pg.get("user"),
    "password": _pg.get("password"),
    "database": _pg.get("database"),
    "options":  _pg.get("options"),
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



class UnescapedJSONResponse(JSONResponse):
    def render(self, content: any) -> bytes:
        # ensure_ascii=False: 保留非 ASCII 字符（如中文）
        # separators: 优化 JSON 格式，可选
        # default: 处理其他特殊类型的默认序列化逻辑
        return json.dumps(
            content,
            ensure_ascii=False,
            allow_nan=False,
            indent=None,
            separators=(",", ":"),
            default=str
        ).encode("utf-8")

# ====================== 统一返回格式工具函数 ======================
async def download_image(url: str, save_dir: str = UPLOAD_DIR) -> str:
    """
    异步下载网络图片到本地临时目录（支持空URL）
    :param url: 图片网络地址（允许空字符串/None）
    :param save_dir: 保存目录
    :return: 本地文件路径（空URL返回空字符串）
    """
    # 处理空URL
    if not url or url.strip() == "":
        return ""

    try:
        # 生成唯一文件名
        file_ext = url.split(".")[-1].split("?")[0]  # 处理URL带参数的情况
        if len(file_ext) > 5:  # 如果扩展名过长，使用默认
            file_ext = "jpg"
        file_name = f"{uuid.uuid4()}.{file_ext}"
        file_path = os.path.join(save_dir, file_name)

        # 异步下载图片
        async with aiohttp.ClientSession() as session:
            async with session.get(url, timeout=aiohttp.ClientTimeout(total=30)) as response:
                if response.status != 200:
                    print(f"⚠️ 图片下载失败：{url}，状态码：{response.status}，返回空路径")
                    return ""

                # 保存图片到本地
                with open(file_path, "wb") as f:
                    f.write(await response.read())

        # 验证图片文件有效性
        try:
            with Image.open(file_path) as img:
                img.verify()  # 验证图片完整性
        except Exception as e:
            os.remove(file_path)
            print(f"⚠️ 图片文件无效：{url}，错误：{str(e)}，返回空路径")
            return ""

        return file_path

    except Exception as e:
        print(f"⚠️ 下载图片失败：{url}，错误：{str(e)}，返回空路径")
        return ""


async def download_images(urls: List[str]) -> List[str]:
    """
    批量下载图片（支持空URL）
    :param urls: 图片URL列表（允许包含空字符串/None）
    :return: 本地文件路径列表（空URL对应空字符串）
    """
    if not urls:
        return []

    # 并发下载所有图片（空URL会被跳过）
    tasks = [download_image(url) for url in urls]
    local_paths = await asyncio.gather(*tasks)

    return local_paths


# 工具函数：生成文档并转换为HTML
def generate_and_convert_to_html(generate_func, *args, **kwargs):
    """
    通用函数：生成DOCX文档并转换为HTML（修复文件占用问题）
    :param generate_func: 文档生成函数
    :return: HTML内容
    """
    # 1. 创建唯一的临时目录（避免文件冲突）
    temp_dir = tempfile.mkdtemp(prefix="docx_html_temp_")
    docx_path = os.path.join(temp_dir, f"temp_doc_{uuid.uuid4().hex[:8]}.docx")
    html_path = os.path.join(temp_dir, f"temp_html_{uuid.uuid4().hex[:8]}.html")

    try:
        # 2. 生成DOCX文档（使用普通文件，而非临时文件句柄）
        generate_func(*args, save_path=docx_path, **kwargs)

        # 3. 强制释放文件句柄（关键修复）
        import gc
        gc.collect()  # 强制垃圾回收，释放docx库的文件句柄
        time.sleep(0.1)  # 短暂等待，确保系统释放文件锁

        # 4. 转换为HTML（使用文件路径，而非句柄）
        converter = DocxHtmlConverter()
        html_content = converter.docx_to_single_html(docx_path, html_path)

        # # 5. 读取HTML内容（确保内容完整）
        # with open(html_path, 'r', encoding='utf-8') as f:
        #     final_html = f.read()

        return html_content

    finally:
        # 6. 强制清理所有临时文件（延迟清理，确保文件锁释放）
        def clean_temp_files():
            try:
                # 再次强制垃圾回收
                gc.collect()
                time.sleep(0.2)

                # 删除文件
                if os.path.exists(docx_path):
                    os.remove(docx_path)
                if os.path.exists(html_path):
                    os.remove(html_path)
                # 删除目录
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception as e:
                print(f"⚠️ 清理临时文件警告：{e}")

        # 使用线程延迟清理（避免当前进程占用）
        import threading
        cleanup_thread = threading.Thread(target=clean_temp_files)
        cleanup_thread.daemon = True
        cleanup_thread.start()


def local_upload_path_to_web_path(local_abs_path: str, request: Request) -> str:
    """
    将uploads本地绝对路径转换为Web路径，并自动生成完整URL

    Args:
        local_abs_path: 本地绝对路径
        request: FastAPI的Request对象（用于动态获取域名/端口）

    Returns:
        包含web_path和full_url的字典
    """
    if WEB_File_Path:
        # return STATIC_WEB_PREFIX+os.path.basename(local_abs_path)
        return UPLOAD_DIR + "/" +os.path.basename(local_abs_path)

    local_abs_path = os.path.normpath(local_abs_path)
    uploads_local_dir = os.path.normpath(UPLOAD_DIR)

    # 检查路径合法性
    if not local_abs_path.startswith(uploads_local_dir):
        raise ValueError(f"路径 {local_abs_path} 不在uploads目录下")

    # 生成Web路径
    relative_path = local_abs_path[len(uploads_local_dir):]

    # 从Request对象动态获取：协议(http/https) + 域名/IP + 端口 + Web路径
    full_url = request.url_for("uploads", path=relative_path.lstrip(os.sep))

    return str(full_url)  # 转为字符串，方便使用



def is_web_path_(path_str):
    """
    判断路径是否为网页URL路径
    :param path_str: 待判断的路径字符串
    :return: True（是网页路径）/ False（不是）
    """
    # 去除首尾空格，避免干扰
    path_str = path_str.strip()
    # 解析URL，判断是否有有效的协议头
    parsed = urlparse(path_str)
    # 常见的网页协议头
    web_schemes = {'http', 'https', 'ftp', 'ftps'}
    return parsed.scheme in web_schemes

def is_local_path_(path_str):
    """
    判断路径是否为本地文件路径
    :param path_str: 待判断的路径字符串
    :return: True（是本地路径）/ False（不是）
    """
    path_str = path_str.strip()
    # 如果是网页路径，直接返回False
    if is_web_path_(path_str):
        return False

    # 判断是否符合本地路径特征
    # 1. Windows路径特征：包含盘符（如C:\）或反斜杠
    if os.name == 'nt':  # Windows系统
        # 匹配盘符格式（如C:、D:），或包含反斜杠，或是相对路径
        has_drive = len(path_str) >= 2 and path_str[1] == ':' and path_str[0].isalpha()
        has_backslash = '\\' in path_str
        return has_drive or has_backslash or os.path.exists(path_str)
    else:  # Linux/macOS系统
        # 以/开头（绝对路径），或存在相对路径，或文件实际存在
        is_abs = path_str.startswith('/')
        return is_abs or os.path.exists(path_str)


def judge_path_type(path_str):
    """
    综合判断路径类型，返回类型描述
    :param path_str: 待判断的路径字符串
    :return: 字符串（web/local/unknown）
    """
    if not path_str:
        return 'unknown'

    if is_web_path_(path_str):
        return 'web'
    elif is_local_path_(path_str):
        return 'local'
    else:
        return 'unknown'


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


def is_ends_with_path_separator(s: str) -> bool:
    """
    判断字符串是否以路径分隔符（/、\\、//、\\\\等）结尾

    Args:
        s: 要检查的字符串

    Returns:
        bool: 如果以路径分隔符结尾返回True，否则返回False
    """
    # 处理空字符串的特殊情况
    if not s:
        return False

    # 正则表达式解释：
    # $ 表示匹配字符串结尾
    # [/\\]+ 表示匹配一个或多个 / 或 \（\需要双重转义）
    pattern = r'[/\\]+$'

    # 使用re.search检查是否匹配
    match = re.search(pattern, s)

    return match is not None

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


def docx_to_html(file_path: str):
    """Word转HTML的实现函数"""

    try:
        # abs_file_path = os.path.abspath("temp.docx")
        # if not os.path.exists(abs_file_path):
        #     return f"<p>转换失败：文件不存在: {abs_file_path}</p>"
        if judge_path_type(file_path) == 'web':
            # new_file_path = convert_path1_to_path2(file_path)
            new_file_path = file_path
            response = requests.get(new_file_path, timeout=30)
            # 校验响应状态码（200表示成功）
            response.raise_for_status()
            temp_docx_filename = generate_unique_filename("temp.docx")
            abs_file_path = os.path.join(UPLOAD_DIR, temp_docx_filename)
            # abs_file_path = os.path.abspath("temp.docx")

            # 将文件内容写入本地
            with open(abs_file_path, 'wb') as f:
                f.write(response.content)

            # print(f"文件下载成功，本地路径：{save_path}")
        else:
            abs_file_path = file_path

        # 文件大小检查
        file_size = os.path.getsize(abs_file_path)
        if file_size > 10 * 1024 * 1024:
            print(f"警告：文件过大（{file_size / 1024 / 1024:.2f}MB），可能转换失败")

        converter = DocxHtmlConverter()
        temp_html_filename = generate_unique_filename("temp.html")
        temp_html_path = os.path.join(UPLOAD_DIR, temp_html_filename)

        # 执行DOCX转HTML
        html_content = converter.docx_to_single_html(abs_file_path, temp_html_path)

        # 读取并返回HTML内容
        # html_content = ""
        # print(html_content)
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
                    pass
                except Exception as e:
                    print(f"警告：无法删除临时文件 {temp_html_path} - {e}")
        result_html = html_content or ""
        return result_html, abs_file_path
    except Exception as e:
        print(f"Word转HTML失败: {str(e)}")
        return f"<p>转换失败：{str(e)}</p>"


def convert_html_to_docx(html_content: str) -> Tuple[bool, Union[io.BytesIO, str], str]:
    """HTML转DOCX的实现函数"""
    try:
        if not html_content.strip():
            return False, "HTML内容不能为空"

        converter = DocxHtmlConverter()
        temp_docx_filename = generate_unique_filename("html2docx.docx")
        temp_docx_path = os.path.join(UPLOAD_DIR, temp_docx_filename)

        converter.html_text_to_docx(html_content, temp_docx_path)

        if not os.path.exists(temp_docx_path):
            return False, f"转换失败：未生成文件 {temp_docx_path}", temp_docx_path

        # 读取文件到内存流
        docx_stream = io.BytesIO()
        with open(temp_docx_path, 'rb') as f:
            docx_stream.write(f.read())
        docx_stream.seek(0)


        # # 删除临时文件
        # try:
        #     if os.path.exists(temp_docx_path):
        #         os.remove(temp_docx_path)
        # except Exception as e:
        #     print(f"警告：无法删除临时文件 {temp_docx_path} - {e}")

        return True, docx_stream, temp_docx_path
    except PermissionError:
        return False, "权限错误：无法创建/读取临时文件", ""
    except Exception as e:
        return False, f"HTML转DOCX失败：{str(e)}", ""


def get_html_heading_levels(html_content):
    """
    判断HTML中包含的标题层级

    参数:
        html_content: str - 待解析的HTML文本

    返回:
        dict - 包含两个键：
            - 'existing_levels': list - 实际存在的标题层级（如[1,2,4]）
            - 'max_level': int - 最大的标题层级（如4）
    """
    # 处理空输入
    if not html_content or not isinstance(html_content, str):
        return {'existing_levels': [], 'max_level': 0}

    # 解析HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # 查找所有h1-h6标签
    headings = soup.find_all(re.compile(r'^h[1-6]$', re.IGNORECASE))

    # 提取标题层级
    existing_levels = []
    for heading in headings:
        # 提取h标签后的数字（如h3提取3）
        level = int(heading.name.lower().replace('h', ''))
        if level not in existing_levels:
            existing_levels.append(level)

    # 排序并计算最大层级
    existing_levels.sort()
    max_level = max(existing_levels) if existing_levels else 0

    return existing_levels, max_level


def limit_html_heading_levels(html_content, max_allowed_level):
    """
    将HTML中超过指定层级的标题替换为指定的最大层级；层级为0时移除标题标签，仅保留内容

    参数:
        html_content: str - 待处理的HTML文本
        max_allowed_level: int - 允许的最大标题层级（0-6）：
                               0 = 移除标题标签，仅保留内容（不丢失任何文本）
                               1-6 = 替换为对应层级标题

    返回:
        str - 处理后的HTML文本
    """
    # 校验输入参数
    if not isinstance(max_allowed_level, int) or max_allowed_level < 0 or max_allowed_level > 6:
        raise ValueError("max_allowed_level必须是0-6之间的整数")
    if not html_content or not isinstance(html_content, str):
        return html_content

    # 解析HTML（使用html5lib解析器，更好地保留结构）
    soup = BeautifulSoup(html_content, 'html5lib')

    # 查找所有h1-h6标签
    headings = soup.find_all(re.compile(r'^h[1-6]$', re.IGNORECASE))

    # 替换标题标签
    for heading in headings:
        current_level = int(heading.name.lower().replace('h', ''))

        # 层级为0：移除标题标签，仅保留其内部所有内容（不丢失任何文本/子标签）
        if max_allowed_level == 0:
            # 提取标题内的所有子节点（文本+标签）
            contents = heading.contents
            # 将标题标签替换为其内部内容（直接插入，不包裹任何标签）
            heading.replace_with(*contents)

        # 层级1-6：替换超过层级的标题
        elif current_level > max_allowed_level:
            # 创建新的标题标签
            new_heading = soup.new_tag(f'h{max_allowed_level}')
            # 保留原标签的所有内容和属性
            new_heading.contents = heading.contents
            new_heading.attrs = heading.attrs
            # 替换原标签
            heading.replace_with(new_heading)

    # 返回处理后的HTML文本（格式化输出，更易读）
    return soup.prettify()


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
      "node_type" varchar(16) COLLATE "pg_catalog"."default" DEFAULT 'main' COMMENT '节点类型：main/branch',
      "split_id" int4 COMMENT '拆分接口返回的树节点id，用于合并时还原树结构'
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
    COMMENT ON COLUMN "yxdl_docx_title_trees"."split_id" IS '拆分接口返回的树节点id，用于合并时还原树结构';
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
def download_image_to_base64(image_url, base_url=None, timeout=10):
    """
    下载图片到临时文件 → 转换为Base64 → 立即删除临时文件
    """
    temp_file = None
    temp_file_path = None
    try:
        # 清理URL中的多余字符
        image_url = image_url.strip().split()[0]
        if image_url.startswith(('"', "'")) and image_url.endswith(('"', "'")):
            image_url = image_url[1:-1]

        # 处理 data: URI（已是base64编码，无需下载）
        if image_url.startswith('data:'):
            try:
                header, data_part = image_url.split(',', 1)
                meta = header[5:]  # 去掉 "data:"
                parts = meta.split(';')
                content_type = parts[0] if parts[0] else 'image/jpeg'
                if 'base64' in parts:
                    return data_part, content_type
                else:
                    import urllib.parse
                    decoded = urllib.parse.unquote_to_bytes(data_part)
                    return base64.b64encode(decoded).decode('utf-8'), content_type
            except Exception as e:
                print(f"解析 data: URI 失败: {str(e)}")
                return None, None

        # 处理相对路径
        if base_url and not image_url.startswith(('http://', 'https://')):
            image_url = f"{base_url.rstrip('/')}/{image_url.lstrip('/')}"

        # 模拟浏览器请求头
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }

        # 下载图片（启用SSL验证）
        response = requests.get(
            image_url,
            headers=headers,
            timeout=timeout,
            stream=True,
            verify=True
        )
        response.raise_for_status()

        # 获取图片MIME类型
        content_type = response.headers.get('Content-Type', 'image/jpeg')

        # 创建临时文件并关闭句柄（避免Windows占用）
        suffix = f".{content_type.split('/')[-1]}" if '/' in content_type else '.jpg'
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        temp_file_path = temp_file.name
        temp_file.close()

        # 写入临时文件
        with open(temp_file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        # 转换为Base64
        with open(temp_file_path, 'rb') as f:
            image_data = f.read()
            base64_encoded = base64.b64encode(image_data).decode('utf-8')

        return base64_encoded, content_type

    except Exception as e:
        print(f"下载/转换图片失败 {image_url}: {str(e)}")
        return None, None

    finally:
        # 清理临时文件（增加重试机制）
        if temp_file_path and os.path.exists(temp_file_path):
            retry = 3
            while retry > 0:
                try:
                    os.unlink(temp_file_path)
                    print(f"临时文件已清理：{temp_file_path}")
                    break
                except Exception as e:
                    retry -= 1
                    if retry == 0:
                        print(f"清理临时文件失败 {temp_file_path}: {str(e)}")
                    else:
                        import time
                        time.sleep(0.1)


def html_img_url_to_base64(html_text, base_url=None, timeout=10):
    """
    一对一精准替换：每个img标签独立处理，确保所有匹配项都被替换
    """
    temp_dir = tempfile.mkdtemp(prefix="img_base64_re_")
    try:
        # 步骤1：匹配所有完整的img标签（保留原始内容）
        # 正则：匹配完整的<img ...>标签
        full_img_pattern = re.compile(r'<img[^>]+>', re.IGNORECASE | re.DOTALL)

        # 步骤2：提取所有img标签的列表（有序，保证一对一）
        img_tags = full_img_pattern.findall(html_text)
        if not img_tags:
            print("未找到任何img标签，直接返回原HTML")
            return html_text, {"success": 0, "fail": 0}

        # 步骤3：为每个img标签生成替换后的版本
        replacement_map = {}  # 原始标签 → 替换后的标签
        success_count = 0
        fail_count = 0

        for original_img_tag in img_tags:
            # 跳过已处理的相同标签（但保留顺序）
            if original_img_tag in replacement_map:
                continue

            # 提取当前标签的src值
            src_pattern = re.compile(r'src\s*=\s*(?:"([^"]+)"|\'([^\']+)\'|([^\s>]+))', re.IGNORECASE)
            src_match = src_pattern.search(original_img_tag)

            if not src_match:
                replacement_map[original_img_tag] = original_img_tag
                fail_count += 1
                print(f"跳过无src的img标签：{original_img_tag[:50]}...")
                continue

            # 获取纯净的src值
            src_values = src_match.groups()
            img_url = next((v for v in src_values if v is not None), "").strip()

            if not img_url:
                replacement_map[original_img_tag] = original_img_tag
                fail_count += 1
                print(f"跳过空src的img标签：{original_img_tag[:50]}...")
                continue

            # 下载并转换为Base64
            base64_str, content_type = download_image_to_base64(img_url, base_url, timeout)

            if base64_str and content_type:
                # 替换当前标签的src属性（只替换当前标签内的src）
                new_img_tag = src_pattern.sub(
                    f'src="data:{content_type};base64,{base64_str}"',
                    original_img_tag,
                    count=1  # 只替换当前标签内的第一个src
                )
                replacement_map[original_img_tag] = new_img_tag
                success_count += 1
                print(f"成功替换图片：{img_url} → 标签已更新")
            else:
                replacement_map[original_img_tag] = original_img_tag
                fail_count += 1
                print(f"替换失败：{img_url} → 保留原标签")

        # 步骤4：按顺序替换HTML中的所有img标签（一对一）
        processed_html = html_text
        for original_img_tag in img_tags:
            # 用replace替换当前标签（保证顺序和数量一致）
            processed_html = processed_html.replace(
                original_img_tag,
                replacement_map[original_img_tag],
                1  # 每次只替换一个，避免批量替换错位
            )

        # 统计信息
        stats = {
            "success": success_count,
            "fail": fail_count,
            "total": len(img_tags)
        }
        return processed_html, stats

    finally:
        # 清理临时目录
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                print(f"\n临时目录已清理：{temp_dir}")
            except Exception as e:
                print(f"\n清理临时目录失败 {temp_dir}: {str(e)}")

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
    if node.eid in eid_path_map:
        node.file_path = eid_path_map[node.eid]
    else:
        logger.warning(f"assign_file_path_to_tree: eid={node.eid!r} 未在 files 中找到匹配，可用 eid={list(eid_path_map.keys())}")
    if node.children:
        for child in node.children:
            assign_file_path_to_tree(child, eid_path_map)


def get_next_batch_count(record_id: int) -> int:
    """查询 record_id 下已有的最大 batch_count，返回 +1 后的值（首次返回 1）"""
    sql = """
        SELECT COALESCE(MAX(batch_count), 0) + 1
        FROM "yxdl_docx_title_trees"
        WHERE record_id = %s;
    """
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, (record_id,))
            return cursor.fetchone()[0]


def process_split_tree_nodes(
        nodes: List[TreeItem],
        record_id: int,
        current_time: datetime.datetime,
        file_base_path: str,
        convert_html: bool = True,
        parent_id: Optional[int] = None,
        batch_count: int = 1,
) -> List[Dict[str, Any]]:
    """
    递归处理拆分后的树节点，入库，返回带层级结构的节点信息

    Args:
        nodes: 树节点列表
        record_id: 记录ID
        current_time: 当前时间
        file_base_path: 文件基础路径
        convert_html: 是否立即将 DOCX 转换为 HTML 写入数据库；
                      False 时跳过转换，html_content 存空字符串，
                      is_conversion_completion=0，由 get_html_by_node 懒转换。
        parent_id: 父节点数据库ID，NULL表示根节点
        batch_count: 本次导入批次号，同一次拆分共享同一值，用于排序（大值排前）

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
        return []

    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

    result_nodes = []

    for node in nodes:
        if not isinstance(node, TreeItem):
            logger.warning(f"process_split_tree_nodes: 跳过非TreeItem节点 type={type(node)}")
            continue

        # 1. 生成节点标题
        node_title = (
            node.text.strip()
            if (node.text and isinstance(node.text, str))
            else f"节点_{node.eid or '未知'}_{node.level}_{node.idx}"
        )
        node_file_path = node.file_path or ""
        level = node.level

        # 2. 按开关决定是否立即转换 DOCX → HTML
        if convert_html and node_file_path:
            try:
                html_content, _ = docx_to_html(node_file_path)
            except Exception as e:
                logger.error(f"process_split_tree_nodes: docx_to_html 失败 path={node_file_path} err={e}")
                html_content = ""
        else:
            if convert_html and not node_file_path:
                logger.warning(f"process_split_tree_nodes: 节点 eid={node.eid} file_path 为空，跳过转换")
            html_content = ""
        is_conversion_completion = 1 if html_content else 0

        # 3. 插入数据库
        insert_tree_sql = """
        INSERT INTO "yxdl_docx_title_trees"
        (record_id, title_text, html_content, create_time, update_time,
         level, eid, idx, node_type, origin_file_path, is_conversion_completion, parent_id, batch_count, split_id)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        RETURNING id;
        """
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(insert_tree_sql, (
                    record_id, node_title, html_content,
                    current_time, current_time,
                    node.level, node.eid, node.idx,
                    "branch", node_file_path, is_conversion_completion,
                    parent_id, batch_count, node.id
                ))
                node_id = cursor.fetchone()[0]
                conn.commit()
        logger.info(f"process_split_tree_nodes: 插入节点 id={node_id} eid={node.eid} level={level}")

        # 4. 构建返回字典
        current_node = {
            "name": node_title,
            "node_id": node_id,
            "level": level,
            "file_name": node_file_path,
            "children": []
        }

        # 5. 递归处理子节点，透传 convert_html / parent_id / batch_count
        children = node.children or []
        if children:
            current_node["children"] = process_split_tree_nodes(
                children, record_id, current_time, file_base_path,
                convert_html=convert_html,
                parent_id=node_id,
                batch_count=batch_count,
            )

        result_nodes.append(current_node)

    return result_nodes


def _tree_item_to_dict(node: Dict[str, Any], file_name: str) -> Dict[str, Any]:
    """
    辅助函数：将节点数据转换为目标字典格式
    """
    # 构建基础字段
    node_dict = {
        "eid": node.get("eid") or "",
        "text": node.get("title_text") or "",
        "level": node.get("level", 0),
        "id": node.get("id", -1),  # 数据库id，无则默认-1
        "idx": node.get("idx", 0),
        "file_info": {
            "is_had_title": 1  # 固定值，按示例要求
        },
        "file_name": file_name,  # 文件名（从路径提取或传入）
        "children": []  # 初始化子节点
    }
    return node_dict


def _build_nested_dict(nodes: List[Dict[str, Any]], file_name: str) -> List[Dict[str, Any]]:
    """
    递归构建嵌套的字典结构（核心函数）
    """
    result = []
    # 先按level分组
    level_map = {}
    for node in nodes:
        level = node.get("level", 0)
        if level not in level_map:
            level_map[level] = []
        level_map[level].append(node)

    # 排序层级
    sorted_levels = sorted(level_map.keys())
    if not sorted_levels:
        return result

    # 递归构建父子关系
    def build_children(parent_level: int, parent_node: Dict[str, Any] = None):
        """递归为父节点添加子节点"""
        # 找到下一级
        next_level = parent_level + 1
        if next_level not in level_map:
            return []

        # 筛选当前父节点的子节点（按idx顺序）
        child_nodes = level_map[next_level]
        child_result = []

        for child in child_nodes:
            # 转换为目标字典格式
            child_dict = _tree_item_to_dict(child, file_name)
            # 递归添加子节点的子节点
            child_dict["children"] = build_children(next_level, child)
            child_result.append(child_dict)

        return child_result

    # 处理根节点（最小层级）
    root_level = sorted_levels[0]
    root_nodes = level_map[root_level]

    for root_node in root_nodes:
        root_dict = _tree_item_to_dict(root_node, file_name)
        # 为根节点添加子节点
        root_dict["children"] = build_children(root_level, root_node)
        result.append(root_dict)

    return result


def recover_split_tree_nodes(record_id: int) -> List[Dict[str, Any]]:
    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

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
                cursor.execute(select_sql, (record_id,))
                node_records = cursor.fetchall()
    except Exception as e:
        raise RuntimeError(f"查询数据库失败：{str(e)}") from e

    if not node_records:
        return []

    # 先用 build_simplified_tree 组装嵌套结构（依赖 parent_id）
    nested = build_simplified_tree(node_records)

    # 递归将嵌套 dict 转为 TreeItem 兼容的格式（字段重命名）
    def _remap(nodes: List[Dict]) -> List[Dict]:
        result = []
        for node in nodes:
            # 每个节点独立选取自己的文件路径
            if node.get("is_conversion_completion") == 1 and node.get("update_file_path"):
                file_path = node["update_file_path"]
            else:
                file_path = node.get("origin_file_path") or ""

            file_name = os.path.splitext(os.path.basename(file_path))[0] if file_path else ""

            remapped = {
                "eid":       node.get("eid") or "",
                "text":      node.get("title_text") or "",   # ← 关键：title_text → text
                "level":     node.get("level", 0),
                "id":        node.get("split_id"),           # ← 使用拆分接口返回的原始id，用于合并时还原树
                "idx":       node.get("idx", 0),
                "parent_id": node.get("parent_id"),
                "file_name": file_name,                      # ← 关键：逐节点取，不共用
                "file_path": file_path,
                "file_info": {"is_had_title": 1},
                "update_file_path": node.get("update_file_path") or "",
                "node_type": node.get("node_type") or "",
                "is_conversion_completion": node.get("is_conversion_completion", 0),
                "children":  _remap(node.get("children", [])),
            }
            result.append(remapped)
        return result

    return _remap(nested)


def build_simplified_tree(rows):
    items = [dict(row) for row in rows]
    for item in items:
        item['children'] = []

    id_map = {item['id']: item for item in items}

    tree = []
    for item in sorted(items, key=lambda x: x['idx']):
        pid = item.get('parent_id')
        if pid and pid in id_map:
            id_map[pid]['children'].append(item)
        else:
            tree.append(item)

    # batch_count 降序（新批次排前）；同批次内按 idx 升序（文档顺序）
    def _sort_key(x):
        return (-(x.get('batch_count') or 0), x['idx'])

    def sort_children(node):
        node['children'].sort(key=_sort_key)
        for child in node['children']:
            sort_children(child)

    fake_root = {'children': tree}
    sort_children(fake_root)
    tree.sort(key=_sort_key)

    return tree


def get_tree_node_file_paths(record_id: int) -> List[str]:
    """保持不变，适配新的recover_split_tree_nodes函数"""
    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

    select_sql = """
    SELECT 
        CASE 
            WHEN is_conversion_completion = 1 AND update_file_path IS NOT NULL AND update_file_path != '' 
            THEN update_file_path
            ELSE origin_file_path
        END AS file_path
    FROM "yxdl_docx_title_trees" 
    WHERE record_id = %s;
    """

    file_paths = []
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(select_sql, (record_id,))
                raw_paths = [row[0] for row in cursor.fetchall()]
                file_paths = [path for path in raw_paths if path and isinstance(path, str) and path.strip() != ""]
    except Exception as e:
        raise RuntimeError(f"查询文件路径失败：{str(e)}") from e

    unique_file_paths = []
    seen = set()
    for path in file_paths:
        if path not in seen:
            seen.add(path)
            unique_file_paths.append(path)

    return unique_file_paths


def get_tree_node_file_paths(record_id: int) -> List[str]:
    """保持不变，仅适配修复后的recover_split_tree_nodes"""
    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

    select_sql = """
    SELECT
        CASE
            WHEN is_conversion_completion = 1 AND update_file_path IS NOT NULL AND update_file_path != ''
            THEN update_file_path
            ELSE origin_file_path
        END AS file_path
    FROM "yxdl_docx_title_trees"
    WHERE record_id = %s;
    """

    file_paths = []
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(select_sql, (record_id,))
                raw_paths = [row[0] for row in cursor.fetchall()]
                file_paths = [path for path in raw_paths if path and isinstance(path, str) and path.strip() != ""]
    except Exception as e:
        raise RuntimeError(f"查询文件路径失败：{str(e)}") from e

    unique_file_paths = []
    seen = set()
    for path in file_paths:
        if path not in seen:
            seen.add(path)
            unique_file_paths.append(path)

    return unique_file_paths


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
    html_content, temp_file_docx = docx_to_html(file_path)

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
        node_type,
        parent_id,
        batch_count
    )
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    RETURNING id;
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(insert_tree_sql, (
                record_id,
                DEFAULT_MAIN_NODE["title"],
                html_content,
                current_time,
                current_time,
                DEFAULT_MAIN_NODE["level"],
                DEFAULT_MAIN_NODE["eid"],
                DEFAULT_MAIN_NODE["idx"],
                "main",
                None,  # 根节点，parent_id=NULL
                1,     # 首次导入 batch_count=1
            ))
            node_id = cursor.fetchone()[0]
            conn.commit()

    print(f"成功创建主节点：ID={node_id}, 标题={DEFAULT_MAIN_NODE['title']}")
    return node_id

def deduplicate_dict_list(dict_list):
    """
    对包含字典的列表进行去重

    Args:
        dict_list: 包含字典的列表

    Returns:
        去重后的新列表，保留第一次出现的字典
    """
    seen = set()  # 用于记录已经出现过的字典特征
    result = []  # 存储去重后的结果

    for d in dict_list:
        # 将字典转换为可哈希的元组（排序后），确保键值对顺序不影响去重
        # sorted(d.items()) 保证 {'a':1, 'b':2} 和 {'b':2, 'a':1} 被视为同一个字典
        tuple_repr = tuple(sorted(d.items()))

        if tuple_repr not in seen:
            seen.add(tuple_repr)
            result.append(d)

    return result


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
                file_id=split_file_id,
                had_title=0,
                rm_outline_in_doc=1
            )

            tree_nodes = [TreeItem(**item) for item in split_result.data.get("tree", [])]

            # 2. 构建 eid-文件路径 映射
            eid_path_map = build_eid_path_mapping(split_result.data.get("files", []))

            # 3. 为每个树节点分配文件路径
            for node in tree_nodes:
                assign_file_path_to_tree(node, eid_path_map)
            print(tree_nodes)
            batch_count = get_next_batch_count(record_id)
            node_ids = process_split_tree_nodes(
                nodes=tree_nodes,
                record_id=record_id,
                current_time=current_time,
                file_base_path=abs_file_path,
                batch_count=batch_count,
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
            if file_path and os.path.exists(file_path):
                os.remove(file_path)
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


@app.post("/doc_editor/route_generate_tree", summary="文件路径生成标题树节点")
async def route_generate_tree(
        # 替换原有的file参数，新增文件来源参数
        file_source_type: str = Body("url", description="文件来源类型：url-从URL下载，static-从静态路径读取"),
        file_source: str = Body(..., description="文件来源：URL地址 或 服务器静态文件路径"),
        process_mode: ProcessMode = Body("split", description="处理模式：single-单个主节点，split-接口拆分多节点（默认）")
) -> JSONResponse:
    """
    获取文件（URL下载/静态路径读取）并生成标题树节点
    - file_source_type: url（从URL下载）/ static（从静态路径读取）
    - file_source: 对应类型的文件地址/路径
    - 默认模式（不传参数）：split，调用拆分接口生成多个分支节点
    - single模式：生成单个主节点
    """
    file_path = ""
    split_file_id = ""
    file_content = b""
    original_filename = ""

    try:
        # 1. 根据文件来源类型获取文件内容和文件名
        if file_source_type == "url":
            # 从URL下载文件
            async with aiohttp.ClientSession() as session:
                async with session.get(file_source) as response:
                    if response.status != 200:
                        return unified_response(
                            code=400,
                            message=f"下载文件失败，HTTP状态码：{response.status}",
                            data={}
                        )
                    # 获取文件名（从URL或响应头提取）
                    content_disposition = response.headers.get("Content-Disposition", "")
                    if "filename=" in content_disposition:
                        original_filename = content_disposition.split("filename=")[-1].strip('"\'')
                    else:
                        # 从URL路径提取文件名
                        original_filename = file_source.split("/")[-1]
                    # 读取文件内容
                    file_content = await response.read()

        elif file_source_type == "static":
            # 从服务器静态路径读取文件
            static_file_path = os.path.abspath(file_source)
            if not os.path.exists(static_file_path):
                return unified_response(
                    code=400,
                    message=f"静态文件不存在：{static_file_path}",
                    data={}
                )
            if not static_file_path.lower().endswith(".docx"):
                ext = static_file_path.split('.')[-1] if '.' in static_file_path else '无后缀'
                return unified_response(
                    code=400,
                    message=f"仅支持docx格式文件，当前文件格式：{ext}",
                    data={}
                )
            # 读取文件内容
            with open(static_file_path, "rb") as f:
                file_content = f.read()
            # 获取文件名
            original_filename = os.path.basename(static_file_path)

        else:
            return unified_response(
                code=400,
                message=f"不支持的文件来源类型：{file_source_type}，仅支持url/static",
                data={}
            )

        # 2. 校验文件格式（补充URL下载文件的格式校验）
        if not original_filename.lower().endswith(".docx"):
            ext = original_filename.split('.')[-1] if '.' in original_filename else '无后缀'
            return unified_response(
                code=400,
                message=f"仅支持docx格式文件，当前文件格式：{ext}",
                data={}
            )

        # 3. 保存文件（和原有逻辑一致）
        new_filename = generate_unique_filename(original_filename)
        file_path = os.path.join(UPLOAD_DIR, new_filename)
        abs_file_path = os.path.abspath(file_path)

        with open(abs_file_path, "wb") as f:
            f.write(file_content)

        # 4. 生成文件记录（和原有逻辑一致）
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
                message="文件获取成功，生成单个主节点",
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
                file_id=split_file_id,
                had_title=0,
                rm_outline_in_doc=1
            )


            tree_nodes = [TreeItem(**item) for item in split_result.data.get("tree", [])]

            # 2. 构建 eid-文件路径 映射
            eid_path_map = build_eid_path_mapping(split_result.data.get("files", []))

            # 3. 为每个树节点分配文件路径
            for node in tree_nodes:
                assign_file_path_to_tree(node, eid_path_map)
            print(tree_nodes)
            batch_count = get_next_batch_count(record_id)
            node_ids = process_split_tree_nodes(
                nodes=tree_nodes,
                record_id=record_id,
                current_time=current_time,
                file_base_path=abs_file_path,
                batch_count=batch_count,
            )

            # 返回结果（保持原有结构）
            return unified_response(
                code=200,
                message=f"文件获取拆分成功，共生成{len(node_ids)}个分支节点",
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
            if file_path and os.path.exists(file_path):
                os.remove(file_path)
            # # 清理拆分接口的远端文件
            # if split_file_id and process_mode == "split":
            #     delete_request = DeleteRequest(id=split_file_id)
            #     call_docx_delete(delete_request)
        except Exception as cleanup_e:
            print(f"清理临时文件失败：{cleanup_e}")

        return unified_response(
            code=500,
            message=f"文件处理失败（模式：{process_mode}）：{str(e)}",
            data={"process_mode": process_mode, "split_file_id": split_file_id}
        )
def merge_html_texts(html_list: list[str]) -> str:
    """
    合并多个 HTML 文本，返回合并后的单个 HTML 字符串。

    参数:
        html_list: 待合并的 HTML 字符串列表
        strategy:  合并策略
            - "body"     : 提取每个文档的 <body> 内容，拼接后包裹成完整 HTML（默认）
            - "full"     : 保留第一个文档的结构，将其余文档的 <body> 追加进去
            - "concat"   : 直接拼接原始 HTML 字符串（最简单，不解析）

    返回:
        合并后的 HTML 字符串
    """

    merged_body_parts = []
    for html in html_list:
        soup = BeautifulSoup(html, "html.parser")
        body = soup.body
        if body:
            merged_body_parts.append(body.decode_contents())
        else:
            # 没有 <body> 标签时，取全部内容
            merged_body_parts.append(str(soup))

    merged_body = "\n".join(merged_body_parts)
    return f"<!DOCTYPE html>\n<html>\n<body>\n{merged_body}\n</body>\n</html>"

@app.post("/doc_editor/route_docx2html_marge", summary="文件路径docx转化html")
async def route_docx2html_marge(
        # 替换原有的file参数，新增文件来源参数
        file_source_type: str = Body("url", description="文件来源类型：url-从URL下载，static-从静态路径读取"),
        file_source: str = Body(..., description="文件来源：URL地址 或 服务器静态文件路径"),
) -> JSONResponse:
    """
    获取文件（URL下载/静态路径读取）并生成标题树节点
    - file_source_type: url（从URL下载）/ static（从静态路径读取）
    - file_source: 对应类型的文件地址/路径
    - 默认模式（不传参数）：split，调用拆分接口生成多个分支节点
    - single模式：生成单个主节点
    """
    file_path = ""
    split_file_id = ""
    file_content = b""
    original_filename = ""

    try:
        # 1. 根据文件来源类型获取文件内容和文件名
        if file_source_type == "url":
            # 从URL下载文件
            async with aiohttp.ClientSession() as session:
                async with session.get(file_source) as response:
                    if response.status != 200:
                        return unified_response(
                            code=400,
                            message=f"下载文件失败，HTTP状态码：{response.status}",
                            data={}
                        )
                    # 获取文件名（从URL或响应头提取）
                    content_disposition = response.headers.get("Content-Disposition", "")
                    if "filename=" in content_disposition:
                        original_filename = content_disposition.split("filename=")[-1].strip('"\'')
                    else:
                        # 从URL路径提取文件名
                        original_filename = file_source.split("/")[-1]
                    # 读取文件内容
                    file_content = await response.read()

        elif file_source_type == "static":
            # 从服务器静态路径读取文件
            static_file_path = os.path.abspath(file_source)
            if not os.path.exists(static_file_path):
                return unified_response(
                    code=400,
                    message=f"静态文件不存在：{static_file_path}",
                    data={}
                )
            if not static_file_path.lower().endswith(".docx"):
                ext = static_file_path.split('.')[-1] if '.' in static_file_path else '无后缀'
                return unified_response(
                    code=400,
                    message=f"仅支持docx格式文件，当前文件格式：{ext}",
                    data={}
                )
            # 读取文件内容
            with open(static_file_path, "rb") as f:
                file_content = f.read()
            # 获取文件名
            original_filename = os.path.basename(static_file_path)

        else:
            return unified_response(
                code=400,
                message=f"不支持的文件来源类型：{file_source_type}，仅支持url/static",
                data={}
            )

        # 2. 校验文件格式（补充URL下载文件的格式校验）
        if not original_filename.lower().endswith(".docx"):
            ext = original_filename.split('.')[-1] if '.' in original_filename else '无后缀'
            return unified_response(
                code=400,
                message=f"仅支持docx格式文件，当前文件格式：{ext}",
                data={}
            )

        # 3. 保存文件（和原有逻辑一致）
        new_filename = generate_unique_filename(original_filename)
        file_path = os.path.join(UPLOAD_DIR, new_filename)
        abs_file_path = os.path.abspath(file_path)

        with open(abs_file_path, "wb") as f:
            f.write(file_content)

        # 4. 生成文件记录（和原有逻辑一致）
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
                    "split"  # process_mode（默认split）
                ))
                record_id = cursor.fetchone()[0]
                conn.commit()
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
            file_id=split_file_id,
            had_title=1,
            rm_outline_in_doc=1
        )
        # 2. 构建 eid-文件路径 映射
        files__ = split_result.data.get("files", [])
        logger.info(files__)
        html_list = []
        for file__ in files__:
            html_content, temp_file_docx_ = docx_to_html(file__)
            html_list.append(html_content)
        total_html_content = merge_html_texts(html_list)
        # html_content, temp_file_docx_ = docx_to_html(result["origin_file_path"])
        # 返回结果（保持原有结构）
        return unified_response(
            code=200,
            message=f"文件html转换成功",
            data={
                "html_content": total_html_content,
            }
        )

    except Exception as e:
        # 清理临时文件
        try:
            if file_path and os.path.exists(file_path):
                os.remove(file_path)
            # # 清理拆分接口的远端文件
            # if split_file_id and process_mode == "split":
            #     delete_request = DeleteRequest(id=split_file_id)
            #     call_docx_delete(delete_request)
        except Exception as cleanup_e:
            print(f"清理临时文件失败：{cleanup_e}")

        return unified_response(
            code=500,
            message=f"文件处理失败：{str(e)}",
            data={"split_file_id": split_file_id}
        )

@app.post("/doc_editor/split_uploads", summary="分片上传文件")
async def split_upload_and_generate_tree(
    request: Request,
    # 文件参数（File类型，带描述）
    file: UploadFile = File(..., description="需要上传的分片文件"),
    # 表单参数（Body类型，带描述和默认值）
    file_no: str = Body(..., description="分片编号"),
    file_sign: str = Body(..., description="文件唯一标识"),
    file_name: str = Body(..., description="原始文件名"),
    files_total_count: str = Body(..., description="分片总数"),
    # 可选：如果有处理模式参数，参考你的示例添加
) -> JSONResponse:
    """
    模板文件分片上传接口（FastAPI优化版）

    - 采用FastAPI原生参数定义方式，更直观
    - 保留原有分片上传核心逻辑
    - 支持参数描述和类型校验
    """

    try:
        # 读取配置
        split_file_path = UPLOAD_DIR

        # 拼接文件名（保持原有逻辑）
        full_file_name = f"{file_sign}_{file_name}"
        # new_filename = generate_unique_filename(file_name)
        file_path = UPLOAD_DIR
        abs_file_path = os.path.abspath(file_path)
        path_ = local_upload_path_to_web_path(abs_file_path, request)
        if not abs_file_path.endswith(os.path.sep):
            real_file_path = abs_file_path + os.path.sep + full_file_name
        else:
            real_file_path = abs_file_path + full_file_name

        # 读取上传文件内容
        file_content = await file.read()

        # 调用分片上传核心逻辑
        status, result, msg = file_resp.SplitUpload(
            split_file_path,
            file_no,
            full_file_name,
            files_total_count,
            file_sign,
            file_content
        )

        # 日志记录
        if status == 0:
            print(f"finish upload. sign:{file_sign} | file_no :{file_no} | total count:{files_total_count}")
        if status == 1:
            print(
                f"upload fail. error info:{msg} | sign:{file_sign} | file_no :{file_no} | total count:{files_total_count}")

        # 处理文件路径返回
        if result == 1:
            file_path = f"{path_}{full_file_name}"
            print("--------TemplateUploadFile-finish--------")
            print(f"file path : {file_path}")
        else:
            file_path = ''
            real_file_path = ''
        # 返回响应
        return JSONResponse(content={
            'status': status,
            'is_finish': result,
            'msg': msg,
            "data": {
                "file_path": file_path,
                "real_file_path": real_file_path
            }
        })

    except Exception as e:
        print("--------TemplateUploadFile-fail--------")
        print(f"TemplateUploadFile-失败：{str(e)}")
        return JSONResponse(
            status_code=500,
            content={'status': 1, 'is_finish': 0, 'msg': '接口异常', "data": ""}
        )


# ====================== 保留原有接口 ======================
@app.get("/doc_editor/get_html_by_node/{node_id}", summary="根据节点ID查询HTML文本")
async def get_html_by_node(request: Request,node_id: int) -> JSONResponse:
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
            html_content, temp_file_docx_ = docx_to_html(result["origin_file_path"])
            temp_file_docx = temp_file_docx_
            eid = os.path.splitext(os.path.basename(temp_file_docx_))[0]
            with get_db_connection() as conn:
                with conn.cursor(cursor_factory=RealDictCursor) as cursor:
                    update_sql = """UPDATE "yxdl_docx_title_trees"
SET html_content = %s, update_time = NOW(), is_conversion_completion = 1, update_file_path = %s, eid = %s
WHERE id = %s"""
                    cursor.execute(update_sql, (html_content, temp_file_docx, eid, node_id))
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
                    "html_content": updated_result["html_content"],
                    "temp_file_docx_": temp_file_docx_
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
async def update_html_by_node(request: Request,
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
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute("SELECT id FROM \"yxdl_docx_title_trees\" WHERE id = %s", (node_id,))
                if not cursor.fetchone():
                    return unified_response(
                        code=404,
                        message=f"未找到ID为{node_id}的标题树节点",
                        data={}
                    )
        logger.error(html_content)
        # html内部img转换base64
        html_content, status_ = html_img_url_to_base64(html_content)
        # html转换成docx
        success, result, temp_docx_path_1 = convert_html_to_docx(html_content)
        # 拼接sql
        temp_docx_path_ = temp_docx_path_1
        eid = os.path.splitext(os.path.basename(temp_docx_path_1))[0]
        update_fields = ["html_content = %s", "update_time = %s", "update_file_path = %s", "eid = %s"]
        update_values = [html_content, datetime.datetime.now(), temp_docx_path_, eid]
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
                cursor.execute(update_sql, tuple(update_values))
                conn.commit()

        return unified_response(
            code=200,
            message="节点HTML内容更新成功",
            data={
                "node_id": node_id,
                "updated_title": "标题更新为" + title_text.strip() if title_text else "标题未更新",
                "update_time": current_time.strftime("%Y-%m-%d %H:%M:%S")
            }
        )

    except Exception as e:
        return unified_response(
            code=500,
            message=f"更新节点HTML失败：{str(e)}",
            data={}
        )
def process_split_tree_nodes_with_select(
        tree_nodes_org: List[TreeItem],  # 入参改为tree_nodes_org
        record_id: int,
        current_time: datetime.datetime,
        file_base_path: str
) -> List[Dict[str, Any]]:
    """
    递归处理从数据库查询的树节点（tree_nodes_org），返回带层级结构的节点信息
    （已移除数据库插入操作）

    Args:
        tree_nodes_org: 从数据库查询的TreeItem列表
        record_id: 记录ID
        current_time: 当前时间
        file_base_path: 文件基础路径

    Returns:
        嵌套结构的节点列表，格式：
        [
            {
                "name": "节点标题",
                "node_id": 虚拟ID（基于数据库id/特征生成）,
                "level": 节点层级,
                "file_name": 文件路径,
                "children": [...]
            }
        ]
    """
    # 参数校验
    if not isinstance(tree_nodes_org, list) or not tree_nodes_org:
        return []

    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

    result_nodes = []

    for node in tree_nodes_org:
        if not isinstance(node, TreeItem):
            continue
        try:
            # 1. 生成节点标题（优先使用text，无则基于数据库字段生成）
            node_title = node.text.strip() if (node.text and isinstance(node.text, str)) else f"节点_{node.eid or '未知'}_{node.level}_{node.idx}"

            # 2. 提取文件路径相关字段（适配tree_nodes_org的字段）
            node_file_path = node.file_path or ""
            file_name = node_file_path
            level = node.level

            # 3. 生成虚拟node_id（优先使用数据库返回的id，无则用哈希）
            if node.id is not None:
                node_id = node.id  # 使用数据库查询到的原始ID
            else:
                # 备用：基于节点特征生成哈希ID
                unique_identifier = f"{record_id}_{node.level}_{node.idx}_{node.eid or 'no_eid'}"
                node_id = hash(unique_identifier) % (10 ** 8)

            # 4. 构建当前节点的返回字典
            current_node = {
                "name": node_title,
                "node_id": node_id,  # 优先使用数据库原始ID
                "level": level,
                "file_name": file_name,
                "update_file_path": node.update_file_path,  # 新增数据库返回的更新文件路径
                "is_conversion_completion": node.is_conversion_completion,  # 新增转换完成状态
                "children": []
            }

            # 5. 递归处理子节点并赋值给children
            if node.children:
                child_nodes = process_split_tree_nodes_with_select(
                    tree_nodes_org=node.children,
                    record_id=record_id,
                    current_time=current_time,
                    file_base_path=file_base_path
                )
                current_node["children"] = child_nodes

            # 6. 将当前节点添加到结果列表
            result_nodes.append(current_node)

        except ValueError as ve:
            continue
        except Exception as e:
            continue

    return result_nodes


def process_single_tree_node(
        node: TreeItem,
        record_id: int,
        id_: int,
        current_time: datetime.datetime,
        convert_html: bool = True,
) -> Dict[str, Any]:
    """
    处理单个树节点，更新数据库字段，返回节点信息

    Args:
        node: 单个树节点
        record_id: 记录ID
        id_: 数据库节点 id（WHERE 条件）
        current_time: 当前时间
        convert_html: 是否立即将 DOCX 转换为 HTML 写入数据库；
                      False 时跳过转换，html_content 不更新，
                      is_conversion_completion=0，由 get_html_by_node 懒转换。

    Returns:
        单个节点信息字典，格式：
        {
            "name": "节点标题",
            "node_id": 数据库ID,
            "level": 节点层级,
            "file_name": 文件名,
            "update_success": 是否更新成功
        }
    """
    result_node = {
        "name": "",
        "node_id": None,
        "level": node.level,
        "file_name": node.file_path,
        "update_success": False
    }

    if not isinstance(node, TreeItem):
        result_node["name"] = "无效节点类型"
        return result_node

    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

    try:
        node_title = (
            node.text.strip()
            if (node.text and isinstance(node.text, str))
            else f"节点_{node.eid or '未知'}_{node.level}_{node.idx}"
        )

        # 按开关决定是否立即转换 DOCX → HTML
        if convert_html and node.file_path:
            try:
                node_html_content, _ = docx_to_html(node.file_path)
            except Exception:
                node_html_content = ""
        else:
            node_html_content = ""
        is_conv_done = 1 if node_html_content else 0

        # html_content 仅在 convert_html=True 时写入，否则只更新文件路径等元数据
        if convert_html:
            update_tree_sql = """
            UPDATE "yxdl_docx_title_trees"
            SET update_time = %s, level = %s, origin_file_path = %s,
                html_content = %s, is_conversion_completion = %s, eid = %s
            WHERE record_id = %s AND id = %s
            RETURNING id;
            """
            params = (
                current_time, node.level, node.file_path,
                node_html_content, is_conv_done, node.eid,
                record_id, id_,
            )
        else:
            update_tree_sql = """
            UPDATE "yxdl_docx_title_trees"
            SET update_time = %s, level = %s, origin_file_path = %s,
                is_conversion_completion = %s, eid = %s
            WHERE record_id = %s AND id = %s
            RETURNING id;
            """
            params = (
                current_time, node.level, node.file_path,
                0, node.eid,
                record_id, id_,
            )

        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(update_tree_sql, params)
                update_result = cursor.fetchone()
                conn.commit()

                if update_result:
                    result_node["node_id"] = update_result[0]
                    result_node["update_success"] = True
                else:
                    logger.warning(
                        f"process_single_tree_node: 未匹配到节点 record_id={record_id} id={id_}"
                    )

        result_node["name"] = node_title

    except ValueError:
        pass
    except Exception as e:
        logger.error(f"process_single_tree_node: 失败 eid={node.eid} err={e}")

    return result_node


@app.post("/doc_editor/update_html_by_node_new", summary="更新节点HTML文本")
async def update_html_by_node_new(request: Request,
        node_id: int = Body(..., description="要更新的节点ID"),
        html_content: str = Body(..., description="更新后的HTML文本"),
        title_text: Optional[str] = Body(None, description="可选：更新节点标题文本")
) -> JSONResponse:
    """更新指定节点ID的HTML文本"""
    MAX_LEVEL_NODE = 9
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

        # ── 1. 查节点，先判空再解包 ──────────────────────────────────────
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    "SELECT id, record_id, level FROM \"yxdl_docx_title_trees\" WHERE id = %s",
                    (node_id,)
                )
                record_id_cursor = cursor.fetchone()
                if not record_id_cursor:
                    return unified_response(
                        code=404,
                        message=f"未找到ID为{node_id}的标题树节点",
                        data={}
                    )
                record_id = record_id_cursor[1]
                now_level  = record_id_cursor[2]

        logger.info(f"update_html_by_node_new: node_id={node_id} record_id={record_id} level={now_level}")

        # ── 2. 公共预处理 ────────────────────────────────────────────────
        html_content, status_ = html_img_url_to_base64(html_content)

        # with open("index.html", "w", encoding="utf-8") as f:
        #     f.write(html_content)
        #
        # print("HTML 文件已成功保存！")

        existing_levels, max_level = get_html_heading_levels(html_content)
        max_now_level = MAX_LEVEL_NODE - int(now_level)

        # ── 3. 公共辅助：查库 → 组装 TreeItem 列表 → build_simplified_tree ─
        def _query_and_build_tree(rec_id: int, cur_time: datetime.datetime) -> List[Dict[str, Any]]:
            """
            查询 record_id 下所有节点，组装成嵌套树结构后经
            process_split_tree_nodes_with_select 转换为标准返回格式。
            """
            select_sql = """
                SELECT
                    id, title_text, level, eid, idx, parent_id, batch_count,
                    origin_file_path, update_file_path, is_conversion_completion
                FROM "yxdl_docx_title_trees"
                WHERE record_id = %s
                ORDER BY level ASC, idx ASC;
            """
            try:
                with get_db_connection() as conn:
                    with conn.cursor(cursor_factory=RealDictCursor) as cursor:
                        cursor.execute(select_sql, (rec_id,))
                        node_records = cursor.fetchall()
            except Exception as e:
                raise RuntimeError(f"查询数据库失败：{str(e)}") from e

            # 字段映射：数据库行 → TreeItem
            tree_nodes_org = [
                TreeItem(**{
                    "id":                     item.get("id"),
                    "text":                   item.get("title_text"),
                    "level":                  item.get("level"),
                    "eid":                    item.get("eid"),
                    "idx":                    item.get("idx"),
                    "parent_id":              item.get("parent_id"),
                    "file_path":              item.get("origin_file_path"),
                    "update_file_path":       item.get("update_file_path", ""),
                    "is_conversion_completion": item.get("is_conversion_completion", 0),
                    "children":               [],
                    "file_name":              None,
                    "file_info":              None,
                    "node_type":              ""
                })
                for item in node_records
            ]

            # 用 build_simplified_tree 的逻辑将平铺列表组成嵌套结构，
            # 再映射回 TreeItem（children 字段已填充 dict，需转为 TreeItem）
            nested_dicts = build_simplified_tree(
                # build_simplified_tree 期望每行有 title_text/eid/level/idx，
                # 从 node_records 原始数据直接传入
                node_records
            )

            def _dicts_to_tree_items(nodes_dict: List[Dict]) -> List[TreeItem]:
                result = []
                for d in nodes_dict:
                    item = TreeItem(
                        eid=d.get("eid", ""),
                        level=d.get("level", 1),
                        idx=d.get("idx", 0),
                        text=d.get("text", "") or d.get("title_text", ""),
                        children=_dicts_to_tree_items(d.get("children", []))
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

            return process_split_tree_nodes_with_select(
                tree_nodes_org=nested_tree_items,
                record_id=rec_id,
                current_time=cur_time,
                file_base_path=""
            )

        # ── 4a. 无标题分支（max_level == 0）────────────────────────────
        if max_level == 0:
            success, result, temp_docx_path_1 = convert_html_to_docx(html_content)
            temp_docx_path_ = temp_docx_path_1
            eid = os.path.splitext(os.path.basename(temp_docx_path_1))[0]
            current_time = datetime.datetime.now()

            update_fields = ["html_content = %s", "update_time = %s",
                             "update_file_path = %s", "eid = %s",
                             "is_conversion_completion = %s"]
            update_values = [html_content, current_time, temp_docx_path_, eid, 1]

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
                    cursor.execute(update_sql, tuple(update_values))
                    conn.commit()

            node_ids = _query_and_build_tree(record_id, current_time)

            return unified_response(
                code=200,
                message="节点HTML内容更新成功",
                data={
                    "node_id":       node_id,
                    "node_ids":      node_ids,
                    "updated_title": ("标题更新为" + title_text.strip()) if title_text else "标题未更新",
                    "update_time":   current_time.strftime("%Y-%m-%d %H:%M:%S")
                }
            )

        # ── 4b. 有标题分支（max_level > 0），需拆分 ───────────────────────
        else:
            success, result, temp_docx_path_1 = convert_html_to_docx(html_content)
            if not success:
                return unified_response(code=500, message=f"HTML转DOCX失败：{result}", data={})

            temp_docx_path_ = temp_docx_path_1
            original_filename = os.path.abspath(temp_docx_path_)
            split_file_id = generate_unique_file_id()
            current_time = datetime.datetime.now()

            # 为本次拆分插入新的上传记录，拿到新 new_record_id
            insert_file_sql = """
                INSERT INTO "yxdl_docx_upload_records"
                (original_filename, new_filename, save_path,
                 upload_time, update_time, split_file_id, process_mode)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                RETURNING id;
            """
            with get_db_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(insert_file_sql, (
                        original_filename, original_filename, temp_docx_path_,
                        current_time, current_time, split_file_id, "split"
                    ))
                    new_record_id = cursor.fetchone()[0]
                    conn.commit()

            # result 是 BytesIO，拆分接口需要 bytes，seek(0) 后读出
            result.seek(0)
            file_bytes = result.read()

            # 调用拆分接口
            split_result = call_docx_split(
                file_stream=file_bytes,
                file_name=original_filename,
                file_id=str(node_id),
                had_title=0,
                rm_outline_in_doc=1
            )
            if split_result.status == 1:
                return unified_response(code=500, message=split_result.msg, data={})

            # 构建节点树并分配文件路径
            tree_nodes = [TreeItem(**item) for item in split_result.data.get("tree", [])]
            eid_path_map = build_eid_path_mapping(split_result.data.get("files", []))
            for node in tree_nodes:
                assign_file_path_to_tree(node, eid_path_map)

            if not tree_nodes:
                return unified_response(code=500, message="拆分结果为空", data={})

            batch_count = get_next_batch_count(record_id)

            # 首节点：更新原有节点，record_id 保持不变，取返回值拿到数据库 id
            first_node = tree_nodes.pop(0)
            first_result = process_single_tree_node(
                first_node, record_id, node_id, current_time,
                convert_html=False,
            )

            # 子节点统一用原 record_id 插入，parent_id 指向首节点，batch_count 标记本次批次
            remaining_nodes = (first_node.children or []) + tree_nodes
            process_split_tree_nodes(
                nodes=remaining_nodes,
                record_id=record_id,
                current_time=current_time,
                file_base_path=temp_docx_path_,
                convert_html=False,
                parent_id=first_result.get("node_id"),
                batch_count=batch_count,
            )

            # 用原 record_id 查询，组装最新嵌套树返回
            node_ids = _query_and_build_tree(record_id, current_time)

            return unified_response(
                code=200,
                message="更新成功",
                data={
                    "record_id":     record_id,
                    "node_ids":      node_ids,
                    "node_type":     "branch",
                    "split_file_id": split_file_id
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
        html_content: str = Body(..., description="需要转换的HTML文本"),
        filename: Optional[str] = Body("output.docx", description="下载的DOCX文件名（默认output.docx）"),
) -> Union[JSONResponse, StreamingResponse]:
    """接收HTML文本生成DOCX文件流"""
    try:
        if not html_content.strip():
            return unified_response(
                code=400,
                message="HTML内容不能为空",
                data={}
            )

        update_fields = ["html_content = %s", "update_time = %s"]
        update_values = [html_content, datetime.datetime.now()]
        current_time = update_values[1]


        # 校验文件名格式
        if not filename.lower().endswith(".docx"):
            filename += ".docx"
        filename = os.path.basename(filename).replace('/', '_').replace('\\', '_').replace(':', '_')

        # 调用HTML转DOCX函数
        success, result, path_ = convert_html_to_docx(html_content)
        # logging.error(success, result, path_)
        if not success:
            return unified_response(
                code=500,
                message=result,
                data={}
            )

        # 构造响应头
        encoded_filename = urllib.parse.quote(filename)

        # 2. 构造兼容的响应头（修复核心问题）
        headers = {
            # 使用RFC 5987标准格式，兼容所有浏览器且避免latin-1编码错误
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}; filename={encoded_filename}",
            "Access-Control-Expose-Headers": "Content-Disposition",
            # 时间格式本身是ASCII字符，无需编码，保持原样
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
@app.post("/doc_editor/merge_docx_office_server", summary="合并拆分的DOCX节点")
async def merge_docx_office_server(
        request: Request,
        node_id: int = Body(..., description="要更新的节点ID"),
        html_content: str = Body(..., description="需要转换的HTML文本"),
        filename: Optional[str] = Body("output.docx", description="下载的DOCX文件名（默认output.docx）"),
        title_text: Optional[str] = Body(None, description="可选：更新节点标题文本")
):
    """调用合并接口生成合并后的DOCX文件流"""
    if node_id <= 0:
        return unified_response(code=400, message="节点ID必须为正整数", data={})
    if not html_content.strip():
        return unified_response(code=400, message="HTML内容不能为空", data={})

    # ── 1. 查节点，拿 record_id ──────────────────────────────────────────
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(
                "SELECT id, record_id FROM \"yxdl_docx_title_trees\" WHERE id = %s",
                (node_id,)
            )
            row = cursor.fetchone()
            if not row:
                return unified_response(
                    code=404,
                    message=f"未找到ID为{node_id}的标题树节点",
                    data={}
                )
            result_record_id = row[1]
    logger.info(f"merge_docx_office_server: node_id={node_id} record_id={result_record_id}")

    # ── 2. 与 _query_and_build_tree 完全相同的树构造逻辑 ─────────────────
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
        return unified_response(code=500, message=f"查询数据库失败：{str(e)}", data={})

    if not node_records:
        return unified_response(code=404, message="该记录下无任何节点", data={})

    # 字段映射：数据库行 → TreeItem（与 _query_and_build_tree 逻辑一致）
    tree_nodes_org = [
        TreeItem(**{
            "id":                       item.get("split_id"),           # ← 使用拆分接口返回的原始id，用于合并时还原树
            "text":                     item.get("title_text"),
            "level":                    item.get("level"),
            "eid":                      item.get("eid"),
            "idx":                      item.get("idx"),
            "parent_id":                item.get("parent_id"),
            "file_path":                item.get("origin_file_path"),
            "update_file_path":         item.get("update_file_path", ""),
            "is_conversion_completion": item.get("is_conversion_completion", 0),
            "children":                 [],
            "file_name":                None,
            "file_info":                None,
            "node_type":                ""
        })
        for item in node_records
    ]

    # build_simplified_tree 按 parent_id 组装嵌套结构
    nested_dicts = build_simplified_tree(node_records)

    # ── 2.5 根据最新树结构刷新 level / idx，并批量回写数据库 ────────────
    def _refresh_level_idx(nodes: List[Dict], parent_level: int = 0, counter: List[int] = None) -> List[Dict]:
        """
        递归重算每个节点的 level / idx：
          - level = parent_level + 1
          - idx   = 全局 DFS 访问顺序，从 0 递增，整棵树唯一不重复
        counter 用列表包装以实现跨递归层共享（模拟引用传递）
        """
        if counter is None:
            counter = [0]
        for node in nodes:
            node['level'] = parent_level + 1
            node['idx'] = counter[0]
            counter[0] += 1
            if node.get('children'):
                _refresh_level_idx(node['children'], parent_level + 1, counter)
        return nodes

    nested_dicts = _refresh_level_idx(nested_dicts)

    # 收集所有变更节点（DFS 平铺）
    def _collect_updates(nodes: List[Dict]) -> List[Dict]:
        result = []
        for node in nodes:
            result.append({'id': node['id'], 'level': node['level'], 'idx': node['idx']})
            if node.get('children'):
                result.extend(_collect_updates(node['children']))
        return result

    updates = _collect_updates(nested_dicts)

    if updates:
        try:
            with get_db_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.executemany(
                        """
                        UPDATE "yxdl_docx_title_trees"
                        SET level = %s, idx = %s, update_time = %s
                        WHERE id = %s
                        """,
                        [(u['level'], u['idx'], current_time, u['id']) for u in updates]
                    )
                    conn.commit()
            logger.info(
                f"merge_docx_office_server: 刷新 level/idx 完成，共更新 {len(updates)} 个节点"
            )
        except Exception as e:
            logger.error(f"merge_docx_office_server: 刷新 level/idx 失败 err={e}")
            return unified_response(code=500, message=f"刷新节点层级失败：{str(e)}", data={})

    # 后续 _dicts_to_tree_items / _collect_files 照常执行，此时 nested_dicts 中的
    # level/idx 已经是刷新后的值，TreeItem 会拿到正确数据
    def _dicts_to_tree_items(nodes_dict: List[Dict]) -> List[TreeItem]:
        result = []
        for d in nodes_dict:
            item = TreeItem(
                eid=d.get("eid", ""),
                level=d.get("level", 1),
                idx=d.get("idx", 0),
                text=d.get("text", "") or d.get("title_text", ""),
                children=_dicts_to_tree_items(d.get("children", []))
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

    # 转为标准返回格式（复用 process_split_tree_nodes_with_select）
    tree_ = nested_tree_items  # 直接用 TreeItem 列表，MergeRequest.tree 类型匹配
    print(tree_)
    def _collect_files(nodes: List[TreeItem]) -> List[str]:
        """DFS 遍历 TreeItem 树，按节点顺序收集文件路径（去重保序）"""
        paths: List[str] = []
        seen: set = set()

        def _dfs(node_list):
            for node in node_list:
                if node.is_conversion_completion == 1 and node.update_file_path:
                    file_path = node.update_file_path
                else:
                    file_path = node.file_path or ""
                if file_path and file_path not in seen:
                    seen.add(file_path)
                    paths.append(file_path)
                _dfs(node.children or [])

        _dfs(nodes)
        return paths

    files_ = _collect_files(tree_)

    # ── 4. 构造合并请求并调用合并接口 ───────────────────────────────────
    format_config = {
        "Heading": {
            "Heading1": {
                "use": True,
                "style": {
                    "alignment": "left",
                    "line_spacing": "single",
                    "line_spacing_value": 1,
                    "left_indent": 0,
                    "right_indent": 0,
                    "space_before": 0,
                    "space_after": 0,
                    "first_line_indent": 0,
                    "font_name": "仿宋",
                    "font_size": "初号",
                    "bold": True,
                    "italic": True,
                    "underline": None,
                    "color": None
                }
            },
            "Heading2": {
                "use": False,
                "style": {
                    "alignment": "left",
                    "line_spacing": "single",
                    "line_spacing_value": 1,
                    "left_indent": 0,
                    "right_indent": 0,
                    "space_before": 0,
                    "space_after": 0,
                    "first_line_indent": 0,
                    "font_name": "仿宋",
                    "font_size": "三号",
                    "bold": True,
                    "italic": False,
                    "underline": None,
                    "color": None
                }
            },
            "Heading3": {
                "use": False,
                "style": {
                    "alignment": "left",
                    "line_spacing": "single",
                    "line_spacing_value": 1,
                    "left_indent": 0,
                    "right_indent": 0,
                    "space_before": 0,
                    "space_after": 0,
                    "first_line_indent": 0,
                    "font_name": "仿宋",
                    "font_size": "三号",
                    "bold": True,
                    "italic": False,
                    "underline": None,
                    "color": None
                }
            },
            "Heading4": {
                "use": False,
                "style": {
                    "alignment": "left",
                    "line_spacing": "single",
                    "line_spacing_value": 1,
                    "left_indent": 0,
                    "right_indent": 0,
                    "space_before": 0,
                    "space_after": 0,
                    "first_line_indent": 0,
                    "font_name": "仿宋",
                    "font_size": "四号",
                    "bold": True,
                    "italic": False,
                    "underline": None,
                    "color": None
                }
            },
            "Heading5": {
                "use": False,
                "style": {
                    "alignment": "left",
                    "line_spacing": "single",
                    "line_spacing_value": 1,
                    "left_indent": 0,
                    "right_indent": 0,
                    "space_before": 0,
                    "space_after": 0,
                    "first_line_indent": 0,
                    "font_name": "仿宋",
                    "font_size": "四号",
                    "bold": True,
                    "italic": False,
                    "underline": None,
                    "color": None
                }
            },
            "Heading6": {
                "use": False,
                "style": {
                    "alignment": "left",
                    "line_spacing": "single",
                    "line_spacing_value": 1,
                    "left_indent": 0,
                    "right_indent": 0,
                    "space_before": 0,
                    "space_after": 0,
                    "first_line_indent": 0,
                    "font_name": "仿宋",
                    "font_size": "小四",
                    "bold": True,
                    "italic": False,
                    "underline": None,
                    "color": None
                }
            },
            "Heading7": {
                "use": False,
                "style": {
                    "alignment": "left",
                    "line_spacing": "single",
                    "line_spacing_value": 1,
                    "left_indent": 0,
                    "right_indent": 0,
                    "space_before": 0,
                    "space_after": 0,
                    "first_line_indent": 0,
                    "font_name": "仿宋",
                    "font_size": "小四",
                    "bold": True,
                    "italic": False,
                    "underline": None,
                    "color": None
                }
            },
            "Heading8": {
                "use": False,
                "style": {
                    "alignment": "left",
                    "line_spacing": "single",
                    "line_spacing_value": 1,
                    "left_indent": 0,
                    "right_indent": 0,
                    "space_before": 0,
                    "space_after": 0,
                    "first_line_indent": 0,
                    "font_name": "仿宋",
                    "font_size": "小四",
                    "bold": True,
                    "italic": False,
                    "underline": None,
                    "color": None
                }
            },
            "Heading9": {
                "use": False,
                "style": {
                    "alignment": "left",
                    "line_spacing": "single",
                    "line_spacing_value": 1,
                    "left_indent": 0,
                    "right_indent": 0,
                    "space_before": 0,
                    "space_after": 0,
                    "first_line_indent": 0,
                    "font_name": "仿宋",
                    "font_size": "五号",
                    "bold": True,
                    "italic": False,
                    "underline": None,
                    "color": None
                }
            }
        },
        "Text": {
            "use": True,
            "style": {
                "alignment": "left",
                "line_spacing": "single",
                "line_spacing_value": 1,
                "left_indent": 0,
                "right_indent": 0,
                "space_before": 0,
                "space_after": 0,
                "first_line_indent": 0,
                "font_name": "仿宋",
                "font_size": "小四",
                "bold": None,
                "italic": None,
                "underline": True,
                "color": None
            }
        },
        "Table": {
            "use": False,
            "style": {
                "repeat_header": False,
                "line_break": False,
                "alignment": "left",
                "font_name": "仿宋",
                "font_size": "五号",
                "left_indent": 0,
                "right_indent": 0,
                "first_line_indent": 0,
                "bold": None,
                "italic": None,
                "underline": None,
                "color": None
            }
        },
        "Header": {
            "use": True,
            "show_logo": True,
            "logo": "",
            "show_name": False,
            "name": "123"
        },
        "Footer": {
            "use": False,
            "style": {
                "alignment": "left"
            }
        },
        "other": {
            "numbering": True,
            "use": False
        },
        "Margin": {
            "use": False,
            "top": 2.54,
            "bottom": 2.54,
            "left": 3.18,
            "right": 3.18
        }
    }
    format_args = {
        "config_dict": format_config,
        "token": "984f5b0a2793eeafeeddfd2cd095ad31",
        "key": "984f5b0a2793eeafeeddfd2cd095ad31-1772598822992"
    }

    try:
        merge_request = MergeRequest(tree=tree_, files=files_, format_args=format_args)
        merged_file_message = call_docx_merge(merge_request)
        return merged_file_message
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"文件合并失败：{str(e)}")



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
        html_content, temp_file_docx = docx_to_html(save_path)
        print(save_path)
        save_path2 = os.path.join(UPLOAD_DIR, "专利报告.html")
        with open(save_path2, 'w', encoding='utf-8') as f:
            f.write(html_content)
        try:
            if os.path.exists(save_path):
                os.remove(save_path)
        except Exception as e:
            print(f"警告：无法删除临时文件 {save_path} - {e}")
        return unified_response(
            code=200,
            message="节点HTML内容更新成功",
            data={
                "html_content": html_content
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成文档时出错: {str(e)}")


@app.post("/doc_editor/generate_patent_doc/patent_generator", summary="专利生成器", description="""
    生成包含专利表格和证书图片的Word文档，并转换为单文件HTML
    - 表格列：序号、专利类型、专利名称、专利号、专利权人、授权公告日
    - 图片特性：自动补图填充空白区域，支持奇数图片的两种显示模式
    """
)
async def generate_default_patent_doc_patent_generator(
        patent_data: List[List[Any]] = Body(
            ...,
            description="专利数据二维列表（必填），每行对应一条专利信息，顺序：[序号,专利类型,专利名称,专利号,专利权人,授权公告日]",
            example=[
                ["1", "发明专利", "一种智能温控系统", "ZL202410000001.0", "某某科技有限公司", "2024-01-15"],
                ["2", "实用新型专利", "一种节能电机", "ZL202420000001.0", "某某科技有限公司", "2024-02-20"]
            ],
            min_items=1,  # 至少1条专利数据
            max_items=100  # 最多100条专利数据
        ),
        cert_img_urls: List[str] = Body(
            [],
            description="专利证书图片的网络URL列表（可选）",
            example=["https://example.com/patent1.jpg", "https://example.com/patent2.png"],
            min_items=0
        ),
        fill_empty_space: bool = Body(
            False,
            description="是否自动补图填充页面空白区域（重复使用现有图片）",
            example=True
        ),
        last_img_display_mode: int = Body(
            1,
            description="奇数图片时最后一行的显示模式：1=删除空单元格，2=合并单元格显示大图",
            example=1,
            ge=1,
            le=2
        )
):
    try:
        # 数据校验：确保每行都有6个字段
        for idx, row in enumerate(patent_data):
            if len(row) != 6:
                raise ValueError(
                    f"专利数据第{idx + 1}行必须包含6个字段（序号、专利类型、专利名称、专利号、专利权人、授权公告日）")

        # 下载图片
        cert_img_paths = await download_images(cert_img_urls)

        # 生成文档并转换为HTML
        html_content = generate_and_convert_to_html(
            generate_fully_centered_patent_doc,
            patent_data=patent_data,
            cert_img_paths=cert_img_paths,
            fill_empty_space=fill_empty_space,
            last_img_display_mode=last_img_display_mode
        )

        # 清理下载的图片文件
        try:
            for img_path in cert_img_paths:
                os.remove(img_path)
        except:
            pass

        if not html_content:
            raise HTTPException(status_code=500, detail="HTML转换失败，生成的HTML内容为空")

        return unified_response(
            code=200,
            message="生成专利生成器HTML内容成功",
            data={
                "html_content": html_content
            }
        )

    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"生成专利文档失败：{str(e)}")


@app.post("/doc_editor/generate_patent_doc/financial_report_generator", summary="财报生成器", description="""
    生成包含标题和多张图片的年度报告Word文档，并转换为单文件HTML返回
    - 文档结构：表格布局，第一行是标题，第二行是主图，后续行是多列排版的其他图片
    - HTML特性：图片Base64内嵌，CSS内联，无外部资源依赖
    """
)
async def generate_default_patent_doc_financial_report_generator(
        title_text: str = Body(
            ...,
            description="报告标题文本（显示在表格第一行）",
            example="2024年度销售报告",
            min_length=1,
            max_length=100
        ),
        second_row_img_url: str = Body(
            ...,
            description="第二行主图片的网络URL（必填），支持HTTP/HTTPS",
            example="https://example.com/main_chart.jpg",
            regex=r'^https?://.*\.(jpg|jpeg|png|gif|bmp)$'
        ),
        other_img_urls: List[str] = Body(
            [],
            description="其他图片的网络URL列表（可选），会按列数排版在主图下方",
            example=["https://example.com/img1.jpg", "https://example.com/img2.png"],
            min_items=0
        ),
        columns_per_row: int = Body(
            2,
            description="每行显示的图片列数（主图除外）",
            example=2,
            ge=1,  # 大于等于1
            le=4  # 小于等于4
        )
):
    try:
        # 下载图片
        main_img_path = await download_image(second_row_img_url)
        other_img_paths = await download_images(other_img_urls)

        # 生成文档并转换为HTML
        html_content = generate_and_convert_to_html(
            generate_report_doc,
            title_text=title_text,
            second_row_img_path=main_img_path,
            other_img_paths=other_img_paths,
            columns_per_row=columns_per_row
        )

        # 清理下载的图片文件
        try:
            os.remove(main_img_path)
            for img_path in other_img_paths:
                os.remove(img_path)
        except:
            pass

        if not html_content:
            raise HTTPException(status_code=500, detail="HTML转换失败，生成的HTML内容为空")

        return unified_response(
            code=200,
            message="生成财报生成器HTML内容成功",
            data={
                "html_content": html_content
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"生成报告失败：{str(e)}")


@app.post("/doc_editor/generate_patent_doc/vehicle_generator", summary="车辆生成器",description="""
    生成包含公司车辆信息的Word文档，并转换为单文件HTML
    - 表格列：序号、车牌号、行驶证图片、车辆图片
    - 图片特性：行驶证和车辆图片分别按固定宽度显示
    """)
async def generate_default_patent_doc_vehicle_generator(
    car_data: List[List[Any]] = Body(
        ...,
        description="车辆数据列表，每行：[序号,车牌号,行驶证URL,车辆图片URL]",
        example=[
            ["1", "京A12345", "https://example.com/drive1.jpg", "https://example.com/car1.jpg"],
            ["2", "沪B67890", "https://example.com/drive2.jpg", "https://example.com/car2.jpg"]
        ],
        min_items=1
    ),
    table_title: str = Body(
        "公司车辆信息",
        description="表格标题",
        example="2024年运营车辆信息",
        min_length=1,
        max_length=50
    )
) -> JSONResponse:
    try:
        # 下载所有车辆相关图片
        car_data_with_local_paths = []
        all_img_urls = []
        img_url_map = {}

        # 收集所有图片URL
        for car_row in car_data:
            seq, plate_num, drive_url, car_url = car_row
            all_img_urls.append(drive_url)
            all_img_urls.append(car_url)

        # 批量下载图片
        img_paths = await download_images(all_img_urls)
        img_path_map = dict(zip(all_img_urls, img_paths))

        # 替换URL为本地路径
        for car_row in car_data:
            seq, plate_num, drive_url, car_url = car_row
            drive_path = img_path_map.get(drive_url, "")
            car_path = img_path_map.get(car_url, "")
            car_data_with_local_paths.append([seq, plate_num, drive_path, car_path])

        # 生成文档并转换为HTML
        html_content = generate_and_convert_to_html(
            generate_car_info_doc,
            car_data=car_data_with_local_paths,
            table_title=table_title
        )

        html_content_dict = {
                "html_content": html_content
            }
        # 清理下载的图片文件
        try:
            for img_path in img_paths:
                os.remove(img_path)
        except:
            pass

        if not html_content:
            raise HTTPException(status_code=500, detail="HTML转换失败")

        return unified_response(
            code=200,
            message="生成车辆生成器HTML内容成功",
            data=html_content_dict
        )

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/test-use-config", summary="4测试在业务逻辑中使用配置")
async def test_use_config():
    """
    示例：在实际业务逻辑中读取并使用上传路径配置
    """
    try:
        # 读取上传配置
        uploads_config = get_server_uploads_config()
        local_path = uploads_config["user_local_path"]
        web_path = uploads_config["web_path"]

        # 模拟业务逻辑：拼接文件完整URL
        filename = "test_file.pdf"
        local_file_path = os.path.join(local_path, filename)
        web_file_url = web_path + filename

        return {
            "local_file_path": local_file_path,
            "web_file_url": web_file_url,
            "original_config": uploads_config
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/doc_editor/generator_query_by_type", summary="查询生成器格式")
async def query_format_storage_by_type(request: Request, formant_type: Any = Body(..., description="type类型，format_storage_id"),table_title: str = Body(
        "",
        description="站位数据",
        example=""
    )) -> JSONResponse:
    """
    通过 type 查询配置格式存储数据
    - request.body.type: 类型值（1=默认格式），必填且为整数
    """
    # 1. 解析并验证请求体
    try:
        # 获取请求体并解析为 JSON
        request_body = await request.json()
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="请求体格式错误，必须是有效的 JSON")

    # 检查 type 参数是否存在
    if "formant_type" not in request_body:
        raise HTTPException(status_code=400, detail="缺少必填参数：formant_type")

    # 验证 type 参数类型（必须是整数）
    type_value = formant_type
    if not isinstance(type_value, int):
        # 尝试转换为整数（兼容前端传字符串数字的情况）
        try:
            type_value = int(type_value)
        except (ValueError, TypeError):
            raise HTTPException(status_code=400, detail="参数 type 必须是整数")

    # 2. 执行原生 SQL 查询
    query_sql = """
        SELECT id, format_name, base64_img 
        FROM cfg_format_storage
        WHERE type = %s  and status = 1;
    """

    try:
        # 获取数据库连接并执行查询
        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                # 执行参数化查询，防止 SQL 注入
                cur.execute(query_sql, (type_value,))
                # 获取查询结果并转换为普通字典列表
                results = [dict(row) for row in cur.fetchall()]

        # 3. 返回查询结果
        return unified_response(200,"查询成功", results)

    except psycopg2.Error as e:
        raise HTTPException(status_code=500, detail=f"数据库查询失败：{str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"查询失败：{str(e)}")


@app.post("/doc_editor/file_slicing_download", summary="文件分片下载", response_model=None)
async def html_to_docx_api(request: Request,
        file_path: str = Body(..., description="文件的完整URL路径（如http://xxx/temp.docx）"),
        filename: str = Body(..., description="文件名"),
) -> Union[JSONResponse, StreamingResponse]:
    """接收文件路径分片给文件流"""
    try:


        # 校验文件名格式
        new_file_path = file_path
        response = requests.get(new_file_path, timeout=30)
        # 校验响应状态码（200表示成功）
        response.raise_for_status()
        temp_docx_filename = generate_unique_filename("temp.docx")
        abs_file_path = os.path.join(UPLOAD_DIR, temp_docx_filename)
        # abs_file_path = os.path.abspath("temp.docx")

        # 将文件内容写入本地
        with open(abs_file_path, 'wb') as f:
            f.write(response.content)
        file_pathlib = pathlib.Path(str(abs_file_path))
        response = file_resp.FileResp(request, file_pathlib).start()
        return response

    except Exception as e:
        return unified_response(
            code=500,
            message=f"HTML转DOCX失败：{str(e)}",
            data={}
        )


LIBREOFFICE_PATH = "libreoffice"


@app.post("/test-liboffice/emf-to-png", summary="测试LibreOffice EMF转PNG功能")
async def test_libreoffice_emf2png(
        background_tasks: BackgroundTasks,  # 注入后台任务对象
        file: UploadFile = File(..., description="上传EMF格式文件")
):
    """
    测试 LibreOffice 是否能正常将 EMF 文件转换为 PNG：
    1. 接收上传的 EMF 文件
    2. 调用 LibreOffice 命令行转换
    3. 返回转换后的 PNG 文件
    4. 立即清理本次请求产生的所有临时文件
    """
    # 为每个请求创建独立的临时目录（隔离不同请求的文件）
    temp_dir = tempfile.mkdtemp(prefix="libreoffice_test_")
    emf_file_path = None
    png_file_path = None

    try:
        # 1. 校验文件格式
        if not file.filename.lower().endswith(".emf"):
            raise HTTPException(status_code=400, detail="仅支持上传 .emf 格式文件")

        # 2. 保存上传的 EMF 文件到临时目录（原生同步写入，替代aiofiles）
        emf_filename = os.path.basename(file.filename)
        emf_file_path = os.path.join(temp_dir, emf_filename)
        # 读取文件内容（异步读取，同步写入，兼容FastAPI）
        file_content = await file.read()
        with open(emf_file_path, "wb") as f:
            f.write(file_content)

        # 3. 构造转换命令（带透明+300DPI参数）
        png_filename = os.path.splitext(emf_filename)[0] + ".png"
        png_file_path = os.path.join(temp_dir, png_filename)

        cmd = [
            LIBREOFFICE_PATH,
            "--headless",  # 无界面运行
            "--norestore",  # 防止恢复提示
            "--nolockcheck",  # 避免锁文件问题
            "--convert-to",
            'png:draw_png_Export:{"Translucent":true,"Resolution":300}',  # 透明+300DPI
            emf_file_path,
            "--outdir",
            temp_dir
        ]

        # 4. 执行转换并捕获结果
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60  # 超时时间60秒
        )

        # 检查命令执行状态
        if result.returncode != 0:
            raise HTTPException(
                status_code=500,
                detail=f"LibreOffice 转换失败：\n标准输出：{result.stdout}\n错误输出：{result.stderr}"
            )

        # 检查PNG文件是否生成
        if not os.path.exists(png_file_path):
            raise HTTPException(
                status_code=500,
                detail="转换命令执行成功，但未生成PNG文件"
            )

        # 5. 注册后台清理任务，返回PNG文件
        background_tasks.add_task(cleanup_temp_files, temp_dir=temp_dir)
        return FileResponse(
            path=png_file_path,
            filename=png_filename,
            media_type="image/png"
        )

    except subprocess.TimeoutExpired:
        # 异常时先清理临时文件，再抛错
        cleanup_temp_files(temp_dir)
        raise HTTPException(status_code=500, detail="LibreOffice 转换超时（60秒）")
    except Exception as e:
        # 所有异常场景都先清理临时文件
        cleanup_temp_files(temp_dir)
        raise HTTPException(status_code=500, detail=f"转换过程出错：{str(e)}")


@app.get("/health", summary="接口健康检查")
async def health_check():
    """检查接口是否可用，同时验证LibreOffice是否能调用"""
    try:
        result = subprocess.run(
            [LIBREOFFICE_PATH, "--version"],
            capture_output=True,
            text=True,
            timeout=10
        )
        if result.returncode != 0:
            return {
                "status": "unhealthy",
                "message": "LibreOffice 调用失败",
                "error": result.stderr
            }
        return {
            "status": "healthy",
            "message": "接口和LibreOffice均正常",
            "libreoffice_version": result.stdout.strip()
        }
    except Exception as e:
        return {
            "status": "unhealthy",
            "message": "接口异常",
            "error": str(e)
        }


# 通用清理函数：删除临时目录及所有文件
def cleanup_temp_files(temp_dir: str):
    """删除指定的临时目录，忽略不存在的目录"""
    if os.path.exists(temp_dir):
        try:
            shutil.rmtree(temp_dir)
            print(f"临时目录 {temp_dir} 已清理")
        except Exception as e:
            print(f"清理临时目录失败：{str(e)}")


# ====================== 启动服务 ======================
if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        app=__name__ + ":app",
        host="0.0.0.0",
        port=8080,
        reload=True,
        workers=4
    )