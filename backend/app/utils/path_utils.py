import logging
import os
import re
import uuid
from urllib.parse import urlparse

from fastapi import Request

from app.core.config import UPLOAD_DIR, STATIC_WEB_FRONT_PREFIX, WEB_File_Path

logger = logging.getLogger(__name__)


def is_web_path_(path_str: str) -> bool:
    """判断路径是否为网页 URL"""
    parsed = urlparse(path_str.strip())
    return parsed.scheme in {"http", "https", "ftp", "ftps"}


def is_local_path_(path_str: str) -> bool:
    """判断路径是否为本地文件路径"""
    path_str = path_str.strip()
    if is_web_path_(path_str):
        return False
    if os.name == "nt":
        has_drive = len(path_str) >= 2 and path_str[1] == ":" and path_str[0].isalpha()
        has_backslash = "\\" in path_str
        return has_drive or has_backslash or os.path.exists(path_str)
    return path_str.startswith("/") or os.path.exists(path_str)


def judge_path_type(path_str: str) -> str:
    """返回路径类型：'web' / 'local' / 'unknown'"""
    if not path_str:
        return "unknown"
    if is_web_path_(path_str):
        return "web"
    if is_local_path_(path_str):
        return "local"
    return "unknown"


def is_ends_with_path_separator(s: str) -> bool:
    """判断字符串是否以路径分隔符结尾"""
    if not s:
        return False
    return bool(re.search(r"[/\\]+$", s))


def local_upload_path_to_web_path(local_abs_path: str, request: Request) -> str:
    """将 uploads 本地绝对路径转换为 Web 路径"""
    if WEB_File_Path:
        return UPLOAD_DIR + "/" + os.path.basename(local_abs_path)

    local_abs_path = os.path.normpath(local_abs_path)
    uploads_local_dir = os.path.normpath(UPLOAD_DIR)
    if not local_abs_path.startswith(uploads_local_dir):
        raise ValueError(f"路径 {local_abs_path} 不在uploads目录下")

    relative_path = local_abs_path[len(uploads_local_dir):]
    full_url = request.url_for("uploads", path=relative_path.lstrip(os.sep))
    return str(full_url)


def save_html_and_get_url(html_content: str) -> str:
    """将 HTML 字符串写入 uploads 目录，返回可访问的 URL"""
    filename = f"html_{uuid.uuid4().hex}.html"
    file_path = os.path.join(UPLOAD_DIR, filename)
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    return STATIC_WEB_FRONT_PREFIX.rstrip("/") + "/" + filename
