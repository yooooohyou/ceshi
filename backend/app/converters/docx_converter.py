import io
import logging
import os
import time

import requests

from app.core.config import UPLOAD_DIR, STATIC_WEB_FRONT_PREFIX
from app.utils.file_utils import generate_unique_filename
from app.utils.html_utils import (
    add_nowrap_to_signature_paragraphs,
    hide_mce_anchor_tags,
    html_base64_images_to_urls,
)
from app.utils.path_utils import judge_path_type

logger = logging.getLogger(__name__)


def _wait_until_docx_readable(path: str, attempts: int = 3, delay: float = 0.4) -> bool:
    """
    偶发：拆分服务返回的小 docx 路径在 NFS/共享挂载上短暂不可见，或文件刚写入
    尚未完成（PK\\x03\\x04 zip 头未就绪）。短延迟轮询直到可读，避免 python-docx
    抛 "Package not found" / "File is not a zip file"。
    """
    last_err = None
    for i in range(attempts):
        try:
            if not os.path.exists(path):
                last_err = "path not exist"
            elif os.path.getsize(path) < 32:
                last_err = "file too small"
            else:
                with open(path, "rb") as f:
                    head = f.read(4)
                if head == b"PK\x03\x04":
                    return True
                last_err = f"bad header {head!r}"
        except Exception as e:
            last_err = f"{type(e).__name__}: {e}"
        if i < attempts - 1:
            time.sleep(delay)
    logger.warning(f"docx_to_html: 等待文件可读超时 path={path} last_err={last_err}")
    return False


def docx_to_html(file_path: str):
    """DOCX 转 HTML，返回 (html_str, abs_file_path)"""
    abs_file_path = ""
    try:
        if judge_path_type(file_path) == "web":
            response = requests.get(file_path, timeout=30)
            response.raise_for_status()
            temp_docx_filename = generate_unique_filename("temp.docx")
            abs_file_path = os.path.join(UPLOAD_DIR, temp_docx_filename)
            with open(abs_file_path, "wb") as f:
                f.write(response.content)
        else:
            abs_file_path = file_path

        # 拆分接口返回的小 docx 偶发不存在 / 写入未完成 / 不是合法 zip。
        # 不可用时直接返回空 HTML，不再写"转换失败"到内容里污染下游合并结果。
        if not _wait_until_docx_readable(abs_file_path):
            return "", abs_file_path

        file_size = os.path.getsize(abs_file_path)
        if file_size > 10 * 1024 * 1024:
            logger.warning(f"警告：文件过大（{file_size / 1024 / 1024:.2f}MB），可能转换失败")

        from docxhtmlcoverter import DocxHtmlConverter
        converter = DocxHtmlConverter()
        temp_html_filename = generate_unique_filename("temp.html")
        temp_html_path = os.path.join(UPLOAD_DIR, temp_html_filename)

        try:
            html_content = converter.docx_to_single_html(abs_file_path, temp_html_path)
        except Exception as e_first:
            # python-docx 在边缘场景偶发 Package not found / not a zip：
            # 兼容 NFS/共享存储的短暂可见性问题，再次等待并重试一次。
            logger.warning(
                f"docx_to_html: 首次转换异常将重试 path={abs_file_path} "
                f"err={type(e_first).__name__}: {e_first}"
            )
            time.sleep(0.5)
            if not _wait_until_docx_readable(abs_file_path):
                return "", abs_file_path
            html_content = converter.docx_to_single_html(abs_file_path, temp_html_path)

        if os.path.exists(temp_html_path):
            try:
                with open(temp_html_path, "r", encoding="utf-8") as f:
                    html_content = f.read()
            except UnicodeDecodeError:
                with open(temp_html_path, "r", encoding="gbk") as f:
                    html_content = f.read()
            finally:
                try:
                    os.remove(temp_html_path)
                except Exception as e:
                    logger.warning(f"警告：无法删除临时文件 {temp_html_path} - {e}")

        if html_content:
            html_content, _ = html_base64_images_to_urls(
                html_content, UPLOAD_DIR, STATIC_WEB_FRONT_PREFIX
            )
            html_content = hide_mce_anchor_tags(html_content)
            html_content = add_nowrap_to_signature_paragraphs(html_content)


        return html_content or "", abs_file_path

    except Exception as e:
        logger.warning(
            f"docx_to_html: Word转HTML失败 path={file_path} "
            f"err={type(e).__name__}: {e}"
        )
        return "", abs_file_path


def convert_html_to_docx(html_content: str):
    """HTML 转 DOCX，返回 (success, stream_or_error_msg, docx_path)"""
    try:
        if not html_content.strip():
            return False, "HTML内容不能为空", ""

        from docxhtmlcoverter import DocxHtmlConverter
        converter = DocxHtmlConverter()
        temp_docx_filename = generate_unique_filename("html2docx.docx")
        temp_docx_path = os.path.join(UPLOAD_DIR, temp_docx_filename)

        converter.html_text_to_docx(html_content, temp_docx_path)

        if not os.path.exists(temp_docx_path):
            return False, f"转换失败：未生成文件 {temp_docx_path}", temp_docx_path

        docx_stream = io.BytesIO()
        with open(temp_docx_path, "rb") as f:
            docx_stream.write(f.read())
        docx_stream.seek(0)

        return True, docx_stream, temp_docx_path

    except PermissionError:
        return False, "权限错误：无法创建/读取临时文件", ""
    except Exception as e:
        return False, f"HTML转DOCX失败：{str(e)}", ""