import asyncio
import gc
import logging
import os
import shutil
import tempfile
import threading
import time
import uuid

import aiohttp
from PIL import Image

from app.core.config import UPLOAD_DIR, STATIC_WEB_FRONT_PREFIX

logger = logging.getLogger(__name__)


async def download_image(url: str, save_dir: str = UPLOAD_DIR) -> str:
    """异步下载网络图片到本地（空 URL 返回空字符串）"""
    if not url or url.strip() == "":
        return ""
    try:
        file_ext = url.split(".")[-1].split("?")[0]
        if len(file_ext) > 5:
            file_ext = "jpg"
        file_name = f"{uuid.uuid4()}.{file_ext}"
        file_path = os.path.join(save_dir, file_name)

        async with aiohttp.ClientSession() as session:
            async with session.get(url, timeout=aiohttp.ClientTimeout(total=30)) as response:
                if response.status != 200:
                    logger.warning(f"图片下载失败：{url}，状态码：{response.status}")
                    return ""
                with open(file_path, "wb") as f:
                    f.write(await response.read())

        try:
            with Image.open(file_path) as img:
                img.verify()
        except Exception as e:
            os.remove(file_path)
            logger.warning(f"图片文件无效：{url}，错误：{e}")
            return ""

        return file_path
    except Exception as e:
        logger.warning(f"下载图片失败：{url}，错误：{e}")
        return ""


async def download_images(urls: list) -> list:
    """批量下载图片（并发）"""
    if not urls:
        return []
    tasks = [download_image(url) for url in urls]
    return list(await asyncio.gather(*tasks))


def generate_and_convert_to_html(generate_func, *args, **kwargs) -> str:
    """通用函数：调用文档生成函数后转换为 HTML，并用网络 URL 替换 base64 图片"""
    from app.utils.html_utils import html_base64_images_to_urls

    temp_dir = tempfile.mkdtemp(prefix="docx_html_temp_")
    docx_path = os.path.join(temp_dir, f"temp_doc_{uuid.uuid4().hex[:8]}.docx")
    html_path = os.path.join(temp_dir, f"temp_html_{uuid.uuid4().hex[:8]}.html")

    try:
        generate_func(*args, save_path=docx_path, **kwargs)

        gc.collect()
        time.sleep(0.1)

        from docxhtmlcoverter import DocxHtmlConverter
        converter = DocxHtmlConverter()
        html_content = converter.docx_to_single_html(docx_path, html_path)

        if html_content:
            html_content, _ = html_base64_images_to_urls(
                html_content, UPLOAD_DIR, STATIC_WEB_FRONT_PREFIX
            )
        return html_content
    finally:
        def _clean():
            gc.collect()
            time.sleep(0.2)
            for p in (docx_path, html_path):
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except Exception:
                        pass
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)

        t = threading.Thread(target=_clean, daemon=True)
        t.start()
