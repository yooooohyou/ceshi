import base64
import logging
import os
import re
import shutil
import tempfile
import time
import uuid

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)


def html_base64_images_to_urls(
    html_text: str,
    serve_dir: str,
    base_url: str,
    cleanup_delay: int = 1800,
) -> tuple:
    """将 HTML 中内嵌的 base64 图片保存到 serve_dir 并替换 src 为可访问 URL。

    Returns:
        (new_html, session_img_dir)  — 无图片时 session_img_dir 为 None。
    """
    _b64_src_re = re.compile(
        r'src="data:(image/[^;]+);base64,([^"]+)"',
        re.IGNORECASE,
    )
    _mime_to_ext = {
        "image/png": "png", "image/jpeg": "jpg", "image/gif": "gif",
        "image/webp": "webp", "image/bmp": "bmp", "image/svg+xml": "svg",
        "image/tiff": "tif",
    }

    if not _b64_src_re.search(html_text):
        return html_text, None

    session_id = uuid.uuid4().hex
    session_img_dir = os.path.join(serve_dir, f"docximg_{session_id}")
    os.makedirs(session_img_dir, exist_ok=True)

    def _replace(m):
        mime = m.group(1).lower()
        b64data = m.group(2)
        try:
            img_bytes = base64.b64decode(b64data)
        except Exception:
            return m.group(0)
        ext = _mime_to_ext.get(mime, "png")
        filename = f"{uuid.uuid4().hex}.{ext}"
        filepath = os.path.join(session_img_dir, filename)
        try:
            with open(filepath, "wb") as f:
                f.write(img_bytes)
        except Exception as e:
            logger.warning(f"保存图片失败 {filepath}: {e}")
            return m.group(0)
        url = f"{base_url.rstrip('/')}/docximg_{session_id}/{filename}"
        return f'src="{url}"'

    new_html = _b64_src_re.sub(_replace, html_text)
    return new_html, session_img_dir


def download_image_to_base64(image_url: str, base_url: str = None, timeout: int = 10):
    """下载图片并转换为 base64 字符串，返回 (base64_str, content_type)"""
    temp_file_path = None
    try:
        image_url = image_url.strip().split()[0]
        if image_url.startswith(('"', "'")) and image_url.endswith(('"', "'")):
            image_url = image_url[1:-1]

        if image_url.startswith("data:"):
            try:
                header, data_part = image_url.split(",", 1)
                meta = header[5:]
                parts = meta.split(";")
                content_type = parts[0] if parts[0] else "image/jpeg"
                if "base64" in parts:
                    return data_part, content_type
                else:
                    import urllib.parse
                    decoded = urllib.parse.unquote_to_bytes(data_part)
                    return base64.b64encode(decoded).decode("utf-8"), content_type
            except Exception as e:
                logger.debug(f"解析 data: URI 失败: {e}")
                return None, None

        if base_url and not image_url.startswith(("http://", "https://")):
            image_url = f"{base_url.rstrip('/')}/{image_url.lstrip('/')}"

        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        }
        response = requests.get(image_url, headers=headers, timeout=timeout, stream=True, verify=True)
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "image/jpeg")
        suffix = f".{content_type.split('/')[-1]}" if "/" in content_type else ".jpg"
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        temp_file_path = temp_file.name
        temp_file.close()

        with open(temp_file_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        with open(temp_file_path, "rb") as f:
            base64_encoded = base64.b64encode(f.read()).decode("utf-8")

        return base64_encoded, content_type

    except Exception as e:
        logger.debug(f"下载/转换图片失败 {image_url}: {e}")
        return None, None
    finally:
        if temp_file_path and os.path.exists(temp_file_path):
            for _ in range(3):
                try:
                    os.unlink(temp_file_path)
                    break
                except Exception:
                    time.sleep(0.1)


def html_img_url_to_base64(html_text: str, base_url: str = None, timeout: int = 10):
    """将 HTML 中所有 <img> 的网络 src 替换为 base64 data URI。

    Returns:
        (processed_html, stats_dict)
    """
    temp_dir = tempfile.mkdtemp(prefix="img_base64_re_")
    try:
        full_img_pattern = re.compile(r"<img[^>]+>", re.IGNORECASE | re.DOTALL)
        img_tags = full_img_pattern.findall(html_text)
        if not img_tags:
            logger.debug("未找到任何img标签，直接返回原HTML")
            return html_text, {"success": 0, "fail": 0}

        replacement_map = {}
        success_count = 0
        fail_count = 0
        src_pattern = re.compile(r'src\s*=\s*(?:"([^"]+)"|\'([^\']+)\'|([^\s>]+))', re.IGNORECASE)

        for original_img_tag in img_tags:
            if original_img_tag in replacement_map:
                continue
            src_match = src_pattern.search(original_img_tag)
            if not src_match:
                replacement_map[original_img_tag] = original_img_tag
                fail_count += 1
                continue

            img_url = next((v for v in src_match.groups() if v is not None), "").strip()
            if not img_url:
                replacement_map[original_img_tag] = original_img_tag
                fail_count += 1
                continue

            base64_str, content_type = download_image_to_base64(img_url, base_url, timeout)
            if base64_str and content_type:
                new_img_tag = src_pattern.sub(
                    f'src="data:{content_type};base64,{base64_str}"',
                    original_img_tag,
                    count=1,
                )
                replacement_map[original_img_tag] = new_img_tag
                success_count += 1
            else:
                replacement_map[original_img_tag] = original_img_tag
                fail_count += 1

        processed_html = html_text
        for original_img_tag in img_tags:
            processed_html = processed_html.replace(original_img_tag, replacement_map[original_img_tag], 1)

        return processed_html, {"success": success_count, "fail": fail_count, "total": len(img_tags)}
    finally:
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass


def get_html_heading_levels(html_content: str):
    """返回 (existing_levels, max_level)"""
    if not html_content or not isinstance(html_content, str):
        return [], 0
    soup = BeautifulSoup(html_content, "html.parser")
    headings = soup.find_all(re.compile(r"^h[1-6]$", re.IGNORECASE))
    existing_levels = sorted({int(h.name[1]) for h in headings})
    max_level = max(existing_levels) if existing_levels else 0
    return existing_levels, max_level


def limit_html_heading_levels(html_content: str, max_allowed_level: int) -> str:
    """将超过 max_allowed_level 的标题降级；0 表示去掉标题标签只保留内容"""
    if not isinstance(max_allowed_level, int) or not (0 <= max_allowed_level <= 6):
        raise ValueError("max_allowed_level必须是0-6之间的整数")
    if not html_content or not isinstance(html_content, str):
        return html_content

    soup = BeautifulSoup(html_content, "html5lib")
    for heading in soup.find_all(re.compile(r"^h[1-6]$", re.IGNORECASE)):
        current_level = int(heading.name[1])
        if max_allowed_level == 0:
            heading.replace_with(*heading.contents)
        elif current_level > max_allowed_level:
            new_heading = soup.new_tag(f"h{max_allowed_level}")
            new_heading.contents = heading.contents
            new_heading.attrs = heading.attrs
            heading.replace_with(new_heading)
    return soup.prettify()


def merge_html_texts(html_list: list) -> str:
    """合并多个 HTML 文本，提取各自 <body> 内容后拼接"""
    merged_body_parts = []
    for html in html_list:
        soup = BeautifulSoup(html, "html.parser")
        body = soup.body
        merged_body_parts.append(body.decode_contents() if body else str(soup))
    merged_body = "\n".join(merged_body_parts)
    return f"<!DOCTYPE html>\n<html>\n<body>\n{merged_body}\n</body>\n</html>"
