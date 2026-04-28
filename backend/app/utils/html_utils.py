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


_MCE_ANCHOR_TAG_RE = re.compile(
    r'<a\b[^>]*(?:'
    r'\bclass\s*=\s*(["\'])[^"\']*\bmce-item-anchor\b[^"\']*\1'
    r'|'
    r'\bname\s*=\s*(["\'])_[Tt]oc[^"\']*\2'
    r')[^>]*>',
    re.IGNORECASE,
)
_MCE_ANCHOR_STYLE_RE = re.compile(r'\bstyle\s*=\s*(["\'])([^"\']*)\1', re.IGNORECASE)
_MCE_ANCHOR_DISPLAY_RE = re.compile(r'display\s*:\s*[^;]+;?', re.IGNORECASE)


def hide_mce_anchor_tags(html_content: str) -> str:
    """为 DOCX 转 HTML 产生的目录/书签锚点 <a> 标签追加 style="display:none"。

    覆盖两种形式：
    - <a class="mce-item-anchor" ...>
    - <a name="_TocXXXXXX" ...>（无 mce-item-anchor class 的 Toc 书签）"""
    if not html_content or not (
        "mce-item-anchor" in html_content or "_Toc" in html_content or "_toc" in html_content
    ):
        return html_content

    def _process(match):
        tag = match.group(0)
        style_m = _MCE_ANCHOR_STYLE_RE.search(tag)
        if style_m:
            existing = style_m.group(2).strip()
            if re.search(r'display\s*:', existing, re.IGNORECASE):
                new_style = _MCE_ANCHOR_DISPLAY_RE.sub('display:none;', existing).rstrip(';')
            else:
                new_style = f"{existing};display:none".strip(';')
            quote = style_m.group(1)
            return tag.replace(style_m.group(0), f'style={quote}{new_style}{quote}')
        return re.sub(r'(\s*/?>)$', r' style="display:none"\1', tag, count=1)

    return _MCE_ANCHOR_TAG_RE.sub(_process, html_content)


# 签字栏/日期栏特征段落识别：DOCX 转 HTML 后这类段落往往含大量下划线 span +
# 大量 &nbsp;，前端渲染时连续的 nbsp 会被换行，破坏布局。给这些段落加 white-space:nowrap。
_NOWRAP_PARA_RE = re.compile(r'<p\b([^>]*)>(.*?)</p\s*>', re.IGNORECASE | re.DOTALL)
_UNDERLINE_SPAN_RE = re.compile(
    r'<span\b[^>]*text-decoration\s*:\s*underline[^>]*>', re.IGNORECASE
)
_STYLE_ATTR_RE = re.compile(r'(\bstyle\s*=\s*)(["\'])(.*?)\2', re.IGNORECASE | re.DOTALL)
_DATA_MCE_STYLE_ATTR_RE = re.compile(
    r'(\bdata-mce-style\s*=\s*)(["\'])(.*?)\2', re.IGNORECASE | re.DOTALL
)
_NOWRAP_DECL_RE = re.compile(r'white-space\s*:\s*nowrap', re.IGNORECASE)


def add_nowrap_to_signature_paragraphs(html_content: str) -> str:
    """
    DOCX 转 HTML 后处理：给段落内同时含 ≥10 个 text-decoration:underline 的 <span>
    且 ≥10 个 &nbsp; 的 <p> 加 white-space: nowrap，避免前端渲染时长串 nbsp 换行。
    若 <p> 上同时存在 data-mce-style，则同步追加到 data-mce-style，保持 TinyMCE 一致。
    """
    if not html_content:
        return html_content

    def _replace(m):
        attrs, body = m.group(1), m.group(2)

        if len(_UNDERLINE_SPAN_RE.findall(body)) < 10:
            return m.group(0)
        if body.lower().count('&nbsp;') < 10:
            return m.group(0)

        style_m = _STYLE_ATTR_RE.search(attrs)
        mce_m = _DATA_MCE_STYLE_ATTR_RE.search(attrs)

        style_val = style_m.group(3) if style_m else ''
        mce_val = mce_m.group(3) if mce_m else ''

        style_has = bool(_NOWRAP_DECL_RE.search(style_val))
        mce_has = mce_m is not None and bool(_NOWRAP_DECL_RE.search(mce_val))
        if style_has and (mce_m is None or mce_has):
            return m.group(0)

        def _append_nowrap(css: str) -> str:
            css = (css or '').strip()
            if _NOWRAP_DECL_RE.search(css):
                return css
            if css and not css.endswith(';'):
                css += ';'
            return (css + ' white-space: nowrap;').strip()

        new_attrs = attrs
        if style_m:
            new_attrs = new_attrs.replace(
                style_m.group(0),
                f'{style_m.group(1)}{style_m.group(2)}{_append_nowrap(style_val)}{style_m.group(2)}',
                1,
            )
        else:
            new_attrs = new_attrs + ' style="white-space: nowrap;"'

        if mce_m is not None:
            new_attrs = new_attrs.replace(
                mce_m.group(0),
                f'{mce_m.group(1)}{mce_m.group(2)}{_append_nowrap(mce_val)}{mce_m.group(2)}',
                1,
            )
        return f'<p{new_attrs}>{body}</p>'

    return _NOWRAP_PARA_RE.sub(_replace, html_content)


def get_html_heading_levels(html_content: str):
    """返回 (existing_levels, max_level)"""
    if not html_content or not isinstance(html_content, str):
        return [], 0
    soup = BeautifulSoup(html_content, "html.parser")
    headings = soup.find_all(re.compile(r"^h[1-6]$", re.IGNORECASE))
    existing_levels = sorted({int(h.name[1]) for h in headings})
    max_level = max(existing_levels) if existing_levels else 0
    return existing_levels, max_level


def replace_first_heading_text(html_content: str, new_title: str) -> str:
    """将 HTML 中第一个 h1-h6 标签的文本内容替换为 new_title。
    若找不到任何标题标签，则原样返回。"""
    if not html_content or not new_title:
        return html_content
    soup = BeautifulSoup(html_content, "html.parser")
    heading = soup.find(re.compile(r"^h[1-6]$", re.IGNORECASE))
    if heading is None:
        return html_content
    heading.clear()
    heading.append(new_title)
    return str(soup)


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


# ─── HTML 固定宽度处理 ────────────────────────────────────────────────────────

def fix_html_to_fixed_width(html_content: str, width: int) -> str:
    """将 HTML 中的自适应宽度属性转为固定值，使 HTML 渲染不超过指定宽度，同时保留原始样式。

    处理范围：
    - <style> 块：展开媒体查询（仅保留目标宽度生效的规则），转换 %/vw 为 px
    - 内联 style 属性：转换宽度相关属性的 %/vw 为 px
    - HTML width 属性（table/img/td 等）：转换百分比为 px
    - 注入全局约束样式，确保内容不溢出指定宽度
    """
    soup = BeautifulSoup(html_content, "html.parser")

    for style_tag in soup.find_all("style"):
        if style_tag.string:
            style_tag.string = _css_fix_for_width(style_tag.string, width)

    for tag in soup.find_all(True):
        if tag.get("style"):
            tag["style"] = _inline_style_fix_for_width(tag["style"], width)
        if tag.get("width"):
            tag["width"] = _html_width_attr_to_px(tag["width"], width)

    # 注入全局约束样式（置于最后，优先级高）
    constraint_style = soup.new_tag("style")
    constraint_style.string = (
        f"html,body{{width:{width}px!important;max-width:{width}px!important;"
        f"overflow-x:hidden!important;box-sizing:border-box!important;}}"
        f"img{{max-width:100%!important;height:auto!important;}}"
        f"table{{max-width:100%!important;}}"
        f"*{{box-sizing:border-box!important;}}"
    )

    head_tag = soup.find("head")
    html_tag = soup.find("html")
    if head_tag:
        head_tag.append(constraint_style)
    elif html_tag:
        new_head = soup.new_tag("head")
        new_head.append(constraint_style)
        html_tag.insert(0, new_head)
    else:
        # 片段 HTML：包裹固定宽度容器
        wrapper = soup.new_tag(
            "div",
            style=(
                f"width:{width}px;max-width:{width}px;"
                f"overflow-x:hidden;box-sizing:border-box;"
            ),
        )
        for child in list(soup.children):
            wrapper.append(child.extract())
        soup.append(wrapper)

    return str(soup)


def _css_fix_for_width(css: str, width: int) -> str:
    """处理 CSS 文本：展开媒体查询，转换自适应宽度值为固定 px"""
    css = _flatten_media_queries(css, width)
    css = _convert_css_width_props(css, width)
    return css


def _flatten_media_queries(css: str, width: int) -> str:
    """展开媒体查询，保留在目标宽度下生效的规则，移除不生效的"""
    result_parts = []
    pos = 0
    media_re = re.compile(r'@media\b([^{]*)\{', re.IGNORECASE | re.DOTALL)

    while pos < len(css):
        m = media_re.search(css, pos)
        if not m:
            result_parts.append(css[pos:])
            break

        result_parts.append(css[pos:m.start()])
        condition = m.group(1).strip()
        block_start = m.end()

        depth = 1
        i = block_start
        while i < len(css) and depth > 0:
            if css[i] == "{":
                depth += 1
            elif css[i] == "}":
                depth -= 1
            i += 1

        block_content = css[block_start : i - 1]

        if _media_condition_applies(condition, width):
            result_parts.append(block_content)

        pos = i

    return "".join(result_parts)


def _media_condition_applies(condition: str, width: int) -> bool:
    """判断媒体查询条件在指定宽度下是否生效"""
    cond = condition.lower()

    if re.search(r'\bprint\b|\bspeech\b', cond):
        return False

    max_w = re.search(r'max-width\s*:\s*(\d+(?:\.\d+)?)(px|em|rem)?', cond)
    if max_w:
        val = float(max_w.group(1)) * (16 if (max_w.group(2) or "px") in ("em", "rem") else 1)
        if width > val:
            return False

    min_w = re.search(r'min-width\s*:\s*(\d+(?:\.\d+)?)(px|em|rem)?', cond)
    if min_w:
        val = float(min_w.group(1)) * (16 if (min_w.group(2) or "px") in ("em", "rem") else 1)
        if width < val:
            return False

    return True


_CSS_WIDTH_DECL_RE = re.compile(
    r'((?:^|(?<=[{;]))\s*(?:width|max-width|min-width|flex-basis)\s*:\s*)([^;}\n]+)',
    re.IGNORECASE | re.MULTILINE,
)


def _convert_css_width_props(css: str, width: int) -> str:
    """将 CSS 中宽度属性的 %/vw 值转为固定 px"""
    def replace_value(m):
        return m.group(1) + _units_to_px(m.group(2), width)

    return _CSS_WIDTH_DECL_RE.sub(replace_value, css)


def _inline_style_fix_for_width(style: str, width: int) -> str:
    """将内联 style 中宽度属性的 %/vw 值转为固定 px"""
    def replace_decl(m):
        return m.group(1) + m.group(2) + _units_to_px(m.group(3), width)

    return re.sub(
        r'((?:^|(?<=;))\s*(?:width|max-width|min-width|flex-basis)\s*)(:\s*)([^;]+)',
        replace_decl,
        style,
        flags=re.IGNORECASE,
    )


def _html_width_attr_to_px(value: str, width: int) -> str:
    """将 HTML width 属性的百分比值转为整数 px（不带单位）"""
    value = value.strip()
    m = re.match(r'^(\d+(?:\.\d+)?)%$', value)
    if m:
        return str(int(width * float(m.group(1)) / 100))
    return value


def _units_to_px(value: str, width: int) -> str:
    """将 CSS 值中的 % 和 vw 单位替换为 px"""
    def pct(m):
        return f"{int(width * float(m.group(1)) / 100)}px"

    def vw(m):
        return f"{int(width * float(m.group(1)) / 100)}px"

    value = re.sub(r'(\d+(?:\.\d+)?)%', pct, value)
    value = re.sub(r'(\d+(?:\.\d+)?)vw', vw, value)
    return value