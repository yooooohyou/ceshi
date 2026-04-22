"""
嵌入组件标记工具（Embed Marker）

用途：
    1. 将任意 JSON 数据生成一个带 URL 的 HTML 占位符（可内联/块级）
    2. 支持从 HTML 反向解析出所有标记
    3. HTML 转 DOCX 后，在 DOCX 中可通过 embed_id 文本定位并替换为真实组件
    4. payload 以 JSON 存储（数据库列类型 JSONB），天然支持扩展
       （text / table / image / chart / reference / custom ...）

设计要点：
    - 容器用 <span>/<div>，允许嵌套复合内容（图片、表格等）
    - 可见文本中包含 "【EMB_xxxxxxxx】" 字面量，即便 data-* 属性在
      DOCX 转换中被剥离，仍可通过该字面量在 DOCX 中定位（与分节符
      /分页符占位符 SB_MARKER_XXX / PB_MARKER_XXX 相同思路）
    - data-embed-type 决定由哪个渲染器处理
    - data-embed-version 用于 payload 结构演化时做兼容判断
"""

from __future__ import annotations

import html
import json
import logging
import re
import time
import uuid
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional

logger = logging.getLogger(__name__)


# ─── 常量 ─────────────────────────────────────────────────────────────────────

EMBED_ID_PREFIX = "EMB_"

# 可见占位文本模板；【】让肉眼容易识别，\u2068/\u2069 隔断双向文本（可选，防止复杂脚本混排）
VISIBLE_PLACEHOLDER_FMT = "【{embed_id}】"

# 支持的显示模式
DISPLAY_INLINE = "inline"
DISPLAY_BLOCK = "block"

# 内置组件类型（可继续扩展）
TYPE_TEXT = "text"
TYPE_TABLE = "table"
TYPE_IMAGE = "image"
TYPE_CHART = "chart"
TYPE_LINK = "link"
TYPE_REFERENCE = "reference"  # 引用别处内容
TYPE_CUSTOM = "custom"

BLOCK_TYPES = {TYPE_TABLE, TYPE_IMAGE, TYPE_CHART}

# HTML 端统一 class 名
CLASS_BASE = "yxdl-embed"
CLASS_INLINE = "yxdl-embed-inline"
CLASS_BLOCK = "yxdl-embed-block"

# embed_id 识别正则（允许 8~32 位 hex）
EMBED_ID_RE = re.compile(rf"{EMBED_ID_PREFIX}[0-9A-Fa-f]{{8,32}}")

# 匹配占位符整体（<span>/<div> 开闭，含 data-embed-id）
EMBED_HTML_RE = re.compile(
    r'<(?P<tag>span|div)\b(?P<attrs>[^>]*\bdata-embed-id="(?P<id>EMB_[0-9A-Fa-f]+)"[^>]*)>'
    r'(?P<inner>.*?)'
    r'</(?P=tag)>',
    re.IGNORECASE | re.DOTALL,
)


# ─── 渲染器注册（可选的扩展点） ────────────────────────────────────────────────

# HTML 渲染器：根据 payload 生成占位符"内部可见 HTML"
#   fn(spec) -> str
_HTML_RENDERERS: Dict[str, Callable[["EmbedSpec"], str]] = {}

# DOCX 渲染器：在 python-docx 段落中将占位文本替换为实际组件
#   fn(spec, paragraph, document) -> None
_DOCX_RENDERERS: Dict[str, Callable[["EmbedSpec", Any, Any], None]] = {}


def register_html_renderer(embed_type: str, fn: Callable[["EmbedSpec"], str]) -> None:
    _HTML_RENDERERS[embed_type] = fn


def register_docx_renderer(embed_type: str, fn: Callable[["EmbedSpec", Any, Any], None]) -> None:
    _DOCX_RENDERERS[embed_type] = fn


# ─── 数据模型 ─────────────────────────────────────────────────────────────────

@dataclass
class EmbedSpec:
    """一个嵌入组件的完整规格。可序列化为 JSON 存库。"""

    embed_id: str
    embed_type: str
    url: str
    payload: Dict[str, Any] = field(default_factory=dict)
    version: int = 1
    title: str = ""
    display: str = DISPLAY_INLINE  # inline / block
    record_id: Optional[int] = None
    node_id: Optional[int] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "embed_id":   self.embed_id,
            "embed_type": self.embed_type,
            "url":        self.url,
            "payload":    self.payload,
            "version":    self.version,
            "title":      self.title,
            "display":    self.display,
            "record_id":  self.record_id,
            "node_id":    self.node_id,
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "EmbedSpec":
        return cls(
            embed_id=d["embed_id"],
            embed_type=d["embed_type"],
            url=d.get("url", ""),
            payload=d.get("payload") or {},
            version=int(d.get("version") or 1),
            title=d.get("title") or "",
            display=d.get("display") or DISPLAY_INLINE,
            record_id=d.get("record_id"),
            node_id=d.get("node_id"),
        )


# ─── 内部工具 ─────────────────────────────────────────────────────────────────

def _generate_embed_id() -> str:
    """生成唯一 embed_id（12 位 hex，前缀 EMB_），碰撞概率 ~2^-48，远超业务需要。"""
    return f"{EMBED_ID_PREFIX}{uuid.uuid4().hex[:12].upper()}"


def _is_block(embed_type: str, display: Optional[str]) -> bool:
    if display == DISPLAY_BLOCK:
        return True
    if display == DISPLAY_INLINE:
        return False
    return embed_type in BLOCK_TYPES


def _default_html_renderer(spec: EmbedSpec) -> str:
    """未注册时的兜底渲染：不输出额外内容；外层 anchor 已承载标题和跳转。"""
    return ""


# ─── 核心对外函数 ─────────────────────────────────────────────────────────────

def build_embed_marker(
    data: Dict[str, Any],
    embed_type: str = TYPE_CUSTOM,
    url: Optional[str] = None,
    *,
    title: str = "",
    version: int = 1,
    display: Optional[str] = None,
    record_id: Optional[int] = None,
    node_id: Optional[int] = None,
    embed_id: Optional[str] = None,
    url_builder: Optional[Callable[[str], str]] = None,
) -> tuple[EmbedSpec, str]:
    """
    将 JSON `data` 生成可嵌入 HTML 的占位标记。

    Args:
        data:         组件的业务数据（任意 JSON 可序列化结构）
        embed_type:   组件类型（text/table/image/chart/reference/custom ...）
        url:          显式指定跳转 URL；为 None 时由 url_builder 依据 embed_id 生成；
                      两者都没给则使用默认 `/doc_editor/embeds/{embed_id}`
        title:        用于 HTML 可见文本、DOCX 兜底标题
        version:      payload 结构版本
        display:      inline / block；不传则按 embed_type 自动判定
        record_id:    关联的文档 record_id（可选）
        node_id:      关联的标题树节点 id（可选）
        embed_id:     自定义 embed_id（一般用于"由历史记录再生成标记"的场景）
        url_builder:  自定义 URL 生成器，签名 (embed_id) -> str

    Returns:
        (spec, html_snippet)
        spec          —— 可直接 json.dumps 入库的 EmbedSpec
        html_snippet  —— 可插入 HTML 编辑器的片段

    注意：
        - html_snippet 中的可见文本包含 "【EMB_xxxx】"，用于 DOCX 端定位；
          前端如需隐藏可用 CSS：`.yxdl-embed [data-embed-anchor]::before`
          或服务端在给前端之前做一次文本净化。
        - data 不会写入 data-* 属性，避免 HTML 超长；数据走数据库。
    """
    eid = embed_id or _generate_embed_id()
    if not EMBED_ID_RE.fullmatch(eid):
        raise ValueError(f"非法 embed_id：{eid}")

    if url is None:
        url = url_builder(eid) if url_builder else f"/doc_editor/embeds/{eid}"

    is_block = _is_block(embed_type, display)
    final_display = DISPLAY_BLOCK if is_block else DISPLAY_INLINE

    spec = EmbedSpec(
        embed_id=eid,
        embed_type=embed_type,
        url=url,
        payload=data or {},
        version=version,
        title=title,
        display=final_display,
        record_id=record_id,
        node_id=node_id,
    )

    tag = "div" if is_block else "span"
    cls = f"{CLASS_BASE} {CLASS_BLOCK if is_block else CLASS_INLINE}"

    renderer = _HTML_RENDERERS.get(embed_type, _default_html_renderer)
    inner_html = renderer(spec)

    visible_tag = VISIBLE_PLACEHOLDER_FMT.format(embed_id=eid)

    # 可见内容结构：
    #   <a href=URL target=_blank data-embed-anchor>【EMB_xxx】title</a><inner_html>
    # 锚点 <a> 里必须包含 embed_id 字面量，DOCX 转换后靠这个文本定位
    anchor_text = f"{visible_tag}{html.escape(title)}" if title else visible_tag
    anchor = (
        f'<a href="{html.escape(url)}" target="_blank" '
        f'data-embed-anchor="1">{anchor_text}</a>'
    )

    attrs = (
        f'class="{cls}" '
        f'contenteditable="false" '
        f'data-embed-id="{eid}" '
        f'data-embed-type="{html.escape(embed_type)}" '
        f'data-embed-version="{version}" '
        f'data-embed-display="{final_display}" '
        f'data-embed-url="{html.escape(url)}"'
    )
    if title:
        attrs += f' data-embed-title="{html.escape(title)}"'

    snippet = f"<{tag} {attrs}>{anchor}{inner_html}</{tag}>"
    return spec, snippet


# ─── 反向解析（HTML 扫描） ────────────────────────────────────────────────────

def parse_embed_markers_from_html(html_text: str) -> List[Dict[str, Any]]:
    """
    从 HTML 扫描所有嵌入组件标记，返回元信息列表：
        [{embed_id, embed_type, version, url, display, title, span: (start,end)}, ...]
    不读 payload（payload 只在数据库），只做轻量解析，便于统计/替换。
    """
    results: List[Dict[str, Any]] = []
    for m in EMBED_HTML_RE.finditer(html_text):
        attrs = m.group("attrs")

        def _pick(name: str) -> str:
            mm = re.search(rf'\b{name}="([^"]*)"', attrs, re.IGNORECASE)
            return mm.group(1) if mm else ""

        results.append({
            "embed_id":    m.group("id"),
            "embed_type":  _pick("data-embed-type"),
            "version":     int(_pick("data-embed-version") or 1),
            "url":         _pick("data-embed-url"),
            "display":     _pick("data-embed-display") or DISPLAY_INLINE,
            "title":       _pick("data-embed-title"),
            "span":        (m.start(), m.end()),
        })
    return results


def strip_visible_placeholders(html_text: str) -> str:
    """把 HTML 可见文本里的【EMB_xxx】字面量清理掉（仅用于最终呈现给用户）。"""
    return re.sub(rf"【{EMBED_ID_PREFIX}[0-9A-Fa-f]{{8,32}}】", "", html_text)


# ─── DOCX 端占位符定位 ────────────────────────────────────────────────────────

def find_embed_ids_in_docx_text(text: str) -> List[str]:
    """在 DOCX 段落/表格单元格的纯文本里查找所有 embed_id。"""
    return EMBED_ID_RE.findall(text or "")


def collect_docx_embed_ids(docx_document) -> set:
    """
    快速扫描 Document 段落，返回文件中实际出现的 embed_id 集合（不加载任何 spec）。
    用于在查询 DB 前预判哪些 embed 占位符在文档中确实存在，避免无效的 DB 查询和反序列化。
    """
    ids: set = set()
    prefix = EMBED_ID_PREFIX
    for para in docx_document.paragraphs:
        text = para.text or ""
        if prefix in text:
            ids.update(find_embed_ids_in_docx_text(text))
    return ids


def build_docx_replace_plan(
    docx_document,
    specs_by_id: Dict[str, EmbedSpec],
) -> List[Dict[str, Any]]:
    """
    扫描 python-docx 的 Document，找到所有带 EMB_xxx 字面量的段落，
    生成替换计划：[{paragraph, embed_id, spec, matched_text}, ...]

    调用方拿到计划后，按 spec.embed_type 路由到对应 DOCX 渲染器执行替换。

    性能优化：
        - 先用一个极简的 `EMBED_ID_PREFIX in text` 前缀检查跳过绝大多数段落，
          避免对每个段落都跑正则
        - 每 500 个段落输出一次扫描进度，便于排查大文档慢扫描
    """
    plan: List[Dict[str, Any]] = []
    t0 = time.perf_counter()
    paragraphs = docx_document.paragraphs
    total = len(paragraphs)
    scanned = 0
    matched_paragraphs = 0
    prefix = EMBED_ID_PREFIX

    for para in paragraphs:
        scanned += 1
        text = para.text or ""
        # 快速预过滤：绝大多数段落不含 EMB_ 前缀，直接跳过，省掉正则开销
        if prefix not in text:
            if scanned % 500 == 0:
                logger.info(
                    "build_docx_replace_plan 扫描中 %d/%d 段落 plan=%d 耗时=%.2fs",
                    scanned, total, len(plan), time.perf_counter() - t0,
                )
            continue
        matched_paragraphs += 1
        for eid in set(find_embed_ids_in_docx_text(text)):
            spec = specs_by_id.get(eid)
            if not spec:
                logger.warning(f"build_docx_replace_plan: 未知 embed_id={eid}，已跳过")
                continue
            plan.append({
                "paragraph":     para,
                "embed_id":      eid,
                "spec":          spec,
                "matched_text":  VISIBLE_PLACEHOLDER_FMT.format(embed_id=eid),
            })
        if scanned % 500 == 0:
            logger.info(
                "build_docx_replace_plan 扫描中 %d/%d 段落 plan=%d 耗时=%.2fs",
                scanned, total, len(plan), time.perf_counter() - t0,
            )

    logger.info(
        "build_docx_replace_plan 扫描完成 段落=%d 命中段=%d 计划项=%d 耗时=%.2fs",
        total, matched_paragraphs, len(plan), time.perf_counter() - t0,
    )
    return plan


def render_docx_replace_plan(plan: List[Dict[str, Any]], docx_document) -> int:
    """
    按计划逐项调用已注册的 DOCX 渲染器；返回成功处理数。

    性能/可观测性：
        - 每渲染完一项就打印进度（含 embed_id、类型、单项耗时、累计耗时、ETA），
          便于在日志里直接看到卡在了哪个表格上
    """
    ok = 0
    total = len(plan)
    t0 = time.perf_counter()
    logger.info("render_docx_replace_plan 开始 共 %d 项 embed", total)

    for i, item in enumerate(plan, 1):
        spec: EmbedSpec = item["spec"]
        renderer = _DOCX_RENDERERS.get(spec.embed_type)
        if renderer is None:
            logger.warning(f"render_docx_replace_plan: 类型 {spec.embed_type} 无 DOCX 渲染器")
            continue
        t_item = time.perf_counter()
        try:
            renderer(spec, item["paragraph"], docx_document)
            ok += 1
        except Exception as e:
            logger.error(f"DOCX 渲染失败 embed_id={spec.embed_id}：{e}")
            continue

        elapsed = time.perf_counter() - t0
        item_cost = time.perf_counter() - t_item
        eta = (elapsed / i) * (total - i) if i > 0 else 0.0
        logger.info(
            "render_docx_replace_plan 进度 %d/%d embed_id=%s type=%s"
            " 单项=%.2fs 累计=%.2fs ETA≈%.2fs",
            i, total, spec.embed_id, spec.embed_type, item_cost, elapsed, eta,
        )

    logger.info(
        "render_docx_replace_plan 结束 成功=%d/%d 总耗时=%.2fs",
        ok, total, time.perf_counter() - t0,
    )
    return ok


# ─── 与数据库对接的序列化辅助 ─────────────────────────────────────────────────

def spec_to_db_row(spec: EmbedSpec) -> Dict[str, Any]:
    """EmbedSpec → 入库字段（payload 保持 dict，让 psycopg2 自动转 JSONB）。"""
    return {
        "embed_id":   spec.embed_id,
        "embed_type": spec.embed_type,
        "url":        spec.url,
        "payload":    json.dumps(spec.payload, ensure_ascii=False),
        "version":    spec.version,
        "title":      spec.title,
        "display":    spec.display,
        "record_id":  spec.record_id,
        "node_id":    spec.node_id,
    }


def spec_from_db_row(row: Dict[str, Any]) -> EmbedSpec:
    """入库行 → EmbedSpec（自动反序列化 payload）。"""
    payload = row.get("payload")
    if isinstance(payload, str):
        try:
            payload = json.loads(payload)
        except json.JSONDecodeError:
            payload = {}
    return EmbedSpec(
        embed_id=row["embed_id"],
        embed_type=row["embed_type"],
        url=row.get("url") or "",
        payload=payload or {},
        version=int(row.get("version") or 1),
        title=row.get("title") or "",
        display=row.get("display") or DISPLAY_INLINE,
        record_id=row.get("record_id"),
        node_id=row.get("node_id"),
    )