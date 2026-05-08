import logging
import os
from urllib.parse import quote, urlparse

from fastapi import APIRouter, Body, HTTPException
from fastapi.responses import FileResponse, JSONResponse

from app.core.config import get_server_uploads_config

router = APIRouter()
# 用于无 /doc_editor 前缀的下载路由
watermark_router = APIRouter()
logger = logging.getLogger(__name__)

# 静态文件目录（backend/static）
STATIC_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "static")
)
WATERMARK_URL_PREFIX = "/watermark"


def _get_web_base_url() -> str:
    """从配置 web_front_path 提取可访问下载的 scheme://host 部分。"""
    try:
        cfg = get_server_uploads_config()
        front = (cfg.get("web_front_path") or "").strip()
        if front:
            parsed = urlparse(front)
            if parsed.scheme and parsed.netloc:
                return f"{parsed.scheme}://{parsed.netloc}"
    except Exception as e:
        logger.warning(f"读取 web_front_path 失败：{e}")
    return ""


@router.post(
    "/bid_assistant/company_qualification/template_file_upload1",
    summary="公司资质模板文件上传（基于 static 静态文件返回水印 URL）",
)
async def template_file_upload1(
    file_name: str = Body(..., embed=True, description="static 目录下的文件名"),
):
    """根据传入的 file_name，从 static 目录读取对应文件，返回 /watermark/<file_name> 可下载 URL。"""
    try:
        if not file_name:
            return JSONResponse(content={
                "status": 1,
                "is_finish": 0,
                "msg": "file_name 不能为空",
                "data": "",
            })

        static_file_path = os.path.join(STATIC_DIR, file_name)
        if not os.path.exists(static_file_path):
            return JSONResponse(content={
                "status": 1,
                "is_finish": 0,
                "msg": f"静态文件不存在：static/{file_name}",
                "data": "",
            })

        base_url = _get_web_base_url()
        if not base_url:
            return JSONResponse(content={
                "status": 1,
                "is_finish": 0,
                "msg": "未配置 web_front_path，无法生成下载地址",
                "data": "",
            })

        watermark_url = f"{base_url}{WATERMARK_URL_PREFIX}/{quote(file_name)}"

        return JSONResponse(content={
            "status": 0,
            "is_finish": 1,
            "msg": "成功",
            "data": watermark_url,
        })

    except Exception as e:
        logger.exception("template_file_upload1 处理失败")
        return JSONResponse(status_code=500, content={
            "status": 1,
            "is_finish": 0,
            "msg": f"接口异常：{str(e)}",
            "data": "",
        })


# ─── /watermark/<file_name> 直接下载源文件（不做任何压缩/转换） ─────────────
@watermark_router.get(
    f"{WATERMARK_URL_PREFIX}/{{file_name:path}}",
    summary="下载 static 下的源文件（原样字节流）",
    include_in_schema=False,
)
async def watermark_download(file_name: str):
    safe_name = os.path.basename(file_name)
    if not safe_name or safe_name in {".", ".."}:
        raise HTTPException(status_code=400, detail="非法的文件名")

    abs_path = os.path.join(STATIC_DIR, safe_name)
    if not os.path.isfile(abs_path):
        raise HTTPException(status_code=404, detail=f"文件不存在：{safe_name}")

    return FileResponse(
        path=abs_path,
        filename=safe_name,
        media_type="application/octet-stream",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{quote(safe_name)}",
            "Content-Encoding": "identity",
            "Cache-Control": "no-transform",
        },
    )

