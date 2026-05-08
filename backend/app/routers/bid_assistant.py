import logging
import os
import shutil
from urllib.parse import quote

from fastapi import APIRouter, Body
from fastapi.responses import JSONResponse

from app.core.config import UPLOAD_DIR, get_server_uploads_config

router = APIRouter()
logger = logging.getLogger(__name__)

# 源静态文件目录（backend/static）
STATIC_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "static")
)


def _get_uploads_web_url() -> str:
    """获取 uploads 的 web 访问根 URL，例如 http://10.13.6.180:21001/doc_editor/uploads/。"""
    try:
        cfg = get_server_uploads_config()
        front = (cfg.get("web_front_path") or "").strip()
        if front:
            return front if front.endswith("/") else front + "/"
    except Exception as e:
        logger.warning(f"读取 web_front_path 失败：{e}")
    return ""


@router.post(
    "/bid_assistant/company_qualification/template_file_upload1",
    summary="公司资质模板文件上传（基于 static 静态文件返回 uploads 下载 URL）",
)
async def template_file_upload1(
    file_name: str = Body(..., embed=True, description="static 目录下的文件名"),
):
    """根据传入的 file_name，将 static 下源文件拷贝到 uploads 目录，返回可下载的 web URL。"""
    try:
        if not file_name:
            return JSONResponse(content={
                "status": 1,
                "is_finish": 0,
                "msg": "file_name 不能为空",
                "data": "",
            })

        src_path = os.path.join(STATIC_DIR, file_name)
        if not os.path.isfile(src_path):
            return JSONResponse(content={
                "status": 1,
                "is_finish": 0,
                "msg": f"静态文件不存在：static/{file_name}",
                "data": "",
            })

        web_root = _get_uploads_web_url()
        if not web_root:
            return JSONResponse(content={
                "status": 1,
                "is_finish": 0,
                "msg": "未配置 web_front_path，无法生成下载地址",
                "data": "",
            })

        # 同步拷贝到 uploads 目录（保持原文件名），确保通过 uploads URL 可下载
        try:
            os.makedirs(UPLOAD_DIR, exist_ok=True)
            dst_path = os.path.join(UPLOAD_DIR, file_name)
            if not os.path.isfile(dst_path) or os.path.getmtime(dst_path) < os.path.getmtime(src_path):
                shutil.copyfile(src_path, dst_path)
        except Exception as e:
            logger.exception("拷贝文件到 uploads 失败")
            return JSONResponse(content={
                "status": 1,
                "is_finish": 0,
                "msg": f"拷贝到 uploads 失败：{e}",
                "data": "",
            })

        download_url = f"{web_root}{quote(file_name)}"

        return JSONResponse(content={
            "status": 0,
            "is_finish": 1,
            "msg": "成功",
            "data": download_url,
        })

    except Exception as e:
        logger.exception("template_file_upload1 处理失败")
        return JSONResponse(status_code=500, content={
            "status": 1,
            "is_finish": 0,
            "msg": f"接口异常：{str(e)}",
            "data": "",
        })
