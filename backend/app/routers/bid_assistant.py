import logging
import os
from urllib.parse import quote

from fastapi import APIRouter, Body, Request
from fastapi.responses import JSONResponse

router = APIRouter()
logger = logging.getLogger(__name__)

# 静态文件目录（backend/static）
_STATIC_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "static")
)


@router.post(
    "/bid_assistant/company_qualification/template_file_upload1",
    summary="公司资质模板文件上传（基于 static 静态文件返回水印 URL）",
)
async def template_file_upload1(
    request: Request,
    file_name: str = Body(..., embed=True, description="static 目录下的文件名"),
):
    """根据传入的 file_name，从 static 目录读取对应文件，返回 /watermark/<file_name> 形式的 URL。"""
    try:
        if not file_name:
            return JSONResponse(content={
                "status": 1,
                "is_finish": 0,
                "msg": "file_name 不能为空",
                "data": "",
            })

        static_file_path = os.path.join(_STATIC_DIR, file_name)
        if not os.path.exists(static_file_path):
            return JSONResponse(content={
                "status": 1,
                "is_finish": 0,
                "msg": f"静态文件不存在：static/{file_name}",
                "data": "",
            })

        base_url = f"{request.url.scheme}://{request.url.netloc}"
        watermark_url = f"{base_url}/watermark/{quote(file_name)}"

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
