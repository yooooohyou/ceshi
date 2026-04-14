import logging
import pathlib
import urllib.parse
from typing import Optional, Union

import requests
from fastapi import APIRouter, Body, Request
from fastapi.responses import JSONResponse, StreamingResponse

from app.converters.docx_converter import convert_html_to_docx
from app.core.config import UPLOAD_DIR
from app.models.schemas import unified_response
from app.utils.file_utils import generate_unique_filename
from file_resp import FileResp

router = APIRouter()
logger = logging.getLogger(__name__)


@router.post("/html_to_docx", summary="HTML转DOCX文件流", response_model=None)
async def html_to_docx_api(
    html_content: str = Body(..., description="需要转换的HTML文本"),
    filename: Optional[str] = Body("output.docx", description="下载的DOCX文件名"),
) -> Union[JSONResponse, StreamingResponse]:
    """接收 HTML 文本，生成并返回 DOCX 文件流"""
    import datetime
    import os
    try:
        if not html_content.strip():
            return unified_response(400, "HTML内容不能为空")

        if not filename.lower().endswith(".docx"):
            filename += ".docx"
        filename = os.path.basename(filename).replace("/", "_").replace("\\", "_").replace(":", "_")
        current_time = datetime.datetime.now()

        success, result, path_ = convert_html_to_docx(html_content)
        if not success:
            return unified_response(500, result)

        encoded_filename = urllib.parse.quote(filename)
        headers = {
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}; filename={encoded_filename}",
            "Access-Control-Expose-Headers": "Content-Disposition",
            "X-Update-Time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        response = StreamingResponse(
            result,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers,
        )
        response.status_code = 200
        return response

    except Exception as e:
        return unified_response(500, f"HTML转DOCX失败：{str(e)}")


@router.post("/file_slicing_download", summary="文件分片下载", response_model=None)
async def file_slicing_download_api(
    request: Request,
    file_path: str = Body(..., description="文件的完整URL路径"),
    filename: str = Body(..., description="文件名"),
) -> Union[JSONResponse, StreamingResponse]:
    """接收文件路径，分片返回文件流（支持 Range 下载）"""
    import os
    try:
        response = requests.get(file_path, timeout=30)
        response.raise_for_status()
        temp_docx_filename = generate_unique_filename("temp.docx")
        abs_file_path = os.path.join(UPLOAD_DIR, temp_docx_filename)
        with open(abs_file_path, "wb") as f:
            f.write(response.content)
        return FileResp(request, pathlib.Path(abs_file_path)).start()
    except Exception as e:
        return unified_response(500, f"文件分片下载失败：{str(e)}")
