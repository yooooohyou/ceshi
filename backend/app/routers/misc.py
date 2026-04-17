import logging
import os
import subprocess
import tempfile

from fastapi import APIRouter, BackgroundTasks, Body, File, HTTPException, UploadFile

from app.core.config import UPLOAD_DIR, get_server_uploads_config
from app.core.constants import LIBREOFFICE_PATH
from app.models.schemas import unified_response
from app.utils.file_utils import cleanup_temp_files
from app.utils.html_utils import fix_html_to_fixed_width
from fastapi.responses import FileResponse

router = APIRouter()
logger = logging.getLogger(__name__)


@router.post("/fix-html-width", summary="HTML 固定宽度格式化")
async def fix_html_width_api(
    html_content: str = Body(..., description="需要处理的 HTML 文本"),
    width: int = Body(..., description="目标固定宽度（像素），HTML 渲染不会超过此宽度", gt=0),
):
    """将 HTML 中的自适应宽度属性（百分比、vw 单位、媒体查询等）转换为固定像素值，
    使 HTML 在指定宽度下渲染时不超宽，同时尽可能保留原始视觉样式。

    处理内容：
    - `<style>` 中的媒体查询展开（保留目标宽度适用的规则）
    - CSS 宽度属性（width/max-width/min-width/flex-basis）的 %/vw 转 px
    - 元素内联 `style` 属性中的宽度 %/vw 转 px
    - HTML `width` 属性（table/img/td 等）的百分比转 px
    - 注入全局约束样式确保内容不溢出
    """
    try:
        if not html_content.strip():
            return unified_response(400, "html_content 不能为空")
        result_html = fix_html_to_fixed_width(html_content, width)
        return unified_response(0, "处理成功", {"html": result_html})
    except Exception as e:
        logger.exception("fix_html_width 处理失败")
        raise HTTPException(status_code=500, detail=f"处理失败：{str(e)}")


@router.get("/health", summary="接口健康检查")
async def health_check():
    """检查接口是否可用，同时验证 LibreOffice 是否能调用"""
    try:
        result = subprocess.run(
            [LIBREOFFICE_PATH, "--version"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        if result.returncode != 0:
            return {"status": "unhealthy", "message": "LibreOffice 调用失败", "error": result.stderr}
        return {"status": "healthy", "message": "接口和LibreOffice均正常", "libreoffice_version": result.stdout.strip()}
    except Exception as e:
        return {"status": "unhealthy", "message": "接口异常", "error": str(e)}


@router.get("/test_use_config", summary="测试在业务逻辑中使用配置")
async def test_use_config():
    """示例：在实际业务逻辑中读取并使用上传路径配置"""
    try:
        uploads_config = get_server_uploads_config()
        local_path = uploads_config["user_local_path"]
        web_path = uploads_config["web_backend_path"]
        filename = "test_file.pdf"
        return {
            "local_file_path": os.path.join(local_path, filename),
            "web_file_url": web_path + filename,
            "original_config": uploads_config,
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/test-liboffice/emf-to-png", summary="测试LibreOffice EMF转PNG功能")
async def test_libreoffice_emf2png(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="上传EMF格式文件"),
):
    """测试 LibreOffice 是否能正常将 EMF 文件转换为 PNG"""
    temp_dir = tempfile.mkdtemp(prefix="libreoffice_test_")
    try:
        if not file.filename.lower().endswith(".emf"):
            raise HTTPException(status_code=400, detail="仅支持上传 .emf 格式文件")

        emf_filename = os.path.basename(file.filename)
        emf_file_path = os.path.join(temp_dir, emf_filename)
        file_content = await file.read()
        with open(emf_file_path, "wb") as f:
            f.write(file_content)

        png_filename = os.path.splitext(emf_filename)[0] + ".png"
        png_file_path = os.path.join(temp_dir, png_filename)

        cmd = [
            LIBREOFFICE_PATH, "--headless", "--norestore", "--nolockcheck",
            "--convert-to", 'png:draw_png_Export:{"Translucent":true,"Resolution":300}',
            emf_file_path, "--outdir", temp_dir,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

        if result.returncode != 0:
            raise HTTPException(
                status_code=500,
                detail=f"LibreOffice 转换失败：\n标准输出：{result.stdout}\n错误输出：{result.stderr}",
            )
        if not os.path.exists(png_file_path):
            raise HTTPException(status_code=500, detail="转换命令执行成功，但未生成PNG文件")

        background_tasks.add_task(cleanup_temp_files, temp_dir=temp_dir)
        return FileResponse(path=png_file_path, filename=png_filename, media_type="image/png")

    except subprocess.TimeoutExpired:
        cleanup_temp_files(temp_dir)
        raise HTTPException(status_code=500, detail="LibreOffice 转换超时（60秒）")
    except Exception as e:
        cleanup_temp_files(temp_dir)
        raise HTTPException(status_code=500, detail=f"转换过程出错：{str(e)}")
