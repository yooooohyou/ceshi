import logging
import os
import subprocess
import tempfile

from fastapi import APIRouter, BackgroundTasks, File, HTTPException, UploadFile

from app.core.config import UPLOAD_DIR, get_server_uploads_config
from app.core.constants import LIBREOFFICE_PATH
from app.models.schemas import unified_response
from app.utils.file_utils import cleanup_temp_files
from fastapi.responses import FileResponse

router = APIRouter()
logger = logging.getLogger(__name__)


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
