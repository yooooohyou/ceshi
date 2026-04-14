# ─── 日志初始化（必须最先执行） ──────────────────────────────────────────────
from app.core.logging_setup import log_file_path, logger  # noqa: E402

# ─── 标准库 / 第三方库 ────────────────────────────────────────────────────────
import logging

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

# ─── 应用配置 ─────────────────────────────────────────────────────────────────
from app.core.config import STATIC_WEB_PREFIX, UPLOAD_DIR, system_path
from app.middleware.http_logging import http_log_middleware
from app.models.schemas import unified_response

# ─── FastAPI 应用实例 ─────────────────────────────────────────────────────────
app = FastAPI(title="DOCX文件上传&HTML转换接口", version="1.0")

# ─── CORS ────────────────────────────────────────────────────────────────────
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── 静态文件挂载（仅 Windows 开发环境） ──────────────────────────────────────
if system_path == "Windows":
    app.mount(STATIC_WEB_PREFIX, StaticFiles(directory=UPLOAD_DIR), name="uploads")

# ─── HTTP 日志中间件 ──────────────────────────────────────────────────────────
app.middleware("http")(http_log_middleware)

# ─── 全局异常处理 ─────────────────────────────────────────────────────────────

@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    return unified_response(400, "参数校验失败", {"errors": exc.errors()})


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    if hasattr(exc, "status_code") and hasattr(exc, "detail"):
        return unified_response(exc.status_code, exc.detail)
    return unified_response(500, f"服务器内部错误：{str(exc)}")


# ─── 路由注册 ─────────────────────────────────────────────────────────────────
from app.routers import conversion, document, generation, logs, misc, upload  # noqa: E402

app.include_router(upload.router,     prefix="/doc_editor", tags=["上传"])
app.include_router(document.router,   prefix="/doc_editor", tags=["文档管理"])
app.include_router(conversion.router, prefix="/doc_editor", tags=["转换"])
app.include_router(generation.router, prefix="/doc_editor", tags=["生成器"])
app.include_router(logs.router,       tags=["日志"])
app.include_router(misc.router,       tags=["其他"])

# ─── 启动入口 ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        app="main:app",
        host="0.0.0.0",
        port=8080,
        reload=True,
        workers=4,
        timeout_keep_alive=12000,
    )
