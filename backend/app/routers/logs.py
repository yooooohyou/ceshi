import json
import logging
import time
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, Query, Request
from fastapi.responses import HTMLResponse, StreamingResponse

from app.core.logging_setup import log_file_path
from app.utils.file_utils import read_logs
from app.utils.log_tailer import tail_file

router = APIRouter()
logger = logging.getLogger(__name__)

_VIEWER_HTML_PATH = Path(__file__).resolve().parent.parent / "templates" / "logs_viewer.html"


@router.get("/api/logs", summary="查询日志")
async def get_logs(
    level: Optional[str] = Query(None, description="日志级别（INFO/WARNING/ERROR）"),
    keyword: Optional[str] = Query(None, description="日志关键词"),
    limit: int = Query(100, ge=1, le=1000, description="返回日志条数（1-1000）"),
):
    logs = read_logs(log_file_path, level, keyword, limit)
    return {
        "log_file": log_file_path,
        "total": len(logs),
        "limit": limit,
        "logs": logs,
    }


@router.get("/logs/viewer", summary="日志查看页面（HTML）", response_class=HTMLResponse)
async def logs_viewer():
    html = _VIEWER_HTML_PATH.read_text(encoding="utf-8")
    return HTMLResponse(content=html)


@router.get("/api/logs/stream", summary="实时日志流（SSE）")
async def logs_stream(
    request: Request,
    tail: int = Query(200, ge=0, le=2000, description="初始历史日志条数"),
):
    async def event_stream():
        # 1) 历史日志
        if tail > 0:
            for line in read_logs(log_file_path, None, None, tail):
                yield f"data: {json.dumps({'line': line}, ensure_ascii=False)}\n\n"
        yield "event: ready\ndata: {}\n\n"

        # 2) 实时追加
        last_ping = time.monotonic()
        async for line in tail_file(log_file_path, from_end=True):
            if await request.is_disconnected():
                return
            yield f"data: {json.dumps({'line': line}, ensure_ascii=False)}\n\n"
            now = time.monotonic()
            if now - last_ping > 15:
                yield ": ping\n\n"
                last_ping = now

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )
