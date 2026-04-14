import logging
from typing import Optional

from fastapi import APIRouter, Query

from app.core.logging_setup import log_file_path
from app.utils.file_utils import read_logs

router = APIRouter()
logger = logging.getLogger(__name__)


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
