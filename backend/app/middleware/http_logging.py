import time
import logging
from fastapi import Request
from fastapi.responses import Response

logger = logging.getLogger(__name__)

_LOG_SKIP_PATHS = {"/api/logs", "/api/logs/stream", "/logs/viewer", "/health", "/docs", "/redoc", "/openapi.json"}


async def http_log_middleware(request: Request, call_next):
    """记录所有 HTTP 请求/响应（跳过日志、静态、健康检查路径）"""
    if request.url.path in _LOG_SKIP_PATHS or request.url.path.startswith("/uploads"):
        return await call_next(request)

    start_time = time.time()
    path = request.url.path
    method = request.method
    query = str(request.query_params) if request.query_params else ""

    content_type = request.headers.get("content-type", "")
    if "multipart/form-data" in content_type:
        logger.info("[REQ] %s %s query=%s body=<multipart/form-data>", method, path, query)
    elif "application/json" in content_type:
        try:
            body_bytes = await request.body()
            sample = body_bytes[:200]
            non_printable = sum(1 for b in sample if b < 0x09 or (0x0e <= b <= 0x1f) or b == 0x7f)
            if sample and non_printable / len(sample) > 0.1:
                logger.info("[REQ] %s %s query=%s body=<binary %d bytes>", method, path, query, len(body_bytes))
            else:
                body_str = body_bytes.decode("utf-8", errors="replace")
                if len(body_str) > 1000:
                    body_str = body_str[:1000] + "...(truncated)"
                logger.info("[REQ] %s %s query=%s body=%s", method, path, query, body_str)

            async def receive():
                return {"type": "http.request", "body": body_bytes, "more_body": False}

            request = Request(request.scope, receive)
        except Exception:
            logger.info("[REQ] %s %s query=%s body=<unreadable>", method, path, query)
    else:
        logger.info("[REQ] %s %s query=%s", method, path, query)

    try:
        response = await call_next(request)
    except Exception as exc:
        elapsed = time.time() - start_time
        logger.error("[RESP] %s %s elapsed=%.3fs exception=%s", method, path, elapsed, exc, exc_info=True)
        raise

    elapsed = time.time() - start_time
    resp_content_type = response.headers.get("content-type", "")

    if "application/json" in resp_content_type:
        resp_body = b""
        async for chunk in response.body_iterator:
            resp_body += chunk
        body_str = resp_body.decode("utf-8", errors="replace")
        log_body = body_str[:1000] + "...(truncated)" if len(body_str) > 1000 else body_str
        logger.info("[RESP] %s %s status=%d elapsed=%.3fs body=%s", method, path, response.status_code, elapsed, log_body)
        _skip_headers = {"transfer-encoding", "content-length"}
        safe_headers = {k: v for k, v in response.headers.items() if k.lower() not in _skip_headers}
        return Response(
            content=resp_body,
            status_code=response.status_code,
            headers=safe_headers,
            media_type=resp_content_type,
        )
    else:
        logger.info("[RESP] %s %s status=%d elapsed=%.3fs body=<non-json>", method, path, response.status_code, elapsed)
        return response
