import json
from fastapi.responses import JSONResponse


class UnescapedJSONResponse(JSONResponse):
    """保留非 ASCII 字符（如中文）的 JSON 响应"""

    def render(self, content) -> bytes:
        return json.dumps(
            content,
            ensure_ascii=False,
            allow_nan=False,
            indent=None,
            separators=(",", ":"),
            default=str,
        ).encode("utf-8")


def unified_response(code: int, message: str, data: dict = None) -> JSONResponse:
    """生成统一格式的响应体"""
    return JSONResponse(
        status_code=200,
        content={
            "code": code,
            "message": message,
            "data": data or {},
        },
    )
