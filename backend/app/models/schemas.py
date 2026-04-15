import json
from typing import List
from fastapi.responses import JSONResponse
from pydantic import BaseModel


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


class TreeNodeUpdate(BaseModel):
    node_id: int
    level: int
    children: List["TreeNodeUpdate"] = []

TreeNodeUpdate.model_rebuild()


class UpdateTreeStructureRequest(BaseModel):
    record_id: int
    node_ids: List[TreeNodeUpdate]


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