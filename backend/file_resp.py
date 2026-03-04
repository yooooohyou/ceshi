# -*- coding:utf-8 -*-
# @Time     : 2024/10/31 16:33
# @Author   : xxy
# @File     : file_response
# @Software : PyCharm

import pathlib
import re
from urllib import parse
from fastapi import Request, HTTPException
from fastapi.responses import StreamingResponse, Response


class FileResp:
    """FastAPI文件下载响应类，支持分片下载(Range请求)"""

    def __init__(self, request: Request, file_path: pathlib.Path):
        self.request = request
        self.file_path = file_path

    def start(self) -> Response:
        # 统一文件路径类型
        if not isinstance(self.file_path, pathlib.Path):
            file_path = pathlib.Path(self.file_path)
        else:
            file_path = self.file_path

        # 检查文件是否存在
        if not file_path.exists():
            raise HTTPException(status_code=404, detail="File not found.")

        # 获取文件大小
        file_size = file_path.stat().st_size
        # 编码文件名（避免中文乱码）
        file_name = parse.quote(file_path.name)

        # 获取Range请求头
        range_header = self.request.headers.get("range", None)

        if range_header:
            # 解析Range头（格式：bytes=start-end）
            match = re.match(r'bytes=(\d+)-(\d+)?', range_header)
            if not match:
                raise HTTPException(status_code=400, detail="Invalid Range header")

            start = int(match.group(1))
            end = match.group(2)

            # 处理end为空的情况（bytes=start- 表示到文件末尾）
            if end is None:
                end = file_size - 1
            else:
                end = int(end)

            # 校验Range范围有效性
            if start > end or start >= file_size:
                # 416状态码：请求的范围无法满足
                raise HTTPException(
                    status_code=416,
                    detail=f"Requested range not satisfiable. Valid range: 0-{file_size - 1}"
                )

            # 确保end不超过文件大小
            if end >= file_size:
                end = file_size - 1

            # 定义分片文件读取生成器（避免一次性加载大文件）
            def file_iterator():
                with open(file_path, "rb") as f:
                    f.seek(start)
                    remaining = end - start + 1
                    chunk_size = 8192  # 8KB分片读取
                    while remaining > 0:
                        chunk = f.read(min(chunk_size, remaining))
                        if not chunk:
                            break
                        yield chunk
                        remaining -= len(chunk)

            # 构建206 Partial Content响应
            response = StreamingResponse(
                content=file_iterator(),
                status_code=206
            )
            # 设置分片响应头
            response.headers["Content-Range"] = f"bytes {start}-{end}/{file_size}"
            response.headers["Content-Length"] = str(end - start + 1)
            response.headers["Total-Length"] = str(file_size)
            response.headers["Content-Disposition"] = f"attachment; filename={file_name}"
            response.headers["Content-Type"] = "application/octet-stream"
            response.headers["Access-Control-Expose-Headers"] = "Content-Length, Content-Range, Total-Length"
        else:
            # 普通下载（非分片）
            response = StreamingResponse(
                content=open(file_path, "rb"),
                media_type="application/octet-stream"
            )
            response.headers["Content-Disposition"] = f"attachment; filename={file_name}"
            response.headers["Content-Length"] = str(file_size)

        return response