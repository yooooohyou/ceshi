# -*- coding:utf-8 -*-
# @Time     : 2024/10/31 16:33
# @Author   : xxy
# @File     : file_response
# @Software : PyCharm

import pathlib
import re
from urllib import parse
import json
import os
import shutil
import time
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


def SplitUpload(split_file_path, file_no, file_name, files_total_count, file_sign, file):
    """
    文件分块上传
    :param split_file_path: 文件上传存储根目录
    :param file_no: 当前块序号
    :param file_name: 源文件名称
    :param files_total_count: 文件总块数
    :param file_sign: 文件标识（MD5）
    :param file: 文件（request.FILES.get所得值）
    :return: 当前块文件处理是否成功(0:成功,1:失败)[int],文件最终步骤（合并）是否完成(0:未完成,1:完成)[int],错误信息[str]
    """
    sign_file_lock = sign_file_path = None
    try:
        if file_no and file_name and files_total_count and file_sign and file:
            root_path = os.path.join(split_file_path, file_sign)  # 分块文件存储目录
            sign_file_path = os.path.join(str(root_path), 'sign.json')  # 分块信息文件
            out_file = os.path.join(split_file_path, file_name)  # 最终输出文件
            sign_file_lock = os.path.join(str(root_path), f'sign_lock.json')
            if not os.path.exists(split_file_path):
                return 1, 0, '根目录不存在'

            # 临时文件夹初始化
            if not os.path.exists(root_path):
                os.mkdir(root_path)
                file_dict = {str(x): 0 for x in range(1, int(files_total_count) + 1)}
                with open(sign_file_path, 'w', encoding='utf-8') as f:
                    f.write(json.dumps(file_dict))

            # 保存拆分文件
            with open(os.path.join(str(root_path), str(file_no)), 'wb') as f:
                for chunk in file.chunks():
                    f.write(chunk)

            # 读取并修改拆分文件标识文件(加锁)
            while 1:
                try:
                    os.rename(sign_file_path, sign_file_lock)
                    str_conf = open(sign_file_lock, 'r', encoding='utf-8').read()
                    break
                except FileNotFoundError:
                    time.sleep(0.05)
            sign_data = json.loads(str_conf)
            if str(files_total_count) not in sign_data or str(int(files_total_count)+1) in sign_data:
                os.rename(sign_file_lock, sign_file_path)
                return 1, 0, '文件分块信息错误，请从新上传！'
            sign_data[str(file_no)] = 1
            with open(sign_file_lock, 'w', encoding='utf-8') as f:
                f.write(json.dumps(sign_data))
            os.rename(sign_file_lock, sign_file_path)

            # 判断所有文件是否上传完毕,完成则合并文件
            finish = 1
            for i in range(1, int(files_total_count) + 1):
                if sign_data[str(i)] == 0:
                    finish = 0
                    break
            if finish:
                with open(out_file, 'wb') as outfile:
                    for i in range(1, int(files_total_count) + 1):
                        file_path = os.path.join(str(root_path), str(i))
                        with open(file_path, 'rb') as infile:
                            while True:
                                chunk = infile.read(4096)  # 读取4KB大小的块
                                if not chunk:
                                    break
                                outfile.write(chunk)
                shutil.rmtree(root_path)
            return 0, finish, '成功'
        else:
            return 1, 0, '参数有空值'
    except Exception as e:
        if sign_file_lock and sign_file_path:
            if os.path.exists(sign_file_lock):
                os.rename(sign_file_lock, sign_file_path)
        raise e