# -*- coding:utf-8 -*-
# @Time     : 2024/10/31 16:33
# @Author   : xxy
# @File     : file_response
# @Software : PyCharm

import logging
import pathlib

logger = logging.getLogger(__name__)
import re
from urllib import parse
import json
import os
import shutil
import time
from fastapi import Request, HTTPException, UploadFile
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


def SplitUpload(
        split_file_path: str,
        file_no: str | int,
        file_name: str,
        files_total_count: str | int,
        file_sign: str,
        file  # 可以是 bytes/文件对象，不再依赖 UploadFile
) -> tuple[int, int, str]:
    """
    文件分块上传（同步版本，适配非异步场景）
    :param split_file_path: 文件上传存储根目录
    :param file_no: 当前块序号
    :param file_name: 源文件名称
    :param files_total_count: 文件总块数
    :param file_sign: 文件标识（MD5）
    :param file: 文件数据（bytes 或 可读取的文件对象）
    :return: (处理状态(0成功,1失败), 合并完成状态(0未完成,1完成), 错误信息)
    """
    sign_file_lock = sign_file_path = None
    try:
        # 1. 严格参数校验 + 类型转换
        if not all([file_no, file_name, files_total_count, file_sign, file]):
            return 1, 0, '参数有空值'

        # 统一转换为整数（避免字符串序号导致的问题）
        try:
            file_no = int(file_no)
            files_total_count = int(files_total_count)
        except ValueError:
            return 1, 0, '块序号/总块数必须是数字'

        # 2. 目录初始化校验
        root_path = os.path.join(split_file_path, file_sign)  # 分块存储目录
        sign_file_path = os.path.join(root_path, 'sign.json')  # 分块进度文件
        out_file = os.path.join(split_file_path, file_name)  # 最终合并文件
        sign_file_lock = os.path.join(root_path, 'sign_lock.json')  # 锁文件
        # print(out_file)
        if not os.path.exists(split_file_path):
            return 1, 0, '根目录不存在'

        # 3. 初始化分块目录和进度文件
        if not os.path.exists(root_path):
            os.makedirs(root_path, exist_ok=True)  # 替换 mkdir，支持多级目录
            # 初始化进度字典：key为块序号，value为0（未上传）
            file_dict = {str(i): 0 for i in range(1, files_total_count + 1)}
            with open(sign_file_path, 'w', encoding='utf-8') as f:
                json.dump(file_dict, f)  # 替换 dumps，更高效

        # 4. 保存当前分块文件（同步读取，移除 await）
        chunk_file_path = os.path.join(root_path, str(file_no))
        try:
            with open(chunk_file_path, 'wb') as f:
                if isinstance(file, bytes):
                    # 直接传入 bytes：一次性写入
                    f.write(file)
                else:
                    # 其他可读取的文件对象：同步分块读取
                    while chunk := file.read(4096):
                        f.write(chunk)
        except Exception as e:
            return 1, 0, f'保存分块失败：{str(e)}'

        # 5. 加锁更新进度文件（优化死循环，增加超时）
        lock_timeout = 5  # 锁超时时间（秒）
        start_time = time.time()
        while True:
            try:
                # 尝试重命名获取锁
                os.rename(sign_file_path, sign_file_lock)
                break
            except FileNotFoundError:
                if time.time() - start_time > lock_timeout:
                    return 1, 0, '获取文件锁超时'
                time.sleep(0.05)

        # 读取并更新进度
        with open(sign_file_lock, 'r', encoding='utf-8') as f:
            sign_data = json.load(f)

        # 校验分块信息合法性
        if str(files_total_count) not in sign_data or file_no > files_total_count:
            os.rename(sign_file_lock, sign_file_path)
            return 1, 0, '文件分块信息错误，请重新上传'

        sign_data[str(file_no)] = 1  # 标记当前块已上传

        # 写回进度并释放锁
        with open(sign_file_lock, 'w', encoding='utf-8') as f:
            json.dump(sign_data, f)
        os.rename(sign_file_lock, sign_file_path)

        # 6. 检查是否所有块都上传完成，完成则合并
        finish = 1 if all(v == 1 for v in sign_data.values()) else 0
        if finish:
            try:
                # 合并所有分块
                with open(out_file, 'wb') as outfile:
                    for i in range(1, files_total_count + 1):
                        chunk_path = os.path.join(root_path, str(i))
                        with open(chunk_path, 'rb') as infile:
                            while chunk := infile.read(4096):
                                outfile.write(chunk)
                # 删除分块目录
                shutil.rmtree(root_path, ignore_errors=True)
            except Exception as e:
                return 1, 0, f'文件合并失败：{str(e)}'

        return 0, finish, '成功'

    except Exception as e:
        # 异常时释放锁，避免锁文件残留
        if sign_file_lock and sign_file_path and os.path.exists(sign_file_lock):
            try:
                os.rename(sign_file_lock, sign_file_path)
            except:
                pass
        # 打印异常详情，方便调试
        logger.info(f"分块上传异常：{str(e)}")
        return 1, 0, f'上传失败：{str(e)}'