import datetime
import logging
import os
import random
import shutil
import string
import uuid

logger = logging.getLogger(__name__)


def generate_unique_filename(original_filename: str) -> str:
    """生成唯一文件名，避免冲突"""
    if "." in original_filename:
        file_ext = os.path.splitext(original_filename)[-1].lower()
    else:
        file_ext = ".docx"
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    random_str = "".join(random.choices(string.ascii_letters + string.digits, k=6))
    return f"{timestamp}_{random_str}{file_ext}"


def generate_unique_file_id() -> str:
    """生成唯一的 file_id，用于拆分接口调用"""
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
    random_str = "".join(random.choices(string.ascii_letters + string.digits, k=8))
    return f"docx_{timestamp}_{random_str}"


def cleanup_temp_files(temp_dir: str) -> None:
    """删除指定的临时目录，忽略不存在的目录"""
    if os.path.exists(temp_dir):
        try:
            shutil.rmtree(temp_dir)
            logger.debug(f"临时目录 {temp_dir} 已清理")
        except Exception as e:
            logger.debug(f"清理临时目录失败：{str(e)}")


def read_logs(file_path: str, level: str = None, keyword: str = None, limit: int = 100) -> list:
    """读取日志文件，支持按级别/关键词筛选"""
    if not os.path.exists(file_path):
        return ["日志文件不存在"]

    logs = []
    with open(file_path, "r", encoding="utf-8") as f:
        all_lines = f.readlines()
        for line in reversed(all_lines):
            line = line.strip()
            if not line:
                continue
            if level and f" - {level.upper()} - " not in line:
                continue
            if keyword and keyword not in line:
                continue
            logs.append(line)
            if len(logs) >= limit:
                break

    return logs[::-1]
