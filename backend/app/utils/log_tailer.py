import asyncio
import logging
import os

logger = logging.getLogger(__name__)


async def tail_file(path: str, from_end: bool = True, poll_interval: float = 0.5):
    """异步 tail -f：逐行 yield 日志文件新增内容。

    - 自动检测 RotatingFileHandler 轮转（inode 变化或文件被截断），重新打开。
    - 行缓冲：仅在遇到换行符时 yield 完整行，避免推送半行。
    - 文件不存在时每秒重试，连接保持。
    """
    while True:
        try:
            f = open(path, "r", encoding="utf-8", errors="replace")
        except FileNotFoundError:
            await asyncio.sleep(1.0)
            continue

        try:
            inode = os.fstat(f.fileno()).st_ino
            if from_end:
                f.seek(0, os.SEEK_END)
            buf = ""
            while True:
                chunk = f.readline()
                if chunk:
                    buf += chunk
                    if buf.endswith("\n"):
                        yield buf.rstrip("\n")
                        buf = ""
                    continue

                await asyncio.sleep(poll_interval)
                try:
                    st = os.stat(path)
                except FileNotFoundError:
                    break
                if st.st_ino != inode or f.tell() > st.st_size:
                    # 轮转或截断，跳出内层重新 open
                    break
        finally:
            f.close()

        # 重新打开后从文件头开始读（避免漏掉刚轮转后写入的新行）
        from_end = False
