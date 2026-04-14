from typing import Literal

# 默认主节点配置
DEFAULT_MAIN_NODE = {
    "title": "文档内容",
    "level": 1,
    "eid": "main_node",
    "idx": 0,
}

# 处理模式定义（默认 split）
ProcessMode = Literal["single", "split"]

# LibreOffice 路径
LIBREOFFICE_PATH = "libreoffice"
