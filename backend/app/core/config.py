import os
import platform
import configparser
import logging
from pathlib import Path
from fastapi.staticfiles import StaticFiles

logger = logging.getLogger(__name__)

# ─── 配置文件读取 ────────────────────────────────────────────────────────────

def read_sc_web_config(config_filename: str = "sc_web.conf") -> configparser.ConfigParser:
    """读取 conf/sc_web.conf 配置文件"""
    config_dir = Path(__file__).parent.parent.parent / "conf"
    config_file = config_dir / config_filename

    if not config_file.exists():
        raise FileNotFoundError(f"配置文件不存在: {config_file}")
    if not os.access(config_file, os.R_OK):
        raise PermissionError(f"没有读取权限: {config_file}")

    config = configparser.ConfigParser()
    try:
        config.read(config_file, encoding="utf-8")
    except Exception as e:
        raise Exception(f"解析配置文件失败: {str(e)}")
    return config


def get_server_uploads_config() -> dict:
    """快捷获取 [server_uploads] 节的所有配置"""
    config = read_sc_web_config()
    if "server_uploads" not in config.sections():
        raise Exception("配置文件中未找到[server_uploads]节")
    return {
        "user_local_path":  config.get("server_uploads", "user_local_path",  fallback=""),
        "web_backend_path": config.get("server_uploads", "web_backend_path", fallback=""),
        "web_front_path":   config.get("server_uploads", "web_front_path",   fallback=""),
    }


def get_docx_render_max_workers(default: int = 5) -> int:
    """读取 [docx_render] max_workers；section/key 缺失或非法时回退到 default"""
    try:
        config = read_sc_web_config()
        raw = config.get("docx_render", "max_workers", fallback=str(default))
        value = int(raw)
        return value if value >= 1 else default
    except Exception:
        return default


# ─── 路径常量（根据操作系统自动选择） ────────────────────────────────────────

system_path = platform.system()

if system_path == "Windows":
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
    STATIC_WEB_PREFIX = "/uploads"
    STATIC_WEB_FRONT_PREFIX = "/uploads"
    WEB_File_Path = False
    try:
        os.makedirs(UPLOAD_DIR, exist_ok=True)
    except PermissionError:
        UPLOAD_DIR = os.path.join(os.gettempdir(), "docx_uploads")
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        logger.warning(f"警告：无法在当前目录创建uploads文件夹，已切换到系统临时目录：{UPLOAD_DIR}")
else:
    _uploads_config = get_server_uploads_config()
    UPLOAD_DIR = _uploads_config["user_local_path"]
    STATIC_WEB_PREFIX = _uploads_config["web_backend_path"]
    STATIC_WEB_FRONT_PREFIX = _uploads_config["web_front_path"]
    WEB_File_Path = True


# ─── PostgreSQL 配置 ──────────────────────────────────────────────────────────

_pg = read_sc_web_config()["postgres"]
POSTGRES_CONFIG = {
    "host":     _pg.get("host"),
    "port":     int(_pg.get("port")),
    "user":     _pg.get("user"),
    "password": _pg.get("password"),
    "database": _pg.get("database"),
    "options":  _pg.get("options"),
}