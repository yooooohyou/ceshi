import logging
import sys
from logging.handlers import RotatingFileHandler


def setup_logging() -> str:
    log_format = "%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s"
    log_level = logging.INFO
    log_file = "app.log"

    logging.basicConfig(
        level=log_level,
        format=log_format,
        handlers=[
            logging.StreamHandler(sys.stdout),
            RotatingFileHandler(
                log_file,
                maxBytes=10 * 1024 * 1024,  # 10MB
                backupCount=5,
                encoding="utf-8"
            )
        ]
    )

    logging.getLogger("uvicorn.access").setLevel(logging.WARNING)
    return log_file


log_file_path = setup_logging()
logger = logging.getLogger(__name__)
