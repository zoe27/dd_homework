import sys
from loguru import logger
import config

# 移除默认 handler
logger.remove()

# 控制台输出：简洁格式
logger.add(
    sys.stdout,
    level=config.LOG_LEVEL,
    format="<green>{time:HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan> - <level>{message}</level>",
    colorize=True,
)

# 文件输出：完整格式，按天滚动
logger.add(
    f"{config.LOG_DIR}/dd_homework_{{time:YYYY-MM-DD}}.log",
    level="DEBUG",
    format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
    rotation="00:00",   # 每天零点新建文件
    retention="7 days", # 保留7天
    encoding="utf-8",
)
