"""Logger configuration for the scanning tool using loguru."""

from loguru import logger
import sys

def setup_logging() -> None:
    """Configure loguru logger with console and file handlers."""
    logger.remove()
    logger.add(sys.stdout, level="INFO", format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>", colorize=True)
    logger.add("scanning_tool.log", rotation="10 MB", retention=5, level="INFO", encoding="utf-8", format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}")
    return logger
