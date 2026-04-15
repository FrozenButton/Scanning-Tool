"""Centralized logging infrastructure."""

import sys
from loguru import logger

def setup_logging(level: str = "INFO") -> None:
    """Initialize structured logging across the application."""
    logger.remove()
    logger.add(
        sys.stdout,
        level=level,
        format="<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
        enqueue=True,
    )
    logger.debug("Structured logging initialized.")
