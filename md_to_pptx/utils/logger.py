"""Logger Utility Module.

Provides standard structured logging configuration for console output.
"""

from __future__ import annotations
import logging
import sys


def setup_logger(name: str = "md_to_pptx", level: int = logging.INFO) -> logging.Logger:
    """Setup and configure a logger instance with formatted stream handler.

    Args:
        name: Logger name.
        level: Minimum logging level.

    Returns:
        Configured Logger instance.
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)

    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        handler.setLevel(level)
        formatter = logging.Formatter(
            "[%(asctime)s] [%(levelname)s] [%(name)s]: %(message)s", datefmt="%H:%M:%S"
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger
