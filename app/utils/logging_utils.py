"""
Logging utilities for the Report Generator API
"""

import logging
import sys
from typing import Optional

import colorlog


def setup_logging(log_level: str = "INFO", log_format: Optional[str] = None):
    """
    Setup logging configuration with colored output

    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_format: Custom log format string
    """
    # Default colored format
    if log_format is None:
        log_format = "%(log_color)s%(asctime)s - %(name)s - %(levelname)s - %(message)s"

    # Create colored formatter
    formatter = colorlog.ColoredFormatter(
        log_format,
        datefmt="%Y-%m-%d %H:%M:%S",
        log_colors={
            "DEBUG": "cyan",
            "INFO": "green",
            "WARNING": "yellow",
            "ERROR": "red",
            "CRITICAL": "red,bg_white",
        },
    )

    # Configure root logger
    logger = logging.getLogger()
    logger.setLevel(getattr(logging, log_level.upper()))

    # Remove existing handlers
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # Add console handler with colored formatter
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Set specific loggers to appropriate levels
    logging.getLogger("uvicorn").setLevel(logging.INFO)
    logging.getLogger("fastapi").setLevel(logging.INFO)


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger instance for a specific module

    Args:
        name: Logger name (usually __name__)

    Returns:
        Logger instance
    """
    return logging.getLogger(name)


def print_blue(message: str):
    """Print message in blue color for startup messages"""
    print(f"\033[94m{message}\033[0m")


def print_success(message: str):
    """Print message in green color for success messages"""
    print(f"\033[92m{message}\033[0m")


def print_warning(message: str):
    """Print message in yellow color for warning messages"""
    print(f"\033[93m{message}\033[0m")


def print_error(message: str):
    """Print message in red color for error messages"""
    print(f"\033[91m{message}\033[0m")
