"""Logging configuration for Outlook MCP Server."""

import logging
import os
import sys
from typing import Optional

# Environment variable to control debug logging
DEBUG_ENV_VAR = "OUTLOOK_MCP_DEBUG"
LOG_LEVEL_ENV_VAR = "OUTLOOK_MCP_LOG_LEVEL"

# Default log format
DEFAULT_LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
DEBUG_LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s"


def get_log_level() -> int:
    """Get the configured log level from environment or defaults.
    
    Returns:
        int: Logging level constant (e.g., logging.DEBUG, logging.INFO)
    """
    # Check for explicit log level environment variable
    log_level_str = os.getenv(LOG_LEVEL_ENV_VAR, "").upper()
    
    if log_level_str:
        level_map = {
            "DEBUG": logging.DEBUG,
            "INFO": logging.INFO,
            "WARNING": logging.WARNING,
            "ERROR": logging.ERROR,
            "CRITICAL": logging.CRITICAL,
        }
        return level_map.get(log_level_str, logging.INFO)
    
    # Check for debug mode environment variable
    if os.getenv(DEBUG_ENV_VAR, "").lower() in ("1", "true", "yes", "on"):
        return logging.DEBUG
    
    return logging.INFO


def get_log_format(debug_mode: bool = False) -> str:
    """Get the appropriate log format based on debug mode.
    
    Args:
        debug_mode: Whether debug mode is enabled
        
    Returns:
        str: Log format string
    """
    return DEBUG_LOG_FORMAT if debug_mode else DEFAULT_LOG_FORMAT


def configure_logging(level: Optional[int] = None, format_string: Optional[str] = None) -> None:
    """Configure the root logger with the specified settings.
    
    Args:
        level: Log level to set (defaults to value from get_log_level())
        format_string: Log format string (defaults to value from get_log_format())
    """
    if level is None:
        level = get_log_level()
    
    debug_mode = level == logging.DEBUG
    
    if format_string is None:
        format_string = get_log_format(debug_mode)
    
    # Configure root logger
    logging.basicConfig(
        level=level,
        format=format_string,
        stream=sys.stderr,
        force=True
    )


def get_logger(name: str) -> logging.Logger:
    """Get a logger instance with the specified name.
    
    Args:
        name: Logger name (typically __name__)
        
    Returns:
        logging.Logger: Configured logger instance
    """
    return logging.getLogger(name)


def is_debug_enabled() -> bool:
    """Check if debug logging is enabled.
    
    Returns:
        bool: True if debug logging is enabled
    """
    return get_log_level() == logging.DEBUG


def set_debug_mode(enabled: bool = True) -> None:
    """Enable or disable debug mode at runtime.
    
    Args:
        enabled: Whether to enable debug mode
    """
    level = logging.DEBUG if enabled else logging.INFO
    configure_logging(level=level)


# Auto-configure logging on module import
configure_logging()
