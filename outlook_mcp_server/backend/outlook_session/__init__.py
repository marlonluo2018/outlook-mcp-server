"""
Outlook Session Management Package for MCP Server.

This package provides comprehensive Outlook session management functionality,
including connection handling, folder operations, email operations, and utility functions.
"""

from .session_manager import OutlookSessionManager
from .folder_operations import FolderOperations
from .email_operations import EmailOperations
from .exceptions import (
    OutlookSessionError,
    ConnectionError,
    FolderNotFoundError,
    EmailNotFoundError,
)
from .utils import (
    safe_com_call,
    retry_on_com_error,
    get_outlook_version,
    validate_outlook_installation,
    format_com_error,
    safe_release_com_object,
    get_available_folders,
    parse_folder_path,
    sanitize_folder_name,
    convert_com_time_to_string,
    COMObjectWrapper
)
from ..validation import validate_email_address
from .decorators import (
    safe_com_operation,
    log_com_operation,
    handle_com_errors,
    timeout_com_operation
)

__all__ = [
    'OutlookSessionManager',
    'FolderOperations',
    'EmailOperations',
    'OutlookSessionError',
    'ConnectionError',
    'FolderNotFoundError',
    'EmailNotFoundError',
    'safe_com_call',
    'retry_on_com_error',
    'get_outlook_version',
    'validate_outlook_installation',
    'format_com_error',
    'safe_release_com_object',
    'get_available_folders',
    'parse_folder_path',
    'validate_email_address',
    'sanitize_folder_name',
    'convert_com_time_to_string',
    'COMObjectWrapper',
    'safe_com_operation',
    'log_com_operation',
    'handle_com_errors',
    'timeout_com_operation'
]