"""
Utility functions for Outlook MCP server session management.

This module provides utility functions for handling Outlook sessions, 
COM objects, and common operations that don't fit into the main classes.
"""

# Standard library imports
import time
from functools import wraps
from typing import Any, Dict, List, Optional, Union

# Third-party imports
import pythoncom
import win32com.client

# Local application imports
from ..logging_config import get_logger

logger = get_logger(__name__)


def safe_com_call(func):
    """
    Decorator to safely handle COM object calls with proper error handling.
    
    This decorator:
    - Handles COM errors gracefully
    - Ensures COM objects are properly released
    - Provides detailed logging for debugging
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
            logger.debug(f"COM call successful: {func.__name__}")
            return result
        except pythoncom.com_error as com_err:
            logger.error(f"COM error in {func.__name__}: {com_err}")
            raise
        except AttributeError as attr_err:
            logger.error(f"Attribute error in {func.__name__}: {attr_err}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error in {func.__name__}: {e}")
            raise
    
    return wrapper


def retry_on_com_error(max_attempts: int = 3, initial_delay: float = 1.0, backoff_factor: float = 2.0):
    """
    Decorator to retry COM operations on failure with exponential backoff.
    
    Args:
        max_attempts: Maximum number of retry attempts
        initial_delay: Initial delay between retries in seconds
        backoff_factor: Factor to multiply delay by for each retry
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            delay = initial_delay
            last_exception = None
            
            for attempt in range(max_attempts):
                try:
                    result = func(*args, **kwargs)
                    if attempt > 0:
                        logger.info(f"COM operation succeeded on attempt {attempt + 1}")
                    return result
                except pythoncom.com_error as com_err:
                    last_exception = com_err
                    logger.warning(f"COM operation failed on attempt {attempt + 1}: {com_err}")
                    
                    if attempt < max_attempts - 1:
                        logger.info(f"Retrying in {delay} seconds...")
                        time.sleep(delay)
                        delay *= backoff_factor
                    else:
                        logger.error(f"COM operation failed after {max_attempts} attempts")
                        raise
                except Exception as e:
                    # Don't retry on non-COM errors
                    logger.error(f"Non-COM error in {func.__name__}: {e}")
                    raise
            
            # This should never be reached, but just in case
            if last_exception:
                raise last_exception
                
        return wrapper
    return decorator


def get_outlook_version() -> Optional[str]:
    """
    Get the version of Outlook installed on the system.
    
    Returns:
        str: Outlook version string, or None if Outlook is not available
    """
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        version = outlook.Version
        logger.info(f"Detected Outlook version: {version}")
        return version
    except Exception as e:
        logger.error(f"Failed to get Outlook version: {e}")
        return None
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass


def validate_outlook_installation() -> bool:
    """
    Validate that Outlook is properly installed and accessible.
    
    Returns:
        bool: True if Outlook is accessible, False otherwise
    """
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Try to access a basic folder to verify functionality
        inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
        
        logger.info("Outlook installation validated successfully")
        return True
        
    except Exception as e:
        logger.error(f"Outlook validation failed: {e}")
        return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass


def format_com_error(error: pythoncom.com_error) -> str:
    """
    Format a COM error into a readable string.
    
    Args:
        error: The COM error object
        
    Returns:
        str: Formatted error message
    """
    try:
        if hasattr(error, 'excepinfo') and error.excepinfo:
            return f"COM Error: {error.excepinfo[2]} (0x{error.excepinfo[5]:08X})"
        elif hasattr(error, 'strerror'):
            return f"COM Error: {error.strerror}"
        else:
            return f"COM Error: {str(error)}"
    except:
        return f"COM Error: {str(error)}"


def safe_release_com_object(obj: Any) -> None:
    """
    Safely release a COM object.
    
    Args:
        obj: The COM object to release
    """
    try:
        if obj is not None:
            # Release COM object
            obj = None
            logger.debug("COM object released")
    except Exception as e:
        logger.warning(f"Error releasing COM object: {e}")


def get_available_folders() -> List[Dict[str, Any]]:
    """
    Get a list of available default folder types in Outlook.
    
    Returns:
        List[Dict[str, Any]]: List of folder information dictionaries
    """
    return [
        {"name": "Inbox", "constant": 6, "description": "Inbox folder"},
        {"name": "Sent Items", "constant": 5, "description": "Sent items folder"},
        {"name": "Deleted Items", "constant": 3, "description": "Deleted items folder"},
        {"name": "Drafts", "constant": 16, "description": "Drafts folder"},
        {"name": "Outbox", "constant": 4, "description": "Outbox folder"},
        {"name": "Junk Email", "constant": 23, "description": "Junk email folder"},
        {"name": "Calendar", "constant": 9, "description": "Calendar folder"},
        {"name": "Contacts", "constant": 10, "description": "Contacts folder"},
        {"name": "Tasks", "constant": 13, "description": "Tasks folder"},
        {"name": "Notes", "constant": 12, "description": "Notes folder"},
        {"name": "Journal", "constant": 11, "description": "Journal folder"},
    ]


def parse_folder_path(folder_path: str) -> Dict[str, Any]:
    """
    Parse a folder path string into components.
    
    Args:
        folder_path: The folder path string (e.g., "Inbox/Subfolder" or "mailbox@domain.com/Inbox")
        
    Returns:
        Dict[str, Any]: Parsed path information
    """
    if not folder_path or folder_path.lower() == "inbox":
        return {
            "is_default": True,
            "default_type": "Inbox",
            "path_parts": [],
            "mailbox": None
        }
    
    # Check if it's a mailbox-specific path
    if "/" in folder_path:
        parts = folder_path.split("/")
        
        # Check if first part looks like an email address
        if "@" in parts[0] and "." in parts[0]:
            return {
                "is_default": False,
                "default_type": None,
                "path_parts": parts[1:],
                "mailbox": parts[0]
            }
        else:
            return {
                "is_default": False,
                "default_type": None,
                "path_parts": parts,
                "mailbox": None
            }
    else:
        # Single folder name
        return {
            "is_default": False,
            "default_type": None,
            "path_parts": [folder_path],
            "mailbox": None
        }


def sanitize_folder_name(name: str) -> str:
    """
    Sanitize a folder name for use in Outlook.
    
    Args:
        name: Folder name to sanitize
        
    Returns:
        str: Sanitized folder name
    """
    if not name:
        return ""
    
    # Remove leading/trailing whitespace
    name = name.strip()
    
    # Remove invalid characters for folder names
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for char in invalid_chars:
        name = name.replace(char, '_')
    
    return name


def convert_com_time_to_string(com_time) -> Optional[str]:
    """
    Convert COM time object to string format.
    
    Args:
        com_time: COM time object
        
    Returns:
        str: Time string in ISO format, or None if conversion fails
    """
    try:
        if com_time is None:
            return None
        
        # Convert COM time to Python datetime
        import datetime
        if hasattr(com_time, 'year') and hasattr(com_time, 'month'):
            dt = datetime.datetime(com_time.year, com_time.month, com_time.day,
                                 com_time.hour, com_time.minute, com_time.second)
            return dt.isoformat()
        else:
            return str(com_time)
            
    except Exception as e:
        logger.warning(f"Failed to convert COM time: {e}")
        return None


class COMObjectWrapper:
    """
    Wrapper class for safe COM object handling.
    
    This class provides a context manager for COM objects with automatic
    cleanup and error handling.
    """
    
    def __init__(self, com_object: Any):
        """Initialize with a COM object."""
        self.com_object = com_object
        self._released = False
    
    def __enter__(self):
        """Enter context manager."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Exit context manager and cleanup."""
        self.release()
        return False  # Don't suppress exceptions
    
    def release(self):
        """Release the COM object."""
        if not self._released and self.com_object is not None:
            safe_release_com_object(self.com_object)
            self._released = True
    
    def __getattr__(self, name):
        """Delegate attribute access to the wrapped COM object."""
        if self.com_object is None:
            raise RuntimeError("COM object has been released")
        
        try:
            return getattr(self.com_object, name)
        except AttributeError as e:
            logger.error(f"Attribute '{name}' not found on COM object")
            raise
    
    def __bool__(self):
        """Check if COM object is valid."""
        return self.com_object is not None and not self._released