"""
Decorators for Outlook MCP server session management.

This module provides decorators for handling COM operations, retries, 
and other common patterns used throughout the session management code.
"""

# Standard library imports
import logging
import time
from functools import wraps
from typing import Any, Callable, Optional

# Third-party imports
import pythoncom

# Local application imports
from ..logging_config import get_logger

logger = get_logger(__name__)


def retry_on_com_error(max_attempts: int = 3, initial_delay: float = 1.0, backoff_factor: float = 2.0):
    """
    Decorator to retry COM operations on failure with exponential backoff.
    
    Args:
        max_attempts: Maximum number of retry attempts
        initial_delay: Initial delay between retries in seconds
        backoff_factor: Factor to multiply delay by for each retry
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            delay = initial_delay
            last_exception = None
            
            for attempt in range(max_attempts):
                try:
                    result = func(*args, **kwargs)
                    if attempt > 0:
                        logger.info(f"COM operation '{func.__name__}' succeeded on attempt {attempt + 1}")
                    return result
                except pythoncom.com_error as com_err:
                    last_exception = com_err
                    logger.warning(f"COM operation '{func.__name__}' failed on attempt {attempt + 1}: {com_err}")
                    
                    if attempt < max_attempts - 1:
                        logger.info(f"Retrying '{func.__name__}' in {delay} seconds...")
                        time.sleep(delay)
                        delay *= backoff_factor
                    else:
                        logger.error(f"COM operation '{func.__name__}' failed after {max_attempts} attempts")
                        raise
                except Exception as e:
                    # Don't retry on non-COM errors
                    logger.error(f"Non-COM error in '{func.__name__}': {e}")
                    raise
            
            # This should never be reached, but just in case
            if last_exception:
                raise last_exception
                
        return wrapper
    return decorator


def safe_com_operation(func: Callable) -> Callable:
    """
    Decorator to safely handle COM operations with proper initialization and cleanup.
    
    This decorator ensures:
    - COM is properly initialized before the operation
    - Proper error handling for COM-specific errors
    - Logging of operation success/failure
    """
    @wraps(func)
    def wrapper(*args, **kwargs) -> Any:
        com_initialized = False
        try:
            # Initialize COM if not already done
            if not pythoncom._GetInterfaceCount():
                pythoncom.CoInitialize()
                com_initialized = True
                logger.debug(f"COM initialized for '{func.__name__}'")
            
            # Execute the function
            result = func(*args, **kwargs)
            logger.debug(f"COM operation '{func.__name__}' completed successfully")
            return result
            
        except pythoncom.com_error as com_err:
            logger.error(f"COM error in '{func.__name__}': {com_err}")
            raise
        except AttributeError as attr_err:
            logger.error(f"Attribute error in '{func.__name__}': {attr_err}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error in '{func.__name__}': {e}")
            raise
        finally:
            # Clean up COM if we initialized it
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    logger.debug(f"COM uninitialized for '{func.__name__}'")
                except:
                    pass
    
    return wrapper


def log_com_operation(level: int = logging.INFO, include_args: bool = False, include_result: bool = False):
    """
    Decorator to log COM operations.
    
    Args:
        level: Logging level to use
        include_args: Whether to include function arguments in the log
        include_result: Whether to include function result in the log
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            func_name = func.__name__
            
            # Log operation start
            if include_args:
                logger.log(level, f"Starting COM operation '{func_name}' with args: {args}, kwargs: {kwargs}")
            else:
                logger.log(level, f"Starting COM operation '{func_name}'")
            
            try:
                result = func(*args, **kwargs)
                
                # Log operation success
                if include_result:
                    logger.log(level, f"COM operation '{func_name}' completed successfully with result: {result}")
                else:
                    logger.log(level, f"COM operation '{func_name}' completed successfully")
                
                return result
                
            except Exception as e:
                # Log operation failure
                logger.error(f"COM operation '{func_name}' failed: {e}")
                raise
        
        return wrapper
    return decorator


def handle_com_errors(default_return: Optional[Any] = None, log_errors: bool = True):
    """
    Decorator to handle COM errors gracefully.
    
    Args:
        default_return: Value to return on error (if None, raises exception)
        log_errors: Whether to log errors
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            try:
                return func(*args, **kwargs)
            except pythoncom.com_error as com_err:
                if log_errors:
                    logger.error(f"COM error in '{func.__name__}': {com_err}")
                
                if default_return is not None:
                    return default_return
                else:
                    raise
            except Exception as e:
                if log_errors:
                    logger.error(f"Error in '{func.__name__}': {e}")
                
                if default_return is not None:
                    return default_return
                else:
                    raise
        
        return wrapper
    return decorator


def timeout_com_operation(timeout_seconds: float = 30.0):
    """
    Decorator to add timeout protection to COM operations.
    
    Note: This is a best-effort timeout as Python's COM operations
    are not easily interruptible. It will prevent new operations after
    the timeout but won't interrupt ongoing COM calls.
    
    Args:
        timeout_seconds: Timeout in seconds
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            start_time = time.time()
            
            try:
                # Check if we've exceeded timeout before starting
                if time.time() - start_time > timeout_seconds:
                    raise TimeoutError(f"COM operation '{func.__name__}' timed out before starting")
                
                result = func(*args, **kwargs)
                
                # Check if operation completed within timeout
                elapsed = time.time() - start_time
                if elapsed > timeout_seconds:
                    logger.warning(f"COM operation '{func.__name__}' completed but exceeded timeout ({elapsed:.2f}s > {timeout_seconds}s)")
                
                return result
                
            except TimeoutError:
                logger.error(f"COM operation '{func.__name__}' timed out after {timeout_seconds} seconds")
                raise
            except Exception as e:
                elapsed = time.time() - start_time
                logger.error(f"COM operation '{func.__name__}' failed after {elapsed:.2f}s: {e}")
                raise
        
        return wrapper
    return decorator