"""Utility functions for email processing and validation"""
import logging
from typing import Any, List, Optional
from datetime import datetime
from enum import IntEnum
from functools import wraps
import time
import pythoncom

logger = logging.getLogger(__name__)


class OutlookFolderType(IntEnum):
    """Outlook folder type constants"""
    DELETED_ITEMS = 3
    OUTBOX = 4
    SENT_MAIL = 5
    INBOX = 6
    CALENDAR = 9
    CONTACTS = 10
    TASKS = 13
    DRAFTS = 16


class OutlookItemClass(IntEnum):
    """Outlook item class constants"""
    MAIL_ITEM = 43
    APPOINTMENT_ITEM = 26
    CONTACT_ITEM = 40
    TASK_ITEM = 48


def safe_encode_text(text: Any, field_name: str = "text") -> str:
    """
    Centralized encoding handler with consistent strategy.
    
    Args:
        text: The text to encode (can be bytes, str, or other types)
        field_name: Name of the field being encoded (for logging)
        
    Returns:
        str: Properly encoded string
    """
    if text is None:
        return ""
        
    if isinstance(text, str):
        return text
        
    if isinstance(text, bytes):
        # Try multiple encodings in order of likelihood
        for encoding in ['utf-8', 'cp1252', 'iso-8859-1', 'gbk']:
            try:
                return text.decode(encoding)
            except (UnicodeDecodeError, LookupError):
                continue
        
        # If all encodings fail, use replacement characters
        logger.warning(f"Failed to decode {field_name}, using replacement characters")
        return text.decode('utf-8', errors='replace')
    
    # For any other type, convert to string
    return str(text)


def retry_on_com_error(max_attempts: int = 3, initial_delay: float = 1.0):
    """
    Decorator to retry COM operations on transient errors.
    
    Args:
        max_attempts: Maximum number of retry attempts
        initial_delay: Initial delay between retries (exponential backoff)
        
    Returns:
        Decorated function with retry logic
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            last_exception = None
            
            for attempt in range(max_attempts):
                try:
                    return func(*args, **kwargs)
                except pythoncom.com_error as e:
                    last_exception = e
                    if attempt == max_attempts - 1:
                        logger.error(f"COM error after {max_attempts} attempts in {func.__name__}: {e}")
                        raise
                    
                    delay = initial_delay * (2 ** attempt)
                    logger.warning(
                        f"COM error on attempt {attempt + 1}/{max_attempts} in {func.__name__}, "
                        f"retrying in {delay}s... Error: {e}"
                    )
                    time.sleep(delay)
                except Exception as e:
                    # Don't retry on non-COM errors
                    logger.error(f"Non-COM error in {func.__name__}: {e}")
                    raise
            
            # This shouldn't be reached, but just in case
            if last_exception:
                raise last_exception
                
        return wrapper
    return decorator


def build_dasl_filter(
    search_terms: List[str],
    threshold_date: datetime,
    field_filter: str,
    match_all: bool = True
) -> str:
    """
    Build optimized DASL filter query for Outlook search.
    
    Args:
        search_terms: List of terms to search for
        threshold_date: Date threshold for filtering
        field_filter: Which field to filter ('subject', 'sender', 'recipient', 'body')
        match_all: If True, all terms must match (AND); if False, any term matches (OR)
        
    Returns:
        str: DASL filter string for Outlook Restrict method
    """
    # Field schema mappings
    field_mappings = {
        'subject': 'urn:schemas:httpmail:subject',
        'sender': 'urn:schemas:httpmail:fromname',
        'recipient': 'urn:schemas:httpmail:displayto',
        'body': 'urn:schemas:httpmail:textdescription'
    }
    
    schema = field_mappings.get(field_filter, field_mappings['subject'])
    
    # Build term filters
    if match_all and len(search_terms) > 1:
        # For AND logic: each term must appear in the field
        term_groups = []
        for term in search_terms:
            # Escape single quotes in search terms
            escaped_term = term.replace("'", "''")
            term_groups.append(f'"{schema}" LIKE \'%{escaped_term}%\'')
        filter_logic = ' AND '.join(term_groups)
    else:
        # For OR logic: any term can match
        term_filters = []
        for term in search_terms:
            escaped_term = term.replace("'", "''")
            term_filters.append(f'"{schema}" LIKE \'%{escaped_term}%\'')
        filter_logic = ' OR '.join(term_filters)
    
    # Add date filter
    date_str = threshold_date.strftime('%Y-%m-%d %H:%M:%S')
    date_filter = f'"urn:schemas:httpmail:datereceived" >= \'{date_str}\''
    
    # Combine filters
    combined_filter = f"@SQL=({filter_logic}) AND {date_filter}"
    
    return combined_filter


def get_pagination_info(cache_size: int, per_page: int) -> dict:
    """
    Calculate pagination metadata.
    
    Args:
        cache_size: Total number of items
        per_page: Items per page
        
    Returns:
        dict: Pagination info with total_pages and total_items
    """
    if cache_size == 0:
        return {'total_pages': 0, 'total_items': 0}
        
    total_pages = (cache_size + per_page - 1) // per_page
    return {
        'total_pages': total_pages,
        'total_items': cache_size
    }


def validate_email_address(email: str) -> bool:
    """
    Validate email address format.
    
    Args:
        email: Email address to validate
        
    Returns:
        bool: True if valid, False otherwise
    """
    import re
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email.strip()))


def sanitize_search_term(search_term: str) -> str:
    """
    Sanitize search term to prevent DASL injection.
    
    Args:
        search_term: Raw search term from user
        
    Returns:
        str: Sanitized search term
    """
    if not search_term:
        return ""
    
    # Remove potentially dangerous characters for DASL queries
    # Keep alphanumeric, spaces, and common punctuation
    sanitized = ''.join(c for c in search_term if c.isalnum() or c in ' .-_@')
    
    return sanitized.strip()