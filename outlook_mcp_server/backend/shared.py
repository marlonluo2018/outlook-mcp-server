"""Shared constants and cache management for email operations."""

# Standard library imports
import json
import os
import queue
import threading
import time
from datetime import datetime, timedelta
from typing import Any, Callable, Dict, List, Optional, Union

# Local application imports
from .config import cache_config, connection_config, performance_config
from .logging_config import get_logger

logger = get_logger(__name__)

# Cache base location
CACHE_BASE_DIR = cache_config.CACHE_BASE_DIR
CACHE_EXPIRY_HOURS = cache_config.CACHE_EXPIRY_HOURS
MAX_CACHE_SIZE = performance_config.MAX_CACHE_SIZE
BATCH_SAVE_SIZE = cache_config.BATCH_SAVE_SIZE
CACHE_SAVE_INTERVAL = cache_config.CACHE_SAVE_INTERVAL

# Global cache storage
email_cache = {}

# Email cache insertion order tracking
email_cache_order = []

# Cache save management
_cache_save_thread = None
_cache_save_queue = queue.Queue()
_last_cache_save_time = 0
_cache_save_lock = threading.Lock()


def _get_cache_file() -> str:
    """Get the cache file path."""
    return os.path.join(CACHE_BASE_DIR, "email_cache.json")


def _ensure_cache_dir_exists() -> None:
    """Ensure the cache directory exists."""
    if not os.path.exists(CACHE_BASE_DIR):
        os.makedirs(CACHE_BASE_DIR, exist_ok=True)


# Performance optimization: Cache parsed datetime objects to avoid repeated parsing
_email_time_cache: Dict[str, datetime] = {}

def _parse_email_time(received_time_str: str) -> datetime:
    """Parse email time with caching to avoid repeated parsing."""
    if received_time_str in _email_time_cache:
        return _email_time_cache[received_time_str]
    
    try:
        # Handle different datetime formats
        if 'T' in received_time_str:
            # ISO format: 2025-12-17T23:31:02.980000+00:00 or 2025-12-17 23:31:02.980000
            # Try to parse directly first (handles microseconds and timezone)
            try:
                parsed_time = datetime.fromisoformat(received_time_str)
            except ValueError:
                # If direct parsing fails, try removing microseconds
                if '.' in received_time_str:
                    parts = received_time_str.split('.')
                    # Keep the timezone part if present
                    if '+' in parts[1]:
                        time_part, tz_part = parts[1].split('+', 1)
                        received_time_str_clean = parts[0] + '+' + tz_part
                    elif '-' in parts[1]:
                        time_part, tz_part = parts[1].split('-', 1)
                        received_time_str_clean = parts[0] + '-' + tz_part
                    else:
                        received_time_str_clean = parts[0]
                    parsed_time = datetime.fromisoformat(received_time_str_clean)
                else:
                    parsed_time = datetime.fromisoformat(received_time_str)
        else:
            # Try other formats
            parsed_time = datetime.strptime(received_time_str, "%m/%d/%y %H:%M:%S")
            # Assume UTC for non-ISO formats
            from datetime import timezone
            parsed_time = parsed_time.replace(tzinfo=timezone.utc)
    except (ValueError, TypeError):
        parsed_time = datetime.min
    
    # Cache the result with the original string as key
    _email_time_cache[received_time_str] = parsed_time
    return parsed_time


def add_email_to_cache(email_id: str, email_data: Dict[str, Any]) -> None:
    """Add an email to the cache with size management, sorted by received time.

    Args:
        email_id: The unique identifier for email
        email_data: The email data to store in the cache
    """
    global email_cache, email_cache_order, _email_time_cache

    # If email already exists, remove it from order list first
    if email_id in email_cache:
        try:
            email_cache_order.remove(email_id)
        except ValueError:
            pass

    # Add to cache
    email_cache[email_id] = email_data
    
    # Parse received time once and cache it
    received_time_str = email_data.get("received_time", "")
    email_received_time = _parse_email_time(received_time_str)
    
    # Use binary search for insertion if list is large
    if len(email_cache_order) > performance_config.BINARY_SEARCH_THRESHOLD:  # Use binary search for larger lists
        import bisect
        
        # Create a list of timestamps for binary search (most recent = largest timestamp)
        # We use negative timestamps so bisect works correctly (most recent first)
        timestamps = []
        for id in email_cache_order:
            try:
                timestamp = -_parse_email_time(email_cache.get(id, {}).get("received_time", "")).timestamp()
                timestamps.append(timestamp)
            except (AttributeError, OSError) as e:
                # Skip problematic timestamps
                timestamps.append(float('-inf'))  # Put at the end
        
        try:
            insert_pos = bisect.bisect_left(timestamps, -email_received_time.timestamp())
        except (AttributeError, OSError) as e:
            # Fallback to linear search if timestamp calculation fails
            insert_pos = len(email_cache_order)
    else:
        # Use linear search for small lists
        insert_pos = len(email_cache_order)
        for i, existing_id in enumerate(email_cache_order):
            try:
                existing_time = _parse_email_time(email_cache.get(existing_id, {}).get("received_time", ""))
                if email_received_time > existing_time:  # More recent emails come first
                    insert_pos = i
                    break
            except (AttributeError, OSError):
                # Skip problematic emails
                continue
    
    email_cache_order.insert(insert_pos, email_id)

    # Enforce cache size limit - remove oldest entries if over limit
    while len(email_cache) > MAX_CACHE_SIZE:
        oldest_id = email_cache_order.pop(-1)  # Remove oldest from the end (least recent)
        oldest_email_data = email_cache.pop(oldest_id, None)  # Remove from cache
        
        # Clean up time cache entry for the removed email
        if oldest_email_data:
            oldest_received_time_str = oldest_email_data.get("received_time", "")
            if oldest_received_time_str in _email_time_cache:
                del _email_time_cache[oldest_received_time_str]


def clear_email_cache() -> None:
    """Clear the email cache both in memory and on disk."""
    global email_cache, email_cache_order, _email_time_cache

    # Clear in-memory cache
    email_cache.clear()
    email_cache_order.clear()
    _email_time_cache.clear()  # Clear time cache as well

    # Clear disk cache
    try:
        cache_file = _get_cache_file()
        if os.path.exists(cache_file):
            os.remove(cache_file)
    except Exception as e:
        import logging

        logger = logging.getLogger(__name__)
        logger.warning(f"Failed to clear email cache from disk: {e}")


def clear_cache() -> None:
    """Clear the email cache both in memory and on disk (deprecated - use clear_email_cache)."""
    clear_email_cache()


def get_cache_size() -> int:
    """Get the current size of the email cache.
    
    Returns:
        int: Number of emails in the cache
    """
    return len(email_cache)


def get_cache_stats() -> Dict[str, Any]:
    """Get statistics about the email cache.
    
    Returns:
        dict: Cache statistics including total emails, oldest and newest email times
    """
    oldest_email = None
    newest_email = None
    
    if email_cache_order:
        try:
            oldest_id = email_cache_order[-1]
            newest_id = email_cache_order[0]
            oldest_email = email_cache.get(oldest_id, {}).get("received_time", "")
            newest_email = email_cache.get(newest_id, {}).get("received_time", "")
        except (IndexError, KeyError):
            pass
    
    return {
        "total_emails": len(email_cache),
        "cache_size": len(email_cache),
        "cache_order_size": len(email_cache_order),
        "time_cache_size": len(_email_time_cache),
        "oldest_email": oldest_email,
        "newest_email": newest_email,
    }


def cleanup_cache() -> None:
    """Clean up the email cache by removing expired entries."""
    global email_cache, email_cache_order, _email_time_cache
    
    from datetime import timezone
    current_time = datetime.now(timezone.utc)
    expiry_threshold = current_time - timedelta(hours=CACHE_EXPIRY_HOURS)
    
    # Find expired emails
    expired_ids = []
    for email_id in email_cache_order:
        try:
            received_time_str = email_cache.get(email_id, {}).get("received_time", "")
            if received_time_str:
                received_time = _parse_email_time(received_time_str)
                if received_time < expiry_threshold:
                    expired_ids.append(email_id)
        except (ValueError, TypeError):
            # Skip problematic emails
            continue
    
    # Remove expired emails
    for email_id in expired_ids:
        email_cache_order.remove(email_id)
        email_data = email_cache.pop(email_id, None)
        if email_data:
            received_time_str = email_data.get("received_time", "")
            if received_time_str in _email_time_cache:
                del _email_time_cache[received_time_str]
    
    if expired_ids:
        logger.info(f"Cleaned up {len(expired_ids)} expired emails from cache")


def get_emails_by_date_range(start_date: Union[datetime, str], end_date: Union[datetime, str]) -> List[Dict[str, Any]]:
    """Get emails within a specific date range.
    
    Args:
        start_date: Start date of the range (datetime or ISO string)
        end_date: End date of the range (datetime or ISO string)
        
    Returns:
        list: List of email data dictionaries within the date range
    """
    from datetime import timezone, timedelta
    
    # Convert string dates to datetime if needed
    if isinstance(start_date, str):
        try:
            start_date = datetime.fromisoformat(start_date)
        except ValueError:
            return []
    
    if isinstance(end_date, str):
        try:
            end_date = datetime.fromisoformat(end_date)
        except ValueError:
            return []
    
    # Ensure timezone-aware comparison
    if start_date.tzinfo is None:
        start_date = start_date.replace(tzinfo=timezone.utc)
    if end_date.tzinfo is None:
        end_date = end_date.replace(tzinfo=timezone.utc)
    
    # Add a small buffer to handle timing edge cases (10 seconds)
    start_date = start_date - timedelta(seconds=10)
    end_date = end_date + timedelta(seconds=10)
    
    result = []
    for email_id in email_cache_order:
        try:
            received_time_str = email_cache.get(email_id, {}).get("received_time", "")
            if received_time_str:
                received_time = _parse_email_time(received_time_str)
                # Ensure received_time is timezone-aware
                if received_time.tzinfo is None:
                    received_time = received_time.replace(tzinfo=timezone.utc)
                if start_date <= received_time <= end_date:
                    result.append(email_cache[email_id])
        except (ValueError, TypeError):
            continue
    return result


def get_emails_by_sender(sender: str) -> List[Dict[str, Any]]:
    """Get emails from a specific sender.
    
    Args:
        sender: Sender name or email address to filter by
        
    Returns:
        list: List of email data dictionaries from the specified sender
    """
    result = []
    sender_lower = sender.lower()
    for email_id in email_cache_order:
        try:
            email_data = email_cache.get(email_id, {})
            # Check both "sender" and "from" fields for compatibility
            from_field = email_data.get("from", "") or email_data.get("sender", "")
            if from_field and sender_lower in from_field.lower():
                result.append(email_data)
        except (ValueError, TypeError):
            continue
    return result


def get_emails_by_subject(subject: str) -> List[Dict[str, Any]]:
    """Get emails with a specific subject.
    
    Args:
        subject: Subject text to filter by
        
    Returns:
        list: List of email data dictionaries matching the subject
    """
    result = []
    subject_lower = subject.lower()
    for email_id in email_cache_order:
        try:
            email_data = email_cache.get(email_id, {})
            email_subject = email_data.get("subject", "")
            if email_subject and subject_lower in email_subject.lower():
                result.append(email_data)
        except (ValueError, TypeError):
            continue
    return result


def get_emails_by_date_range_cached(start_date: Union[datetime, str], end_date: Union[datetime, str]) -> List[Dict[str, Any]]:
    """Get emails within a specific date range (cached version).
    
    Args:
        start_date: Start date of the range (datetime or ISO string)
        end_date: End date of the range (datetime or ISO string)
        
    Returns:
        list: List of email data dictionaries within the date range
    """
    return get_emails_by_date_range(start_date, end_date)


def get_emails_by_sender_cached(sender: str) -> List[Dict[str, Any]]:
    """Get emails from a specific sender (cached version).
    
    Args:
        sender: Sender name or email address to filter by
        
    Returns:
        list: List of email data dictionaries from the specified sender
    """
    return get_emails_by_sender(sender)


def get_emails_by_subject_cached(subject: str) -> List[Dict[str, Any]]:
    """Get emails with a specific subject (cached version).
    
    Args:
        subject: Subject text to filter by
        
    Returns:
        list: List of email data dictionaries matching the subject
    """
    return get_emails_by_subject(subject)


def _async_cache_saver() -> None:
    """Background thread for saving cache to disk."""
    while True:
        try:
            # Wait for save request
            force_save = _cache_save_queue.get(timeout=1.0)
            
            # Check if we should save (respect minimum interval)
            current_time = time.time()
            with _cache_save_lock:
                if not force_save and (current_time - _last_cache_save_time) < CACHE_SAVE_INTERVAL:
                    continue
                    
                if email_cache:  # Only save if there's data
                    cache_data = {
                        "cache": email_cache,
                        "cache_order": email_cache_order,
                        "timestamp": datetime.now().isoformat(),
                    }
                    
                    # Use temporary file for atomic write
                    cache_file = _get_cache_file()
                    temp_file = cache_file + '.tmp'
                    
                    with open(temp_file, "w", encoding="utf-8") as f:
                        json.dump(cache_data, f, ensure_ascii=False)
                    
                    # Atomic rename
                    os.replace(temp_file, cache_file)
                    
                    # Update last save time
                    globals()['_last_cache_save_time'] = current_time
                    
        except queue.Empty:
            continue
        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.warning(f"Failed to save email cache in background: {e}")


def save_email_cache(force_save: bool = False) -> None:
    """Save the email cache to disk with optimized asynchronous batching.
    
    Args:
        force_save: If True, save immediately regardless of batch size
    """
    global _cache_save_thread, _last_cache_save_time
    
    try:
        # Initialize counter if not exists
        if not hasattr(save_email_cache, '_pending_save_count'):
            save_email_cache._pending_save_count = 0
        
        # PERFORMANCE OPTIMIZATION: Only save if we have cache data
        if email_cache:
            save_email_cache._pending_save_count += 1
            
            # UVX COMPATIBILITY: Save more frequently for UVX environments
            # Reduce batch size for UVX to ensure data persistence
            uvx_batch_size = BATCH_SAVE_SIZE // 2  # Half of normal batch size
            
            # Save only if forced or reached batch size
            if force_save or save_email_cache._pending_save_count >= uvx_batch_size:
                # PERFORMANCE OPTIMIZATION: Check time interval to avoid too frequent saves
                current_time = time.time()
                if not force_save and (current_time - _last_cache_save_time) < CACHE_SAVE_INTERVAL:
                    # Skip this save to reduce I/O
                    return
                
                # Start background saver thread if not running
                if _cache_save_thread is None or not _cache_save_thread.is_alive():
                    _cache_save_thread = threading.Thread(target=_async_cache_saver, daemon=True)
                    _cache_save_thread.start()
                
                # Queue save request
                _cache_save_queue.put(force_save)
                
                # PERFORMANCE OPTIMIZATION: Update last save time immediately
                # This prevents multiple saves in quick succession
                _last_cache_save_time = current_time
                
                # Reset counter after queuing save
                save_email_cache._pending_save_count = 0
                
    except Exception as e:
        import logging
        logger = logging.getLogger(__name__)
        logger.warning(f"Failed to queue email cache save: {e}")


def load_email_cache() -> None:
    """Load the email cache from disk if it exists and is not expired."""
    global email_cache, email_cache_order
    try:
        cache_file = _get_cache_file()
        if not os.path.exists(cache_file):
            return

        with open(cache_file, "r", encoding="utf-8") as f:
            cache_data = json.load(f)

        # Check if cache is expired
        cache_timestamp = datetime.fromisoformat(cache_data.get("timestamp", "2000-01-01T00:00:00"))
        if datetime.now() - cache_timestamp > timedelta(hours=CACHE_EXPIRY_HOURS):
            return

        # Load the cache
        if isinstance(cache_data.get("cache"), dict):
            email_cache = cache_data["cache"]

            # Load cache order if available, otherwise rebuild it from keys
            if isinstance(cache_data.get("cache_order"), list):
                email_cache_order = cache_data["cache_order"]
                # Ensure order list only contains keys that exist in cache
                email_cache_order = [id for id in email_cache_order if id in email_cache]
            else:
                # Fallback: use cache keys (order not preserved)
                email_cache_order = list(email_cache.keys())
            
            import logging
            logger = logging.getLogger(__name__)
            logger.info(f"Loaded {len(email_cache)} emails from persistent cache")
        else:
            # Initialize empty cache if data is invalid
            email_cache = {}
            email_cache_order = []
    except Exception as e:
        import logging

        logger = logging.getLogger(__name__)
        logger.warning(f"Failed to load email cache: {e}")
        # Initialize empty cache on error
        email_cache = {}
        email_cache_order = []


def get_email_from_cache(email_identifier: Union[int, str]) -> Optional[Dict[str, Any]]:
    """Get an email from cache by its cache number (1-based) or email ID.

    Args:
        email_identifier: Either the 1-based position of email in the cache (int)
                          or the email ID string (str)

    Returns:
        The email data dictionary, or None if not found

    Raises:
        ValueError: If email_number is out of range
    """
    global email_cache, email_cache_order

    # Handle email_id (string) case
    if isinstance(email_identifier, str):
        return email_cache.get(email_identifier)

    # Handle email_number (int) case
    email_number = email_identifier

    # Check if email_number is within valid range
    if email_number < 1 or email_number > len(email_cache_order):
        raise ValueError(
            f"Email number {email_number} is out of range. Available range: 1-{len(email_cache_order)}"
        )

    # Convert to 0-based index and get the email_id
    email_id = email_cache_order[email_number - 1]

    # Return the email data
    return email_cache.get(email_id)


def immediate_save_cache() -> None:
    """Immediately save the email cache to disk for UVX compatibility.
    
    This function ensures cache persistence between UVX process instances
    by bypassing the batching mechanism and writing directly to disk.
    """
    global email_cache, email_cache_order
    
    if not email_cache:
        return
        
    try:
        # Use temporary file for atomic write
        cache_file = _get_cache_file()
        temp_file = cache_file + '.tmp'
        
        cache_data = {
            "cache": email_cache,
            "cache_order": email_cache_order,
            "timestamp": datetime.now().isoformat(),
        }
        
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(cache_data, f, ensure_ascii=False)
        
        # Atomic rename
        os.replace(temp_file, cache_file)
        
        import logging
        logger = logging.getLogger(__name__)
        logger.info(f"Immediately saved {len(email_cache)} emails to cache")
        
    except Exception as e:
        import logging
        logger = logging.getLogger(__name__)
        logger.warning(f"Failed to immediately save email cache: {e}")


def refresh_email_cache_with_new_data() -> bool:
    """Improved cache loading workflow:
    1. Clear both memory and disk cache
    2. Load fresh data into memory
    3. Immediately save to disk once data is loaded
    
    This ensures cache consistency and prevents stale data issues.
    
    Returns:
        bool: True if cache refresh was successful, False otherwise
    """
    try:
        # Step 1: Clear both memory and disk cache for fresh start
        clear_email_cache()
        
        # Step 2: Data loading happens externally - this function just prepares the cache
        # The actual email loading should be done by the caller
        
        # Step 3: Once data is loaded, save immediately (done by caller)
        
        logger = logging.getLogger(__name__)
        logger.info("Cache refreshed - ready for fresh data loading")
        return True
        
    except Exception as e:
        logger = logging.getLogger(__name__)
        logger.error(f"Failed to refresh email cache: {e}")
        return False


# Load cache when module is imported
load_email_cache()


# Logging configuration
def configure_logging():
    import logging
    import sys

    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    # Remove any existing handlers to avoid duplicates
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # Create console handler with formatting
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger