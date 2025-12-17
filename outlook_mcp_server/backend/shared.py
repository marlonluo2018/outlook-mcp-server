from typing import Optional

# Global configuration constants
MAX_DAYS = 30
MAX_EMAILS = 1000
MAX_LOAD_TIME = 58  # seconds
LAZY_LOADING_ENABLED = True  # Enable lazy loading for email details

# Connection configuration
CONNECT_TIMEOUT = 30  # seconds
MAX_RETRIES = 3
INITIAL_BACKOFF = 1  # seconds
MAX_BACKOFF = 16  # seconds

# Cache configuration
import os
import json
import threading
import queue
import time
from datetime import datetime, timedelta

# Cache base location
CACHE_BASE_DIR = os.path.join(
    os.getenv("LOCALAPPDATA", os.path.expanduser("~")), "outlook_mcp_server"
)
CACHE_EXPIRY_HOURS = 6  # Cache expires after 6 hours
MAX_CACHE_SIZE = 1000  # Maximum number of emails to keep in cache
BATCH_SAVE_SIZE = 200  # Increased batch size for better performance (reduced I/O operations)
CACHE_SAVE_INTERVAL = 15.0  # Increased interval to reduce disk I/O for better performance

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


def _ensure_cache_dir_exists():
    """Ensure the cache directory exists."""
    if not os.path.exists(CACHE_BASE_DIR):
        os.makedirs(CACHE_BASE_DIR, exist_ok=True)


# Performance optimization: Cache parsed datetime objects to avoid repeated parsing
_email_time_cache = {}

def _parse_email_time(received_time_str: str):
    """Parse email time with caching to avoid repeated parsing."""
    if received_time_str in _email_time_cache:
        return _email_time_cache[received_time_str]
    
    try:
        # Handle different datetime formats
        if 'T' in received_time_str:
            # ISO format: 2025-12-17 23:31:02.980000
            # Remove microseconds if present
            if '.' in received_time_str:
                parts = received_time_str.split('.')
                received_time_str = parts[0]
            parsed_time = datetime.fromisoformat(received_time_str)
        else:
            # Try other formats
            parsed_time = datetime.strptime(received_time_str, "%m/%d/%y %H:%M:%S")
    except (ValueError, TypeError):
        parsed_time = datetime.min
    
    # Cache the result
    _email_time_cache[received_time_str] = parsed_time
    return parsed_time


def add_email_to_cache(email_id: str, email_data: dict):
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
    if len(email_cache_order) > 20:  # Use binary search for larger lists
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
        del email_cache[oldest_id]  # Remove from cache
        
        # Note: Time cache cleanup is handled when email_data is retrieved before deletion
        # No additional cleanup needed here


def clear_email_cache():
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


def _async_cache_saver():
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


def save_email_cache(force_save=False):
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


def load_email_cache():
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


def get_email_from_cache(email_number: int) -> Optional[dict]:
    """Get an email from cache by its cache number (1-based).

    Args:
        email_number: The 1-based position of email in the cache

    Returns:
        The email data dictionary, or None if not found

    Raises:
        ValueError: If email_number is out of range
    """
    global email_cache, email_cache_order

    # Check if email_number is within valid range
    if email_number < 1 or email_number > len(email_cache_order):
        raise ValueError(
            f"Email number {email_number} is out of range. Available range: 1-{len(email_cache_order)}"
        )

    # Convert to 0-based index and get the email_id
    email_id = email_cache_order[email_number - 1]

    # Return the email data
    return email_cache.get(email_id)


def immediate_save_cache():
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