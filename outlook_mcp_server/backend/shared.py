# Global configuration constants
MAX_DAYS = 30
MAX_EMAILS = 1000
MAX_LOAD_TIME = 58  # seconds

# Connection configuration
CONNECT_TIMEOUT = 30  # seconds
MAX_RETRIES = 3
INITIAL_BACKOFF = 1  # seconds
MAX_BACKOFF = 16  # seconds

# Cache configuration
import os
import json
from datetime import datetime, timedelta

# Cache base location
CACHE_BASE_DIR = os.path.join(os.getenv('LOCALAPPDATA', os.path.expanduser('~')), 'outlook_mcp_server')
CACHE_EXPIRY_HOURS = 1  # Cache expires after 1 hour
MAX_CACHE_SIZE = 1000  # Maximum number of emails to keep in cache

# Global cache storage
email_cache = {}

# Email cache insertion order tracking
email_cache_order = []

def _get_cache_file() -> str:
    """Get the cache file path."""
    return os.path.join(CACHE_BASE_DIR, 'email_cache.json')


def _ensure_cache_dir_exists():
    """Ensure the cache directory exists."""
    if not os.path.exists(CACHE_BASE_DIR):
        os.makedirs(CACHE_BASE_DIR, exist_ok=True)


def add_email_to_cache(email_id: str, email_data: dict):
    """Add an email to the cache with size management.
    
    Args:
        email_id: The unique identifier for the email
        email_data: The email data to store in the cache
    """
    global email_cache, email_cache_order
    
    # If email already exists, remove it from order list first
    if email_id in email_cache:
        try:
            email_cache_order.remove(email_id)
        except ValueError:
            pass
    
    # Add to cache and update order
    email_cache[email_id] = email_data
    email_cache_order.append(email_id)
    
    # Enforce cache size limit - remove oldest entries if over limit
    while len(email_cache) > MAX_CACHE_SIZE:
        oldest_id = email_cache_order.pop(0)  # Remove oldest from order list
        del email_cache[oldest_id]  # Remove from cache

def save_email_cache():
    """Save the email cache to disk."""
    try:
        _ensure_cache_dir_exists()
        
        # Only save if we have cache data
        if email_cache:
            cache_data = {
                'cache': email_cache,
                'cache_order': email_cache_order,
                'timestamp': datetime.now().isoformat()
            }
            
            with open(_get_cache_file(), 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        import logging
        logger = logging.getLogger(__name__)
        logger.warning(f"Failed to save email cache: {e}")

def load_email_cache():
    """Load the email cache from disk if it exists and is not expired."""
    global email_cache, email_cache_order
    try:
        cache_file = _get_cache_file()
        if not os.path.exists(cache_file):
            return
            
        with open(cache_file, 'r', encoding='utf-8') as f:
            cache_data = json.load(f)
            
        # Check if cache is expired
        cache_timestamp = datetime.fromisoformat(cache_data.get('timestamp', '2000-01-01T00:00:00'))
        if datetime.now() - cache_timestamp > timedelta(hours=CACHE_EXPIRY_HOURS):
            return
            
        # Load the cache
        if isinstance(cache_data.get('cache'), dict):
            email_cache = cache_data['cache']
            
            # Load cache order if available, otherwise rebuild it from keys
            if isinstance(cache_data.get('cache_order'), list):
                email_cache_order = cache_data['cache_order']
                # Ensure order list only contains keys that exist in cache
                email_cache_order = [id for id in email_cache_order if id in email_cache]
            else:
                # Fallback: use cache keys (order not preserved)
                email_cache_order = list(email_cache.keys())
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

def clear_email_cache():
    """Clear the email cache both in memory and on disk."""
    global email_cache, email_cache_order
    
    # Clear in-memory cache
    email_cache.clear()
    email_cache_order.clear()
    
    # Clear disk cache
    try:
        cache_file = _get_cache_file()
        if os.path.exists(cache_file):
            os.remove(cache_file)
    except Exception as e:
        import logging
        logger = logging.getLogger(__name__)
        logger.warning(f"Failed to clear email cache from disk: {e}")


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
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger