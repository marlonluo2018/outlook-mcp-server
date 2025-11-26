# Global configuration constants
MAX_DAYS = 30
MAX_EMAILS = 1000
MAX_LOAD_TIME = 58  # seconds

# Connection configuration
CONNECT_TIMEOUT = 30  # seconds
MAX_RETRIES = 3
INITIAL_BACKOFF = 1  # seconds
MAX_BACKOFF = 16  # seconds



# Global email cache dictionary (key=EntryID, value=formatted email)
email_cache = {}

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