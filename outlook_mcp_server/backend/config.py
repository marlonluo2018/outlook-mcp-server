"""Centralized configuration for Outlook MCP Server.

This module consolidates all configuration constants to enable easy
customization and environment-specific settings.
"""

from typing import Optional
import os


class CacheConfig:
    """Cache configuration settings.
    
    These settings control how emails are cached in memory and on disk
    to improve performance and reduce Outlook COM interface calls.
    """

    MAX_DAYS = 30
    MAX_EMAILS = 1000
    MAX_LOAD_TIME = 58
    LAZY_LOADING_ENABLED = True

    CACHE_EXPIRY_HOURS = 6
    BATCH_SAVE_SIZE = 200
    CACHE_SAVE_INTERVAL = 15.0

    @property
    def CACHE_BASE_DIR(self) -> str:
        """Get cache base directory."""
        return os.path.join(
            os.getenv("LOCALAPPDATA", os.path.expanduser("~")),
            "outlook_mcp_server"
        )


class ConnectionConfig:
    """Connection configuration settings.
    
    These settings control how the application connects to Outlook
    and handles connection failures.
    """

    CONNECT_TIMEOUT = 30
    CONNECTION_TIMEOUT = 30
    MAX_RETRIES = 3
    RETRY_DELAY = 1.0
    INITIAL_BACKOFF = 1
    MAX_BACKOFF = 16
    HEARTBEAT_INTERVAL = 60


class PerformanceConfig:
    """Performance optimization settings.
    
    These settings control various performance optimizations including
    cache management, search algorithms, and concurrent operations.
    """

    MAX_CACHE_SIZE = 1000
    BINARY_SEARCH_THRESHOLD = 100
    CACHE_CLEANUP_THRESHOLD = 0.8
    LAZY_LOAD_BATCH_SIZE = 100
    MAX_CONCURRENT_OPERATIONS = 5


class DisplayConfig:
    """Display formatting settings.
    
    These settings control how email information is displayed to users,
    including text truncation and date formatting.
    """

    SEPARATOR_LINE_LENGTH = 60
    MAX_SUBJECT_LENGTH = 100
    MAX_SENDER_LENGTH = 50
    PREVIEW_LENGTH = 200
    DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
    SEPARATOR_LINE = "=" * 60


class BatchConfig:
    """Batch processing settings for Outlook operations.
    
    These settings control batch sizes for various operations to balance
    performance and resource usage when working with Outlook COM interface.
    
    Use Cases:
    - DEFAULT_BATCH_SIZE (50): Used for general batch processing when adding
      emails to cache from search results. Provides balanced performance for
      medium-sized datasets.
    
    - FAST_MODE_BATCH_SIZE (100): Used in fast mode email extraction (minimal
      metadata) when loading emails to cache. Processes more emails per batch
      since minimal data extraction is faster.
    
    - FULL_EXTRACTION_BATCH_SIZE (25): Used in full extraction mode (complete
      metadata including body, attachments) when loading emails to cache.
      Smaller batch size due to heavier processing per email.
    
    - OUTLOOK_BCC_LIMIT (500): Maximum number of BCC recipients per email
      in Outlook. Used for batch forwarding operations to split recipients
      into multiple emails.
    """

    MAX_BATCH_SIZE = 100
    MAX_EMAIL_NUMBER = 2000
    MAX_PAGE_NUMBER = 100
    DEFAULT_BATCH_SIZE = 50
    FAST_MODE_BATCH_SIZE = 100
    FULL_EXTRACTION_BATCH_SIZE = 25
    OUTLOOK_BCC_LIMIT = 500
    IMAGE_EMBEDDING_SIZE_THRESHOLD = 102400


class OutlookConfig:
    """Outlook COM configuration.
    
    These constants represent Outlook item types used when creating
    or identifying Outlook objects via COM interface.
    """

    OL_MAIL_ITEM = 0
    OL_CONTACT_ITEM = 2
    OL_DISTRIBUTION_LIST_ITEM = 7
    OL_JOURNAL_ITEM = 4
    OL_NOTE_ITEM = 5
    OL_POST_ITEM = 6
    OL_TASK_ITEM = 3
    OL_FOLDER_INBOX = 6
    OL_FOLDER_SENT = 5
    OL_FOLDER_DRAFTS = 16
    OL_FOLDER_DELETED = 3


class EmailFormatConfig:
    """Email format configuration.
    
    These constants represent different email body formats supported
    by Outlook for email composition and display.
    """

    PLAIN_TEXT = 1
    HTML = 2
    RICH_TEXT = 3
    OL_FORMAT_PLAIN = 1
    OL_FORMAT_HTML = 2
    OL_FORMAT_RICH_TEXT = 3


class AttachmentConfig:
    """Attachment configuration.
    
    These constants represent different attachment types in Outlook,
    including inline/embedded attachments and file references.
    """

    BY_VALUE = 1
    BY_REFERENCE = 4
    EMBEDDING = 5
    OLE = 6


class EmailMetadataConfig:
    """Email metadata configuration.
    
    These constants represent various metadata properties for emails,
    including importance levels, sensitivity settings, and flag status.
    """

    IMPORTANCE_LOW = 0
    IMPORTANCE_NORMAL = 1
    IMPORTANCE_HIGH = 2

    SENSITIVITY_NORMAL = 0
    SENSITIVITY_PERSONAL = 1
    SENSITIVITY_PRIVATE = 2
    SENSITIVITY_CONFIDENTIAL = 3

    FLAG_NO_FLAG = 0
    FLAG_FLAGGED = 1
    FLAG_COMPLETED = 2
    FLAG_STATUS_UNFLAGGED = 0
    FLAG_STATUS_FLAGGED = 1
    FLAG_STATUS_COMPLETE = 2


class ValidationConfig:
    """Validation configuration.
    
    These settings control validation rules for email addresses and
    other user inputs to ensure data integrity and prevent errors.
    """

    MAX_EMAIL_LENGTH = 254
    MAX_EMAIL_LOCAL_PART_LENGTH = 64
    MIN_EMAIL_LENGTH = 3
    MAX_SEARCH_TERM_LENGTH = 100
    MAX_FOLDER_NAME_LENGTH = 100
    MIN_SEARCH_TERM_LENGTH = 1


cache_config = CacheConfig()
connection_config = ConnectionConfig()
performance_config = PerformanceConfig()
display_config = DisplayConfig()
batch_config = BatchConfig()
outlook_config = OutlookConfig()
email_format_config = EmailFormatConfig()
attachment_config = AttachmentConfig()
email_metadata_config = EmailMetadataConfig()
validation_config = ValidationConfig()
