import pytest
import os
from outlook_mcp_server.backend.config import (
    CacheConfig,
    ConnectionConfig,
    PerformanceConfig,
    DisplayConfig,
    BatchConfig,
    OutlookConfig,
    EmailFormatConfig,
    AttachmentConfig,
    EmailMetadataConfig,
    ValidationConfig,
    cache_config,
    connection_config,
    performance_config,
    display_config,
    batch_config,
    outlook_config,
    email_format_config,
    attachment_config,
    email_metadata_config,
    validation_config
)


class TestCacheConfig:
    """Test suite for CacheConfig."""

    def test_max_days(self):
        """Test MAX_DAYS constant."""
        assert CacheConfig.MAX_DAYS == 30

    def test_max_emails(self):
        """Test MAX_EMAILS constant."""
        assert CacheConfig.MAX_EMAILS == 1000

    def test_max_load_time(self):
        """Test MAX_LOAD_TIME constant."""
        assert CacheConfig.MAX_LOAD_TIME == 58

    def test_lazy_loading_enabled(self):
        """Test LAZY_LOADING_ENABLED constant."""
        assert CacheConfig.LAZY_LOADING_ENABLED is True

    def test_cache_expiry_hours(self):
        """Test CACHE_EXPIRY_HOURS constant."""
        assert CacheConfig.CACHE_EXPIRY_HOURS == 6

    def test_batch_save_size(self):
        """Test BATCH_SAVE_SIZE constant."""
        assert CacheConfig.BATCH_SAVE_SIZE == 200

    def test_cache_save_interval(self):
        """Test CACHE_SAVE_INTERVAL constant."""
        assert CacheConfig.CACHE_SAVE_INTERVAL == 15.0

    def test_cache_base_dir(self):
        """Test CACHE_BASE_DIR property."""
        base_dir = cache_config.CACHE_BASE_DIR
        assert isinstance(base_dir, str)
        assert "outlook_mcp_server" in base_dir
        assert os.path.exists(os.path.dirname(base_dir))


class TestConnectionConfig:
    """Test suite for ConnectionConfig."""

    def test_max_retries(self):
        """Test MAX_RETRIES constant."""
        assert ConnectionConfig.MAX_RETRIES == 3

    def test_retry_delay(self):
        """Test RETRY_DELAY constant."""
        assert ConnectionConfig.RETRY_DELAY == 1.0

    def test_connection_timeout(self):
        """Test CONNECTION_TIMEOUT constant."""
        assert ConnectionConfig.CONNECTION_TIMEOUT == 30

    def test_heartbeat_interval(self):
        """Test HEARTBEAT_INTERVAL constant."""
        assert ConnectionConfig.HEARTBEAT_INTERVAL == 60


class TestPerformanceConfig:
    """Test suite for PerformanceConfig."""

    def test_max_cache_size(self):
        """Test MAX_CACHE_SIZE constant."""
        assert PerformanceConfig.MAX_CACHE_SIZE == 1000

    def test_cache_cleanup_threshold(self):
        """Test CACHE_CLEANUP_THRESHOLD constant."""
        assert PerformanceConfig.CACHE_CLEANUP_THRESHOLD == 0.8

    def test_lazy_load_batch_size(self):
        """Test LAZY_LOAD_BATCH_SIZE constant."""
        assert PerformanceConfig.LAZY_LOAD_BATCH_SIZE == 100

    def test_max_concurrent_operations(self):
        """Test MAX_CONCURRENT_OPERATIONS constant."""
        assert PerformanceConfig.MAX_CONCURRENT_OPERATIONS == 5


class TestDisplayConfig:
    """Test suite for DisplayConfig."""

    def test_max_subject_length(self):
        """Test MAX_SUBJECT_LENGTH constant."""
        assert DisplayConfig.MAX_SUBJECT_LENGTH == 100

    def test_max_sender_length(self):
        """Test MAX_SENDER_LENGTH constant."""
        assert DisplayConfig.MAX_SENDER_LENGTH == 50

    def test_preview_length(self):
        """Test PREVIEW_LENGTH constant."""
        assert DisplayConfig.PREVIEW_LENGTH == 200

    def test_date_format(self):
        """Test DATE_FORMAT constant."""
        assert DisplayConfig.DATE_FORMAT == "%Y-%m-%d %H:%M:%S"

    def test_separator_line(self):
        """Test SEPARATOR_LINE constant."""
        assert DisplayConfig.SEPARATOR_LINE == "=" * 60


class TestBatchConfig:
    """Test suite for BatchConfig."""

    def test_max_batch_size(self):
        """Test MAX_BATCH_SIZE constant."""
        assert BatchConfig.MAX_BATCH_SIZE == 100

    def test_max_email_number(self):
        """Test MAX_EMAIL_NUMBER constant."""
        assert BatchConfig.MAX_EMAIL_NUMBER == 2000

    def test_max_page_number(self):
        """Test MAX_PAGE_NUMBER constant."""
        assert BatchConfig.MAX_PAGE_NUMBER == 100


class TestOutlookConfig:
    """Test suite for OutlookConfig."""

    def test_ol_mail_item(self):
        """Test OL_MAIL_ITEM constant."""
        assert OutlookConfig.OL_MAIL_ITEM == 0

    def test_ol_folder_inbox(self):
        """Test OL_FOLDER_INBOX constant."""
        assert OutlookConfig.OL_FOLDER_INBOX == 6

    def test_ol_folder_sent(self):
        """Test OL_FOLDER_SENT constant."""
        assert OutlookConfig.OL_FOLDER_SENT == 5

    def test_ol_folder_drafts(self):
        """Test OL_FOLDER_DRAFTS constant."""
        assert OutlookConfig.OL_FOLDER_DRAFTS == 16

    def test_ol_folder_deleted(self):
        """Test OL_FOLDER_DELETED constant."""
        assert OutlookConfig.OL_FOLDER_DELETED == 3


class TestEmailFormatConfig:
    """Test suite for EmailFormatConfig."""

    def test_plain_text(self):
        """Test PLAIN_TEXT constant."""
        assert EmailFormatConfig.PLAIN_TEXT == 1

    def test_html(self):
        """Test HTML constant."""
        assert EmailFormatConfig.HTML == 2

    def test_rich_text(self):
        """Test RICH_TEXT constant."""
        assert EmailFormatConfig.RICH_TEXT == 3


class TestAttachmentConfig:
    """Test suite for AttachmentConfig."""

    def test_by_value(self):
        """Test BY_VALUE constant."""
        assert AttachmentConfig.BY_VALUE == 1

    def test_by_reference(self):
        """Test BY_REFERENCE constant."""
        assert AttachmentConfig.BY_REFERENCE == 4

    def test_embedding(self):
        """Test EMBEDDING constant."""
        assert AttachmentConfig.EMBEDDING == 5

    def test_ole(self):
        """Test OLE constant."""
        assert AttachmentConfig.OLE == 6


class TestEmailMetadataConfig:
    """Test suite for EmailMetadataConfig."""

    def test_importance_low(self):
        """Test IMPORTANCE_LOW constant."""
        assert EmailMetadataConfig.IMPORTANCE_LOW == 0

    def test_importance_normal(self):
        """Test IMPORTANCE_NORMAL constant."""
        assert EmailMetadataConfig.IMPORTANCE_NORMAL == 1

    def test_importance_high(self):
        """Test IMPORTANCE_HIGH constant."""
        assert EmailMetadataConfig.IMPORTANCE_HIGH == 2

    def test_sensitivity_normal(self):
        """Test SENSITIVITY_NORMAL constant."""
        assert EmailMetadataConfig.SENSITIVITY_NORMAL == 0

    def test_sensitivity_personal(self):
        """Test SENSITIVITY_PERSONAL constant."""
        assert EmailMetadataConfig.SENSITIVITY_PERSONAL == 1

    def test_sensitivity_private(self):
        """Test SENSITIVITY_PRIVATE constant."""
        assert EmailMetadataConfig.SENSITIVITY_PRIVATE == 2

    def test_sensitivity_confidential(self):
        """Test SENSITIVITY_CONFIDENTIAL constant."""
        assert EmailMetadataConfig.SENSITIVITY_CONFIDENTIAL == 3

    def test_flag_no_flag(self):
        """Test FLAG_NO_FLAG constant."""
        assert EmailMetadataConfig.FLAG_NO_FLAG == 0

    def test_flag_flagged(self):
        """Test FLAG_FLAGGED constant."""
        assert EmailMetadataConfig.FLAG_FLAGGED == 1

    def test_flag_completed(self):
        """Test FLAG_COMPLETED constant."""
        assert EmailMetadataConfig.FLAG_COMPLETED == 2


class TestValidationConfig:
    """Test suite for ValidationConfig."""

    def test_max_email_length(self):
        """Test MAX_EMAIL_LENGTH constant."""
        assert ValidationConfig.MAX_EMAIL_LENGTH == 254

    def test_max_email_local_part_length(self):
        """Test MAX_EMAIL_LOCAL_PART_LENGTH constant."""
        assert ValidationConfig.MAX_EMAIL_LOCAL_PART_LENGTH == 64

    def test_max_search_term_length(self):
        """Test MAX_SEARCH_TERM_LENGTH constant."""
        assert ValidationConfig.MAX_SEARCH_TERM_LENGTH == 100

    def test_max_folder_name_length(self):
        """Test MAX_FOLDER_NAME_LENGTH constant."""
        assert ValidationConfig.MAX_FOLDER_NAME_LENGTH == 100

    def test_min_search_term_length(self):
        """Test MIN_SEARCH_TERM_LENGTH constant."""
        assert ValidationConfig.MIN_SEARCH_TERM_LENGTH == 1
