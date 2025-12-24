import pytest
from outlook_mcp_server.backend.validation import (
    ValidationError,
    validate_search_term,
    validate_days_parameter,
    validate_folder_name,
    validate_email_address,
    validate_email_number,
    validate_page_parameter,
    validate_cache_available,
    validate_not_empty
)
from outlook_mcp_server.backend.config import (
    outlook_config,
    email_format_config,
    attachment_config,
    email_metadata_config,
    batch_config,
    performance_config,
    display_config
)


class TestValidationFunctions:
    """Test suite for validation functions."""

    def test_validate_search_term_valid(self):
        """Test validate_search_term with valid inputs."""
        assert validate_search_term("test") == "test"
        assert validate_search_term("test subject") == "test subject"
        assert validate_search_term("Test with numbers 123") == "Test with numbers 123"

    def test_validate_search_term_empty(self):
        """Test validate_search_term with empty string."""
        with pytest.raises(ValidationError, match="must be a non-empty string"):
            validate_search_term("")

    def test_validate_search_term_whitespace_only(self):
        """Test validate_search_term with whitespace only."""
        with pytest.raises(ValidationError, match="must be a non-empty string"):
            validate_search_term("   ")

    def test_validate_search_term_none(self):
        """Test validate_search_term with None."""
        with pytest.raises(ValidationError, match="must be a non-empty string"):
            validate_search_term(None)

    def test_validate_days_parameter_valid(self):
        """Test validate_days_parameter with valid inputs."""
        assert validate_days_parameter(1) == 1
        assert validate_days_parameter(7) == 7
        assert validate_days_parameter(30) == 30

    def test_validate_days_parameter_negative(self):
        """Test validate_days_parameter with negative value."""
        with pytest.raises(ValidationError, match="must be between"):
            validate_days_parameter(-1)

    def test_validate_days_parameter_zero(self):
        """Test validate_days_parameter with zero."""
        with pytest.raises(ValidationError, match="must be between"):
            validate_days_parameter(0)

    def test_validate_days_parameter_exceeds_max(self):
        """Test validate_days_parameter exceeding maximum."""
        with pytest.raises(ValidationError, match="must be between"):
            validate_days_parameter(31)

    def test_validate_folder_name_valid(self):
        """Test validate_folder_name with valid inputs."""
        assert validate_folder_name("Inbox") == "Inbox"
        assert validate_folder_name("Sent Items") == "Sent Items"
        assert validate_folder_name("Drafts") == "Drafts"

    def test_validate_folder_name_empty(self):
        """Test validate_folder_name with empty string."""
        result = validate_folder_name("")
        assert result is None

    def test_validate_folder_name_none(self):
        """Test validate_folder_name with None."""
        result = validate_folder_name(None)
        assert result is None

    def test_validate_email_address_valid(self):
        """Test validate_email_address with valid inputs."""
        assert validate_email_address("test@example.com") == "test@example.com"
        assert validate_email_address("user.name@example.com") == "user.name@example.com"
        assert validate_email_address("user+tag@example.com") == "user+tag@example.com"

    def test_validate_email_address_empty(self):
        """Test validate_email_address with empty string."""
        with pytest.raises(ValidationError, match="Email address must be a non-empty string"):
            validate_email_address("")

    def test_validate_email_address_none(self):
        """Test validate_email_address with None."""
        with pytest.raises(ValidationError, match="Email address must be a non-empty string"):
            validate_email_address(None)

    def test_validate_email_address_invalid_format(self):
        """Test validate_email_address with invalid format."""
        with pytest.raises(ValidationError, match="Invalid email address format"):
            validate_email_address("invalid-email")

    def test_validate_email_address_too_long(self):
        """Test validate_email_address exceeding maximum length."""
        long_email = "a" * 255 + "@example.com"
        with pytest.raises(ValidationError, match="Email address is too long"):
            validate_email_address(long_email)

    def test_validate_email_address_local_part_too_long(self):
        """Test validate_email_address with local part exceeding maximum length."""
        long_local = "a" * 65
        with pytest.raises(ValidationError, match="Email local part is too long"):
            validate_email_address(f"{long_local}@example.com")

    def test_validate_email_number_valid(self):
        """Test validate_email_number with valid inputs."""
        assert validate_email_number(1, 100) == 1
        assert validate_email_number(10, 100) == 10
        assert validate_email_number(100, 100) == 100

    def test_validate_email_number_negative(self):
        """Test validate_email_number with negative value."""
        with pytest.raises(ValidationError, match="is out of range"):
            validate_email_number(-1, 100)

    def test_validate_email_number_zero(self):
        """Test validate_email_number with zero."""
        with pytest.raises(ValidationError, match="is out of range"):
            validate_email_number(0, 100)

    def test_validate_email_number_exceeds_max(self):
        """Test validate_email_number exceeding maximum."""
        with pytest.raises(ValidationError, match="is out of range"):
            validate_email_number(101, 100)

    def test_validate_page_parameter_valid(self):
        """Test validate_page_parameter with valid inputs."""
        assert validate_page_parameter(1, 10) == 1
        assert validate_page_parameter(5, 10) == 5
        assert validate_page_parameter(10, 10) == 10

    def test_validate_page_parameter_negative(self):
        """Test validate_page_parameter with negative value."""
        with pytest.raises(ValidationError, match="Page parameter must be at least 1"):
            validate_page_parameter(-1, 10)

    def test_validate_page_parameter_zero(self):
        """Test validate_page_parameter with zero."""
        with pytest.raises(ValidationError, match="Page parameter must be at least 1"):
            validate_page_parameter(0, 10)

    def test_validate_page_parameter_exceeds_max(self):
        """Test validate_page_parameter exceeding maximum."""
        with pytest.raises(ValidationError, match="is out of range"):
            validate_page_parameter(11, 10)

    def test_validate_cache_available_with_cache(self):
        """Test validate_cache_available when cache is available."""
        from outlook_mcp_server.backend.shared import email_cache
        email_cache["test_id"] = {"subject": "test"}
        result = validate_cache_available(1)
        assert result is None
        del email_cache["test_id"]

    def test_validate_cache_available_empty(self):
        """Test validate_cache_available when cache is empty."""
        from outlook_mcp_server.backend.shared import email_cache
        email_cache.clear()
        with pytest.raises(ValidationError, match="No emails available"):
            validate_cache_available(0)

    def test_validate_not_empty_valid(self):
        """Test validate_not_empty with valid inputs."""
        assert validate_not_empty("test", "Test field") == "test"
        assert validate_not_empty("  test  ", "Test field") == "test"

    def test_validate_not_empty_empty(self):
        """Test validate_not_empty with empty string."""
        with pytest.raises(ValidationError, match="must not be empty"):
            validate_not_empty("", "Test field")

    def test_validate_not_empty_whitespace_only(self):
        """Test validate_not_empty with whitespace only."""
        with pytest.raises(ValidationError, match="must not be empty"):
            validate_not_empty("   ", "Test field")

    def test_validate_not_empty_none(self):
        """Test validate_not_empty with None."""
        with pytest.raises(ValidationError, match="must be a non-empty string"):
            validate_not_empty(None, "Test field")


class TestOutlookConfigConstants:
    """Test suite for Outlook config constants."""

    def test_ol_mail_item(self):
        """Test OL_MAIL_ITEM constant."""
        assert outlook_config.OL_MAIL_ITEM == 0

    def test_ol_contact_item(self):
        """Test OL_CONTACT_ITEM constant."""
        assert outlook_config.OL_CONTACT_ITEM == 2

    def test_ol_journal_item(self):
        """Test OL_JOURNAL_ITEM constant."""
        assert outlook_config.OL_JOURNAL_ITEM == 4

    def test_ol_note_item(self):
        """Test OL_NOTE_ITEM constant."""
        assert outlook_config.OL_NOTE_ITEM == 5

    def test_ol_post_item(self):
        """Test OL_POST_ITEM constant."""
        assert outlook_config.OL_POST_ITEM == 6

    def test_ol_task_item(self):
        """Test OL_TASK_ITEM constant."""
        assert outlook_config.OL_TASK_ITEM == 3


class TestEmailFormatConfigConstants:
    """Test suite for EmailFormatConfig constants."""

    def test_ol_format_plain(self):
        """Test OL_FORMAT_PLAIN constant."""
        assert email_format_config.OL_FORMAT_PLAIN == 1

    def test_ol_format_html(self):
        """Test OL_FORMAT_HTML constant."""
        assert email_format_config.OL_FORMAT_HTML == 2

    def test_ol_format_rich_text(self):
        """Test OL_FORMAT_RICH_TEXT constant."""
        assert email_format_config.OL_FORMAT_RICH_TEXT == 3


class TestAttachmentConfigConstants:
    """Test suite for AttachmentConfig constants."""

    def test_by_value(self):
        """Test BY_VALUE constant."""
        assert attachment_config.BY_VALUE == 1

    def test_by_reference(self):
        """Test BY_REFERENCE constant."""
        assert attachment_config.BY_REFERENCE == 4

    def test_embedded(self):
        """Test EMBEDDING constant."""
        assert attachment_config.EMBEDDING == 5

    def test_ole(self):
        """Test OLE constant."""
        assert attachment_config.OLE == 6


class TestEmailMetadataConfigConstants:
    """Test suite for EmailMetadataConfig constants."""

    def test_importance_low(self):
        """Test IMPORTANCE_LOW constant."""
        assert email_metadata_config.IMPORTANCE_LOW == 0

    def test_importance_normal(self):
        """Test IMPORTANCE_NORMAL constant."""
        assert email_metadata_config.IMPORTANCE_NORMAL == 1

    def test_importance_high(self):
        """Test IMPORTANCE_HIGH constant."""
        assert email_metadata_config.IMPORTANCE_HIGH == 2

    def test_sensitivity_normal(self):
        """Test SENSITIVITY_NORMAL constant."""
        assert email_metadata_config.SENSITIVITY_NORMAL == 0

    def test_sensitivity_personal(self):
        """Test SENSITIVITY_PERSONAL constant."""
        assert email_metadata_config.SENSITIVITY_PERSONAL == 1

    def test_sensitivity_private(self):
        """Test SENSITIVITY_PRIVATE constant."""
        assert email_metadata_config.SENSITIVITY_PRIVATE == 2

    def test_sensitivity_confidential(self):
        """Test SENSITIVITY_CONFIDENTIAL constant."""
        assert email_metadata_config.SENSITIVITY_CONFIDENTIAL == 3

    def test_flag_status_unflagged(self):
        """Test FLAG_STATUS_UNFLAGGED constant."""
        assert email_metadata_config.FLAG_STATUS_UNFLAGGED == 0

    def test_flag_status_flagged(self):
        """Test FLAG_STATUS_FLAGGED constant."""
        assert email_metadata_config.FLAG_STATUS_FLAGGED == 1

    def test_flag_status_complete(self):
        """Test FLAG_STATUS_COMPLETE constant."""
        assert email_metadata_config.FLAG_STATUS_COMPLETE == 2


class TestBatchConfigConstants:
    """Test suite for BatchConfig constants."""

    def test_outlook_bcc_limit(self):
        """Test OUTLOOK_BCC_LIMIT constant."""
        assert batch_config.OUTLOOK_BCC_LIMIT == 500

    def test_image_embedding_size_threshold(self):
        """Test IMAGE_EMBEDDING_SIZE_THRESHOLD constant."""
        assert batch_config.IMAGE_EMBEDDING_SIZE_THRESHOLD == 102400

    def test_default_batch_size(self):
        """Test DEFAULT_BATCH_SIZE constant."""
        assert batch_config.DEFAULT_BATCH_SIZE == 50

    def test_fast_mode_batch_size(self):
        """Test FAST_MODE_BATCH_SIZE constant."""
        assert batch_config.FAST_MODE_BATCH_SIZE == 100

    def test_full_extraction_batch_size(self):
        """Test FULL_EXTRACTION_BATCH_SIZE constant."""
        assert batch_config.FULL_EXTRACTION_BATCH_SIZE == 25


class TestPerformanceConfigConstants:
    """Test suite for PerformanceConfig constants."""

    def test_binary_search_threshold(self):
        """Test BINARY_SEARCH_THRESHOLD constant."""
        assert performance_config.BINARY_SEARCH_THRESHOLD == 100

    def test_max_cache_size(self):
        """Test MAX_CACHE_SIZE constant."""
        assert performance_config.MAX_CACHE_SIZE == 1000

    def test_cache_cleanup_threshold(self):
        """Test CACHE_CLEANUP_THRESHOLD constant."""
        assert performance_config.CACHE_CLEANUP_THRESHOLD == 0.8

    def test_lazy_load_batch_size(self):
        """Test LAZY_LOAD_BATCH_SIZE constant."""
        assert performance_config.LAZY_LOAD_BATCH_SIZE == 100

    def test_max_concurrent_operations(self):
        """Test MAX_CONCURRENT_OPERATIONS constant."""
        assert performance_config.MAX_CONCURRENT_OPERATIONS == 5


class TestDisplayConfigConstants:
    """Test suite for DisplayConfig constants."""

    def test_separator_line_length(self):
        """Test SEPARATOR_LINE_LENGTH constant."""
        assert display_config.SEPARATOR_LINE_LENGTH == 60

    def test_max_subject_length(self):
        """Test MAX_SUBJECT_LENGTH constant."""
        assert display_config.MAX_SUBJECT_LENGTH == 100

    def test_max_sender_length(self):
        """Test MAX_SENDER_LENGTH constant."""
        assert display_config.MAX_SENDER_LENGTH == 50

    def test_preview_length(self):
        """Test PREVIEW_LENGTH constant."""
        assert display_config.PREVIEW_LENGTH == 200

    def test_date_format(self):
        """Test DATE_FORMAT constant."""
        assert display_config.DATE_FORMAT == "%Y-%m-%d %H:%M:%S"

    def test_separator_line(self):
        """Test SEPARATOR_LINE constant."""
        assert display_config.SEPARATOR_LINE == "=" * 60

