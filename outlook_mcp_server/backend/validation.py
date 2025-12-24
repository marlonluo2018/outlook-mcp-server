"""Unified validation utilities for Outlook MCP Server.

This module consolidates all common validation patterns to eliminate code duplication.
"""

from typing import Optional, List, Union

from .logging_config import get_logger
from .config import (
    outlook_config,
    email_format_config,
    attachment_config,
    email_metadata_config,
    batch_config,
    display_config,
    performance_config,
    validation_config
)

logger = get_logger(__name__)


class ValidationError(Exception):
    """Custom exception for validation errors."""
    pass


class OutlookConstants:
    """Outlook COM constants (deprecated - use config.outlook_config)."""
    OL_MAIL_ITEM = outlook_config.OL_MAIL_ITEM
    OL_CONTACT_ITEM = outlook_config.OL_CONTACT_ITEM
    OL_DISTRIBUTION_LIST_ITEM = outlook_config.OL_DISTRIBUTION_LIST_ITEM
    OL_JOURNAL_ITEM = outlook_config.OL_JOURNAL_ITEM
    OL_NOTE_ITEM = outlook_config.OL_NOTE_ITEM
    OL_POST_ITEM = outlook_config.OL_POST_ITEM
    OL_TASK_ITEM = outlook_config.OL_TASK_ITEM


class BodyFormat:
    """Email body format constants (deprecated - use config.email_format_config)."""
    OL_FORMAT_PLAIN = email_format_config.OL_FORMAT_PLAIN
    OL_FORMAT_HTML = email_format_config.OL_FORMAT_HTML
    OL_FORMAT_RICH_TEXT = email_format_config.OL_FORMAT_RICH_TEXT


class AttachmentType:
    """Attachment type constants (deprecated - use config.attachment_config)."""
    BY_VALUE = attachment_config.BY_VALUE
    BY_REFERENCE = attachment_config.BY_REFERENCE
    EMBEDDED = attachment_config.EMBEDDING
    OLE = attachment_config.OLE


class Importance:
    """Email importance constants (deprecated - use config.email_metadata_config)."""
    LOW = email_metadata_config.IMPORTANCE_LOW
    NORMAL = email_metadata_config.IMPORTANCE_NORMAL
    HIGH = email_metadata_config.IMPORTANCE_HIGH


class Sensitivity:
    """Email sensitivity constants (deprecated - use config.email_metadata_config)."""
    NORMAL = email_metadata_config.SENSITIVITY_NORMAL
    PERSONAL = email_metadata_config.SENSITIVITY_PERSONAL
    PRIVATE = email_metadata_config.SENSITIVITY_PRIVATE
    CONFIDENTIAL = email_metadata_config.SENSITIVITY_CONFIDENTIAL


class FlagStatus:
    """Email flag status constants (deprecated - use config.email_metadata_config)."""
    UNFLAGGED = email_metadata_config.FLAG_STATUS_UNFLAGGED
    FLAGGED = email_metadata_config.FLAG_STATUS_FLAGGED
    COMPLETE = email_metadata_config.FLAG_STATUS_COMPLETE


class BatchLimits:
    """Batch operation limits (deprecated - use config.batch_config)."""
    OUTLOOK_BCC_LIMIT = batch_config.OUTLOOK_BCC_LIMIT
    IMAGE_EMBEDDING_SIZE_THRESHOLD = batch_config.IMAGE_EMBEDDING_SIZE_THRESHOLD


class CacheThresholds:
    """Cache performance thresholds (deprecated - use config.performance_config)."""
    BINARY_SEARCH_THRESHOLD = performance_config.BINARY_SEARCH_THRESHOLD
    MAX_CACHE_SIZE = performance_config.MAX_CACHE_SIZE


class DisplayConstants:
    """Display formatting constants (deprecated - use config.display_config)."""
    SEPARATOR_LINE_LENGTH = display_config.SEPARATOR_LINE_LENGTH


class BatchProcessing:
    """Batch processing constants for email operations (deprecated - use config.batch_config)."""
    DEFAULT_BATCH_SIZE = batch_config.DEFAULT_BATCH_SIZE
    FAST_MODE_BATCH_SIZE = batch_config.FAST_MODE_BATCH_SIZE
    FULL_EXTRACTION_BATCH_SIZE = batch_config.FULL_EXTRACTION_BATCH_SIZE


def validate_search_term(search_term: str) -> str:
    """Validate search term parameter.

    Args:
        search_term: Search term to validate

    Returns:
        Validated search term

    Raises:
        ValidationError: If search term is invalid
    """
    if not search_term or not isinstance(search_term, str):
        raise ValidationError("Search term must be a non-empty string")
    search_term = search_term.strip()
    if not search_term:
        raise ValidationError("Search term must be a non-empty string")
    return search_term


def validate_days_parameter(days: int, min_days: int = 1, max_days: int = 30) -> int:
    """Validate days parameter.

    Args:
        days: Days parameter to validate
        min_days: Minimum allowed days (default: 1)
        max_days: Maximum allowed days (default: 30)

    Returns:
        Validated days parameter

    Raises:
        ValidationError: If days parameter is invalid
    """
    if not isinstance(days, int):
        raise ValidationError("Days parameter must be an integer")
    if days < min_days or days > max_days:
        raise ValidationError(f"Days parameter must be between {min_days} and {max_days}")
    return days


def validate_folder_name(folder_name: Optional[str]) -> Optional[str]:
    """Validate and normalize folder name.

    Args:
        folder_name: Folder name to validate

    Returns:
        Normalized folder name or None

    Raises:
        ValidationError: If folder name is invalid
    """
    if folder_name is None:
        return None
    if not isinstance(folder_name, str):
        raise ValidationError("Folder name must be a string or None")
    folder_name = folder_name.strip()
    if folder_name.lower() in ["null", ""]:
        return None
    return folder_name if folder_name else None


def validate_email_address(email: str) -> str:
    """Validate email address format with comprehensive checks.

    Args:
        email: Email address to validate

    Returns:
        Validated email address (trimmed and lowercased domain)

    Raises:
        ValidationError: If email address is invalid
    """
    import re

    if not email or not isinstance(email, str):
        raise ValidationError("Email address must be a non-empty string")

    email = email.strip()

    if len(email) > validation_config.MAX_EMAIL_LENGTH:
        raise ValidationError(f"Email address is too long (maximum {validation_config.MAX_EMAIL_LENGTH} characters)")

    if len(email.split("@")[0]) > validation_config.MAX_EMAIL_LOCAL_PART_LENGTH:
        raise ValidationError(f"Email local part is too long (maximum {validation_config.MAX_EMAIL_LOCAL_PART_LENGTH} characters)")

    pattern = r"^[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}$"
    if not re.match(pattern, email):
        raise ValidationError(f"Invalid email address format: {email}")

    return email


def validate_email_addresses(emails: Union[str, List[str]]) -> List[str]:
    """Validate one or more email addresses.

    Args:
        emails: Single email string or list of email strings

    Returns:
        List of validated email addresses

    Raises:
        ValidationError: If any email address is invalid
    """
    if emails is None:
        raise ValidationError("Email addresses cannot be None")

    if isinstance(emails, str):
        emails = [emails]

    if not isinstance(emails, list):
        raise ValidationError("Email addresses must be a string or list of strings")

    validated_emails = []
    for email in emails:
        if not isinstance(email, str) or not email.strip():
            raise ValidationError("All email addresses must be non-empty strings")
        validated_emails.append(validate_email_address(email))

    return validated_emails


def validate_email_number(email_number: int, cache_size: int) -> int:
    """Validate email number against cache size.

    Args:
        email_number: Email number to validate
        cache_size: Current cache size

    Returns:
        Validated email number

    Raises:
        ValidationError: If email number is out of range
    """
    if not isinstance(email_number, int):
        raise ValidationError("Email number must be an integer")
    if email_number < 1 or email_number > cache_size:
        raise ValidationError(
            f"Email number {email_number} is out of range. Available range: 1-{cache_size}"
        )
    return email_number


def validate_page_parameter(page: int, total_pages: int) -> int:
    """Validate page parameter.

    Args:
        page: Page number to validate
        total_pages: Total number of pages available

    Returns:
        Validated page number

    Raises:
        ValidationError: If page number is invalid
    """
    if not isinstance(page, int):
        raise ValidationError("Page parameter must be an integer")
    if page < 1:
        raise ValidationError("Page parameter must be at least 1")
    if total_pages > 0 and page > total_pages:
        raise ValidationError(
            f"Page {page} is out of range. Total pages: {total_pages}"
        )
    return page


def validate_not_empty(value: str, field_name: str = "Field") -> str:
    """Validate that a string value is not empty or whitespace.

    Args:
        value: String value to validate
        field_name: Name of the field for error messages

    Returns:
        Validated string value

    Raises:
        ValidationError: If value is empty or whitespace
    """
    if value is None or not isinstance(value, str):
        raise ValidationError(f"{field_name} must be a non-empty string")
    value = value.strip()
    if not value:
        raise ValidationError(f"{field_name} must not be empty")
    return value


def sanitize_search_term(search_term: str) -> str:
    """Sanitize search term to prevent DASL injection.

    Args:
        search_term: Raw search term from user

    Returns:
        Sanitized search term
    """
    if not search_term:
        return ""
    return "".join(c for c in search_term if c.isalnum() or c in " .-_@").strip()


def normalize_email_address(email: str) -> str:
    """Normalize email address for comparison.

    Handles case sensitivity, display name formats, and extra whitespace.

    Args:
        email: Email address to normalize

    Returns:
        Normalized email address for comparison
    """
    if not email:
        return ""

    normalized = email.strip().rstrip(";").strip()

    if "<" in normalized and ">" in normalized:
        start = normalized.find("<")
        end = normalized.find(">")
        if start < end:
            normalized = normalized[start + 1 : end]

    return normalized.lower()


def validate_recipients_list(
    recipients: Optional[Union[str, List[str]]]
) -> Optional[List[str]]:
    """Validate and normalize recipients list.

    Args:
        recipients: Single email string, list of emails, or None

    Returns:
        List of validated emails or None

    Raises:
        ValidationError: If recipients are invalid
    """
    if recipients is None:
        return None

    if isinstance(recipients, str):
        if not recipients.strip():
            return None
        recipients = [recipients]

    if not isinstance(recipients, list):
        raise ValidationError("Recipients must be a string or list of strings")

    validated_emails = []
    for email in recipients:
        if isinstance(email, str) and email.strip():
            validated_emails.append(email.strip())

    return validated_emails if validated_emails else None


def get_folder_path_safe(folder_name: Optional[str] = None) -> str:
    """Get safe folder path, defaulting to Inbox if not provided.

    Args:
        folder_name: Folder name to use

    Returns:
        Safe folder path
    """
    return folder_name if folder_name else "Inbox"


def validate_cache_available(cache_size: int) -> None:
    """Validate that email cache is available and not empty.

    Args:
        cache_size: Current cache size

    Raises:
        ValidationError: If cache is empty
    """
    if cache_size == 0:
        raise ValidationError("No emails available - please list emails first.")


def execute_cache_loading_operation(
    operation_func,
    operation_name: str,
    validation_func=None,
    validation_params=None,
    message_suffix: str = None,
    **operation_kwargs
) -> dict:
    """Unified cache loading operation wrapper for tool functions.

    This function provides a consistent interface for all cache loading operations
    including list, search, and load tools. It handles validation, error handling,
    and message formatting uniformly.

    Args:
        operation_func: The backend function to execute (e.g., list_recent_emails, search_email_by_subject)
        operation_name: Name of the operation for logging and error messages
        validation_func: Optional validation function to call before executing the operation
        validation_params: Optional dictionary of parameters to pass to validation function
        message_suffix: Optional suffix to append to success message (e.g., " (max 30 days)")
        **operation_kwargs: Keyword arguments to pass to the operation function

    Returns:
        dict: Response containing success or error message:
        {
            "type": "text",
            "text": "Found X emails in 'Inbox' from last 7 days (max 30 days). Use 'view_email_cache_tool' to view them."
        }

    Example:
        result = execute_cache_loading_operation(
            operation_func=list_recent_emails,
            operation_name="list_recent_emails",
            validation_func=lambda: validate_days_parameter(days),
            validation_params=None,
            message_suffix=" (max 30 days)",
            folder_name=folder_path,
            days=days
        )
    """
    try:
        if validation_func and validation_params:
            validation_func(**validation_params)
        elif validation_func:
            validation_func()
    except ValidationError as e:
        return {"type": "text", "text": f"Validation error: {str(e)}"}
    
    logger.info(f"{operation_name} called with {operation_kwargs}")
    
    try:
        emails, message = operation_func(**operation_kwargs)
        logger.info(f"{operation_name} returned: {len(emails)} emails, message: {message}")
        
        if message_suffix:
            message = message + message_suffix
        
        if "view_email_cache_tool" not in message:
            message = message + ". Use 'view_email_cache_tool' to view them."
        
        return {"type": "text", "text": message}
    except Exception as e:
        logger.error(f"Error in {operation_name}: {e}")
        import traceback
        traceback.print_exc()
        return {"type": "text", "text": f"Error retrieving emails: {str(e)}"}
