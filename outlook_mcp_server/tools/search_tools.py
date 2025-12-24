"""Email search tools for Outlook MCP Server."""

# Type imports
from typing import Any, Callable, Dict, Optional

# Local application imports
from ..backend import email_search
from ..backend.logging_config import get_logger
from ..backend.validation import (
    ValidationError,
    execute_cache_loading_operation,
    get_folder_path_safe,
    validate_days_parameter,
    validate_folder_name,
    validate_search_term
)

logger = get_logger(__name__)

# Import specific functions from the email_search module
list_recent_emails = email_search.list_recent_emails
search_email_by_subject = email_search.search_email_by_subject
search_email_by_sender = email_search.search_email_by_sender
search_email_by_recipient = email_search.search_email_by_recipient
search_email_by_body = email_search.search_email_by_body


def list_recent_emails_tool(days: int = 7, folder_name: Optional[str] = None) -> Dict[str, Any]:
    """Load emails into cache and return count message.

    Args:
        days: Days to look back (1-30, default:7, max:30)
        folder_name: Folder to search (default:Inbox, or use full path like "user@company.com/Inbox")

    Returns:
        dict: Response containing email count message:
        {
            "type": "text",
            "text": "Found X emails from last N days (max 30 days). Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"
    """
    def validate_params() -> Optional[str]:
        validate_days_parameter(days)
        return validate_folder_name(folder_name)
    
    validated_folder = validate_params()
    folder_path = get_folder_path_safe(validated_folder)
    
    return execute_cache_loading_operation(
        operation_func=list_recent_emails,
        operation_name="list_recent_emails_tool",
        message_suffix=" (max 30 days)",
        folder_name=folder_path,
        days=days
    )


def search_email_by_subject_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Dict[str, Any]:
    """Search email subjects and load matching emails into cache.

    This function only searches the email subject field. It does not search in the email body,
    sender name, recipients, or other fields.

    Args:
        search_term: Plain text search term (colons are allowed as part of regular text)
        days: Number of days to look back (1-30, default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last N days (max 30 days). Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"

    """
    def validate_params() -> Optional[str]:
        validate_search_term(search_term)
        validate_days_parameter(days)
        return validate_folder_name(folder_name)
    
    validated_folder = validate_params()
    
    return execute_cache_loading_operation(
        operation_func=search_email_by_subject,
        operation_name="search_email_by_subject_tool",
        message_suffix=f" from last {days} days (max 30 days)",
        search_term=search_term,
        days=days,
        folder_name=validated_folder,
        match_all=match_all
    )


def search_email_by_sender_name_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Dict[str, Any]:
    """Search emails by sender name and load matching emails into cache.

    This function only searches the sender name field. It does not search in the email body,
    subject, recipients, or other fields.

    Search by name only, not email address.

    Args:
        search_term: Plain text search term for sender name (colons are allowed as part of regular text)
        days: Number of days to look back (1-30, default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last N days (max 30 days). Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"

    """
    def validate_params() -> Optional[str]:
        validate_search_term(search_term)
        validate_days_parameter(days)
        return validate_folder_name(folder_name)
    
    validated_folder = validate_params()
    
    return execute_cache_loading_operation(
        operation_func=search_email_by_sender,
        operation_name="search_email_by_sender_name_tool",
        message_suffix=f" from last {days} days (max 30 days)",
        search_term=search_term,
        days=days,
        folder_name=validated_folder,
        match_all=match_all
    )


def search_email_by_recipient_name_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Dict[str, Any]:
    """Search emails by recipient name and load matching emails into cache.

    This function only searches the recipient (To) field. It does not search in the email body,
    subject, sender, or other fields.

    Search by name only, not email address.

    Args:
        search_term: Plain text search term for recipient name (colons are allowed as part of regular text)
        days: Number of days to look back (1-30, default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last N days (max 30 days). Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"

    """
    def validate_params() -> Optional[str]:
        validate_search_term(search_term)
        validate_days_parameter(days)
        return validate_folder_name(folder_name)
    
    validated_folder = validate_params()
    
    return execute_cache_loading_operation(
        operation_func=search_email_by_recipient,
        operation_name="search_email_by_recipient_name_tool",
        message_suffix=f" from last {days} days (max 30 days)",
        search_term=search_term,
        days=days,
        folder_name=validated_folder,
        match_all=match_all
    )


def search_email_by_body_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Dict[str, Any]:
    """Search emails by body content and load matching emails into cache.

    This function searches the email body content. It does not search in the subject,
    sender name, recipients, or other fields.

    Note: Searching email body is slower than searching other fields as it requires
    loading the full content of each email.

    Args:
        search_term: Plain text search term (colons are allowed as part of regular text)
                    For exact phrase matching, enclose the term in quotes (e.g., "red hat partner day")
                    For word-based matching, use the term without quotes (e.g., red hat partner day)
        days: Number of days to look back (1-30, default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"
        {
            "type": "text",
            "text": "Found X matching emails from last N days (max 30 days). Use 'view_email_cache_tool' to view them."
        }

    """
    def validate_params() -> Optional[str]:
        validate_search_term(search_term)
        validate_days_parameter(days)
        return validate_folder_name(folder_name)
    
    validated_folder = validate_params()
    
    return execute_cache_loading_operation(
        operation_func=search_email_by_body,
        operation_name="search_email_by_body_tool",
        message_suffix=f" from last {days} days (max 30 days)",
        search_term=search_term,
        days=days,
        folder_name=validated_folder,
        match_all=match_all
    )