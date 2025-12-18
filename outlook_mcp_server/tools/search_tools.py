"""Email search tools for Outlook MCP Server."""

from typing import Optional
from ..backend import email_search

# Import specific functions from the email_search module
list_recent_emails = email_search.list_recent_emails
search_email_by_subject = email_search.search_email_by_subject
search_email_by_sender = email_search.search_email_by_sender
search_email_by_recipient = email_search.search_email_by_recipient
search_email_by_body = email_search.search_email_by_body


def list_recent_emails_tool(days: int = 7, folder_name: Optional[str] = None) -> dict:
    """Load emails into cache and return count message.

    Args:
        days: Days to look back (1-30, default:7)
        folder_name: Folder to search (default:Inbox, or use full path like "user@company.com/Inbox")

    Returns:
        dict: Response containing email count message:
        {
            "type": "text",
            "text": "Found X emails from last Y days. Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"
    """
    # Input validation
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    
    # Ensure folder_name is a string or None
    if folder_name is not None and not isinstance(folder_name, str):
        raise ValueError("Folder name must be a string or None")
    
    # Log the request for debugging
    import logging
    logger = logging.getLogger(__name__)
    logger.info(f"list_recent_emails_tool called with days={days}, folder_name={folder_name}")
    
    # If folder_name is None, use the default
    if folder_name is None:
        folder_name = "Inbox"
    
    try:
        logger.info(f"Calling list_recent_emails with folder={folder_name}, days={days}")
        emails, message = list_recent_emails(folder_name=folder_name, days=days)
        logger.info(f"list_recent_emails returned: {len(emails)} emails, message: {message}")
        
        # Add debugging info
        from ..backend.shared import email_cache
        debug_info = f"\n\nDEBUG: Cache contains {len(email_cache)} emails"
        
        return {"type": "text", "text": message + debug_info}
    except Exception as e:
        logger.error(f"Error in list_recent_emails_tool: {e}")
        import traceback
        traceback.print_exc()
        return {"type": "text", "text": f"Error retrieving emails: {str(e)}"}


def search_email_by_subject_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> dict:
    """Search email subjects and load matching emails into cache.

    This function only searches the email subject field. It does not search in the email body,
    sender name, recipients, or other fields.

    Args:
        search_term: Plain text search term (colons are allowed as part of regular text)
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last Y days. Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"

    """
    if not search_term or not isinstance(search_term, str):
        raise ValueError("Search term must be a non-empty string")
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    emails, note = search_email_by_subject(search_term, days, folder_name, match_all=match_all)
    result = f"Found {len(emails)} matching emails{note}. Use 'view_email_cache_tool' to view them."
    return {"type": "text", "text": result}


def search_email_by_sender_name_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> dict:
    """Search emails by sender name and load matching emails into cache.

    This function only searches the sender name field. It does not search in the email body,
    subject, recipients, or other fields.

    Search by name only, not email address.

    Args:
        search_term: Plain text search term for sender name (colons are allowed as part of regular text)
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last Y days. Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"

    """
    if not search_term or not isinstance(search_term, str):
        raise ValueError("Search term must be a non-empty string")
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    emails, note = search_email_by_sender(search_term, days, folder_name, match_all=match_all)
    result = f"Found {len(emails)} matching emails{note}. Use 'view_email_cache_tool' to view them."
    return {"type": "text", "text": result}


def search_email_by_recipient_name_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> dict:
    """Search emails by recipient name and load matching emails into cache.

    This function only searches the recipient (To) field. It does not search in the email body,
    subject, sender, or other fields.

    Search by name only, not email address.

    Args:
        search_term: Plain text search term for recipient name (colons are allowed as part of regular text)
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last Y days. Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"

    """
    if not search_term or not isinstance(search_term, str):
        raise ValueError("Search term must be a non-empty string")
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    emails, note = search_email_by_recipient(search_term, days, folder_name, match_all=match_all)
    result = f"Found {len(emails)} matching emails{note}. Use 'view_email_cache_tool' to view them."
    return {"type": "text", "text": result}


def search_email_by_body_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> dict:
    """Search emails by body content and load matching emails into cache.

    This function searches the email body content. It does not search in the subject,
    sender name, recipients, or other fields.

    Note: Searching email body is slower than searching other fields as it requires
    loading the full content of each email.

    Args:
        search_term: Plain text search term (colons are allowed as part of regular text)
                    For exact phrase matching, enclose the term in quotes (e.g., "red hat partner day")
                    For word-based matching, use the term without quotes (e.g., red hat partner day)
        days: Number of days to look back (default: 7, max: 30)
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
            "text": "Found X matching emails from last Y days. Use 'view_email_cache_tool' to view them."
        }

    """
    if not search_term or not isinstance(search_term, str):
        raise ValueError("Search term must be a non-empty string")
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    emails, note = search_email_by_body(search_term, days, folder_name, match_all=match_all)
    result = f"Found {len(emails)} matching emails{note}. Use 'view_email_cache_tool' to view them."
    return {"type": "text", "text": result}