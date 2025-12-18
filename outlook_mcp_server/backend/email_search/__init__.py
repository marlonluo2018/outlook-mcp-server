#!/usr/bin/env python3
"""Email search functionality for Outlook MCP server.

This package provides various email search capabilities including:
|- Subject-based searches
|- Sender-based searches  
|- Recipient-based searches
|- Body content searches
|- Shared search utilities
"""

# Import from the modular search components
from .subject_search import search_email_by_subject
from .sender_search import search_email_by_sender
from .recipient_search import search_email_by_recipient
from .body_search import search_email_by_body

# Import from email listing module
from .email_listing import list_recent_emails, get_emails_from_folder

# Import from outlook_session modules (moved from search_utils)
from ..outlook_session.folder_operations import get_folder_emails, list_folders
from ..outlook_session.email_operations import (
    get_email_by_number,
    move_email_to_folder,
    delete_email_by_number,
)

# Direct imports for search functions (previously wrapped in search_utils)
from .sender_search import search_email_by_sender as search_email_by_from
from .recipient_search import search_email_by_recipient as search_email_by_to

# Import shared utilities
from .search_common import get_folder_path_safe, get_date_limit, is_server_search_supported, extract_email_info

# Import unified search
from .unified_search import unified_search

# Import server search
from .server_search import server_side_search

__all__ = [
    # Search functions from modular components
    "search_email_by_subject",
    "search_email_by_sender", 
    "search_email_by_recipient",
    "search_email_by_body",
    
    # Email listing functions
    "list_folders",
    "list_recent_emails",
    "get_emails_from_folder",
    "get_folder_emails",
    "get_email_by_number",
    "move_email_to_folder",
    "delete_email_by_number",
    
    # Search functions from modular components
    "search_email_by_from",
    "search_email_by_to",
    
    # Shared utilities
    "get_folder_path_safe",
    "get_date_limit",
    "is_server_search_supported",
    "extract_email_info",
    
    # Search implementations
    "unified_search",
    "server_side_search",
]