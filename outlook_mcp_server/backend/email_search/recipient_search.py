"""
Email search functionality for recipient-based searches.

This module provides functions to search emails by recipient name.
"""

from typing import Any, Dict, List, Optional, Tuple

from .unified_search import unified_search


def search_email_by_recipient(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Tuple[List[Dict[str, Any]], str]:
    """
    Search emails by recipient name and return list of emails with note.
    
    Args:
        search_term: The recipient name to search for
        days: Number of days to look back (1-30, default: 7)
        folder_name: Folder to search in (defaults to Inbox)
        match_all: If True, requires ALL search terms to match (AND logic).
                  If False, matches ANY search term (OR logic)
    
    Returns:
        Tuple of (list of email dictionaries, status message)
    """
    return unified_search(search_term, days, folder_name, match_all, "recipient")