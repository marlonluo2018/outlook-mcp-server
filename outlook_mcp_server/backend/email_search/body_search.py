"""
Email search functionality for body content searches.

This module provides functions to search emails by body content.
"""

from typing import Any, Dict, List, Optional, Tuple

from .unified_search import unified_search


def search_email_by_body(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Tuple[List[Dict[str, Any]], str]:
    """
    Search emails by body content and return list of emails with note.
    
    Args:
        search_term: The term to search for in email body content
        days: Number of days to look back (1-30, default: 7)
        folder_name: Folder to search in (defaults to Inbox)
        match_all: If True, requires ALL search terms to match (AND logic).
                  If False, matches ANY search term (OR logic)
    
    Returns:
        Tuple of (list of email dictionaries, status message)
    """
    return unified_search(search_term, days, folder_name, match_all, "body")