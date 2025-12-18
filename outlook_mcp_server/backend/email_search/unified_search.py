"""
Unified search functionality for email operations.

This module provides the main unified search function that coordinates
between server-side search implementations.
"""

import logging
from typing import Any, Dict, List, Optional, Tuple

from ..outlook_session.session_manager import OutlookSessionManager
from ..shared import email_cache, add_email_to_cache, clear_email_cache
from ..validators import EmailSearchParams
from .search_common import get_folder_path_safe, is_server_search_supported, extract_email_info
from .server_search import server_side_search

# Set up logging
logger = logging.getLogger(__name__)


def unified_search(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True, search_type: str = "subject"
) -> Tuple[List[Dict[str, Any]], str]:
    """
    Unified search function that prioritizes fast server-side search.
    
    Args:
        search_term: The term to search for
        days: Number of days to look back
        folder_name: Folder to search in (defaults to Inbox)
        match_all: Whether to match all terms (AND logic) or any term (OR logic)
        search_type: Type of search (subject, sender, recipient, body)
    
    Returns:
        Tuple of (list of email dictionaries, status message)
    """
    if not search_term or not isinstance(search_term, str):
        return [], "Search term must be a non-empty string"
    
    if days < 1 or days > 30:
        return [], "Days must be between 1 and 30"
    
    try:
        folder_path = get_folder_path_safe(folder_name)
        
        with OutlookSessionManager() as session:
            folder = session.get_folder(folder_path)
            if not folder:
                return [], f"Folder '{folder_path}' not found"
            
            # Use server-side search only - completely disable client-side search for performance
            results = []
            if is_server_search_supported(search_type):
                try:
                    results = server_side_search(folder, search_term, days, search_type, match_all, session.outlook_namespace)
                    if results:
                        logger.info(f"Server-side search successful: found {len(results)} results")
                    else:
                        logger.info(f"Server-side search completed but no results found")
                except Exception as e:
                    logger.error(f"Server-side search failed: {e}")
                    # Return empty results instead of falling back to slow client-side search
                    return [], f"Search failed for '{search_term}' in '{folder_path}'"
            else:
                # For unsupported search types, return empty rather than using slow client-side
                logger.warning(f"Search type '{search_type}' not supported by server-side search")
                return [], f"Search type '{search_type}' is not supported for performance reasons"
            
            if not results:
                return [], f"No emails found in '{folder_path}' matching '{search_term}'"
            
            # Clear cache before adding new search results
            clear_email_cache()
            
            # Convert results to cache format
            email_list = []
            for item in results:
                try:
                    email_data = extract_email_info(item)
                    add_email_to_cache(email_data["entry_id"], email_data)
                    email_list.append(email_data)
                except Exception as e:
                    logger.warning(f"Failed to cache email: {e}")
                    continue
            
            if not email_list:
                return [], "No valid emails found"
            
            # Sort by received time (newest first)
            email_list.sort(key=lambda x: x.get("received_time", ""), reverse=True)
            
            # Ensure cache order matches the returned email list
            clear_email_cache()
            for email_data in email_list:
                add_email_to_cache(email_data["entry_id"], email_data)
            
            message = f"Found {len(email_list)} emails in '{folder_path}'"
            return email_list, message
            
    except Exception as e:
        error_msg = f"Error searching emails: {e}"
        logger.error(error_msg)
        return [], error_msg