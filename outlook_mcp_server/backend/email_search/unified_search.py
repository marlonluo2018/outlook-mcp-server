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
from .search_common import get_folder_path_safe, is_server_search_supported, extract_email_info, unified_cache_load_workflow
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
            
            # Convert results to email data format - OPTIMIZED
            email_list = []
            
            # Clear COM cache before processing to prevent memory growth
            from .search_common import clear_com_attribute_cache
            clear_com_attribute_cache()
            
            # Process in batches for better performance
            batch_size = 50
            total_results = len(results)
            
            for batch_start in range(0, total_results, batch_size):
                batch_end = min(batch_start + batch_size, total_results)
                batch_items = results[batch_start:batch_end]
                
                # Process batch
                for item in batch_items:
                    try:
                        email_data = extract_email_info(item)
                        if email_data:
                            email_list.append(email_data)
                    except Exception as e:
                        logger.warning(f"Failed to extract email info: {e}")
                        continue
                
                # Clear COM cache periodically to prevent memory growth
                if batch_start % 200 == 0:
                    clear_com_attribute_cache()
            
            if not email_list:
                return [], "No valid emails found"
            
            # Sort by received time (newest first)
            email_list.sort(key=lambda x: x.get("received_time", ""), reverse=True)
            
            # Use unified cache loading workflow for consistent cache management
            success = unified_cache_load_workflow(email_list, f"unified_search({search_term})")
            if success:
                logger.info(f"Unified cache workflow completed successfully for {len(email_list)} search results")
            else:
                logger.warning("Unified cache workflow failed for search results")
            
            message = f"Found {len(email_list)} emails in '{folder_path}'"
            return email_list, message
            
    except Exception as e:
        error_msg = f"Error searching emails: {e}"
        logger.error(error_msg)
        return [], error_msg