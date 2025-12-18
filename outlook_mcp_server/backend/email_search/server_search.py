"""
Server-side search functionality for email operations.

This module provides functions for performing server-side searches using
Outlook's AdvancedSearch functionality, which is more efficient for large folders.
"""

import logging
import time
from datetime import datetime, timedelta, timezone
from typing import Any, List

from ..outlook_session.session_manager import OutlookSessionManager
from .search_common import get_date_limit

# Set up logging
logger = logging.getLogger(__name__)


def server_side_search(
    folder, search_term: str, days: int, search_type: str, match_all: bool, namespace=None
) -> List[Any]:
    """
    Perform server-side search using Outlook's AdvancedSearch functionality.
    
    This is more efficient for large folders as it leverages Outlook's indexing.
    """
    try:
        # Use the provided namespace from the existing session
        # This avoids creating duplicate COM objects which can cause issues
        
        # Build the search criteria with proper formatting
        date_limit = get_date_limit(days)
        
        # Escape single quotes in search term to prevent syntax errors
        escaped_search_term = search_term.replace("'", "''")
        
        # Use SQL format for all conditions to ensure compatibility
        sql_conditions = []
        
        # Add date condition
        sql_conditions.append(f"urn:schemas:httpmail:datereceived >= '{date_limit.strftime('%Y-%m-%d')}'")
        
        # Add content condition based on search type
        if search_type == "subject":
            if match_all:
                sql_conditions.append(f"urn:schemas:httpmail:subject LIKE '%{escaped_search_term}%'")
            else:
                sql_conditions.append(f"urn:schemas:httpmail:subject LIKE '%{escaped_search_term}%'")
        elif search_type == "sender":
            sql_conditions.append(f"urn:schemas:httpmail:fromname LIKE '%{escaped_search_term}%'")
        elif search_type == "recipient":
            sql_conditions.append(f"urn:schemas:httpmail:to LIKE '%{escaped_search_term}%'")
        
        # Combine all conditions with AND
        search_criteria = "@SQL=" + " AND ".join(sql_conditions)
        
        logger.info(f"Server-side search criteria: {search_criteria}")
        
        # Get the folder path correctly
        folder_path = folder.FolderPath if hasattr(folder, 'FolderPath') else str(folder)
        
        logger.info(f"Folder path: {folder_path}")
        logger.info(f"Search criteria: {search_criteria}")
        
        # Try using the folder's Items collection with Restrict method instead of AdvancedSearch
        try:
            # Use Restrict method on the folder's Items collection
            items = folder.Items
            restricted_items = items.Restrict(search_criteria)
            results = list(restricted_items)
            logger.info(f"Restrict method completed: found {len(results)} results")
            return results
        except Exception as e:
            logger.warning(f"Restrict method failed: {e}")
            
            # Fallback to AdvancedSearch with proper scope format
            try:
                # Use Application.AdvancedSearch with proper scope format
                outlook = namespace.Application if hasattr(namespace, 'Application') else namespace
                
                # Create scope in the format "Inbox" or "\\Personal Folders\Inbox"
                scope = folder_path
                logger.info(f"Using scope: {scope}")
                
                search_results = outlook.AdvancedSearch(
                    Scope=scope, 
                    Filter=search_criteria, 
                    SearchSubFolders=True
                )
                
                # Wait for search to complete with timeout
                max_wait_time = 5  # seconds
                start_time = time.time()
                
                while search_results.SearchState != 1:  # 1 = SearchComplete
                    time.sleep(0.1)
                    if time.time() - start_time > max_wait_time:
                        logger.warning("Server-side search timed out")
                        return []
                
                results = list(search_results.Results)
                logger.info(f"AdvancedSearch completed: found {len(results)} results")
                return results
                
            except Exception as e2:
                logger.error(f"AdvancedSearch also failed: {e2}")
                return []
        
        # Wait for search to complete with timeout
        max_wait_time = 5  # seconds
        start_time = time.time()
        
        while search_results.SearchState != 1:  # 1 = SearchComplete
            time.sleep(0.1)
            if time.time() - start_time > max_wait_time:
                logger.warning("Server-side search timed out")
                return []
        
        results = list(search_results.Results)
        logger.info(f"Server-side search completed: found {len(results)} results")
        return results
        
    except Exception as e:
        logger.error(f"Server-side search failed: {e}")
        logger.error(f"Error type: {type(e)}")
        logger.error(f"Error details: {str(e)}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return []