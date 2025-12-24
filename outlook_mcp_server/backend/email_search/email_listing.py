"""
Email listing functionality for retrieving emails from folders.

This module provides functions for listing recent emails and getting emails
from specific folders with various filtering options.
"""

# Standard library imports
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Tuple

# Local application imports
from ..logging_config import get_logger
from ..outlook_session.session_manager import OutlookSessionManager
from ..shared import add_email_to_cache, clear_email_cache, email_cache, email_cache_order
from ..validators import EmailListParams
from .search_common import (
    extract_email_info,
    get_folder_path_safe,
    unified_cache_load_workflow
)

logger = get_logger(__name__)


def list_recent_emails(folder_name: str = "Inbox", days: int = None) -> Tuple[List[Dict[str, Any]], str]:
    """Public interface for listing emails (used by CLI).
    Loads emails into cache and returns (emails, message) tuple.
    
    Uses improved cache workflow:
    1. Clear both memory and disk cache
    2. Load fresh data from Outlook
    3. Save immediately to disk
    """
    try:
        # Default to 30 days if not specified to ensure we get results
        effective_days = days or 30
        params = EmailListParams(days=effective_days, folder_name=folder_name)
        
        # Minimal logging for performance
        if effective_days <= 7:
            logger.debug(f"list_recent_emails: {folder_name}, {days} days")
    except Exception as e:
        logger.error(f"Validation error in list_recent_emails: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    # Load fresh emails from Outlook
    emails, note = get_emails_from_folder_optimized(folder_name=params.folder_name, days=params.days)
    
    # Use unified cache loading workflow for consistent cache management
    # This handles all 3 steps: clear cache, load data, save to disk
    if emails and "Error:" not in note:
        unified_cache_load_workflow(emails, f"list_recent_emails({params.folder_name})")

    days_str = f" from last {params.days} days" if params.days else ""
    
    return emails, f"Found {len(emails)} emails in '{params.folder_name}'{days_str}"


def get_emails_from_folder_optimized(folder_name: str = "Inbox", days: int = 7) -> Tuple[List[Dict[str, Any]], str]:
    """
    Optimized version of get_emails_from_folder with performance improvements.
    
    Key optimizations:
    1. Early termination when emails are older than date limit
    2. Optimized batch processing with better batch size
    3. Reduced COM object attribute access
    4. Streamlined email extraction
    """
    try:
        params = EmailListParams(folder_name=folder_name, days=days)
    except Exception as e:
        logger.error(f"Validation error in get_emails_from_folder: {e}")
        return [], f"Error: Invalid parameters: {e}"

    try:
        with OutlookSessionManager() as session:
            folder = session.get_folder(params.folder_name)
            if not folder:
                return [], f"Error: Folder '{params.folder_name}' not found"
            
            # Minimal logging for performance - only log significant events
            if params.days > 7:
                logger.info(f"Processing {params.folder_name} for {params.days} days")
            
            # Get total item count first (needed for batch size calculation)
            total_items = folder.Items.Count if hasattr(folder.Items, 'Count') else 0
            
            # OPTIMIZATION 1: Dynamic batch size based on folder size for optimal performance
            def get_optimal_batch_size(folder_size, days_requested):
                """Determine optimal batch size based on folder size and search parameters."""
                if folder_size < 100:
                    return 50  # Small folders: larger batches
                elif folder_size < 500:
                    return 75  # Medium folders: balanced approach
                elif days_requested <= 1:
                    return 100  # Recent searches: larger batches
                elif days_requested <= 7:
                    return 50  # Week searches: medium batches
                else:
                    return 25  # Long searches: smaller batches for better responsiveness
            
            batch_size = get_optimal_batch_size(total_items, params.days)
            
            # OPTIMIZATION 2: Adjust max_items based on days requested
            if params.days and params.days <= 1:
                max_items = 200  # For 1-day searches, 200 items is sufficient
            elif params.days and params.days <= 3:
                max_items = 500  # For 3-day searches, increase to 500
            elif params.days and params.days <= 7:
                max_items = 1000  # For 7-day searches, increase to 1000
            else:
                max_items = 2000  # For longer searches, use higher limit
            
            filtered_items = []
            
            # Calculate date limit if needed
            date_limit = None
            if params.days:
                date_limit = datetime.now(timezone.utc) - timedelta(days=params.days)
            
            # MAJOR OPTIMIZATION: Use Restrict method to filter by date first, then process
            items_collection = folder.Items
            
            # OPTIMIZATION: Sort items by received time (newest first) at the Outlook level
            try:
                items_collection.Sort("[ReceivedTime]", True)  # True = descending order (newest first)
            except Exception as e:
                if params.days > 7:  # Only log for longer operations
                    logger.warning(f"Failed to sort items at Outlook level: {e}")
            
            if date_limit:
                # Use Restrict to filter items by date - this is MUCH faster than individual item access
                date_filter = f"@SQL=urn:schemas:httpmail:datereceived >= '{date_limit.strftime('%Y-%m-%d')}'"
                try:
                    filtered_items = items_collection.Restrict(date_filter)
                    # Convert to list to get count and enable indexing
                    filtered_items_list = list(filtered_items)
                    
                    # Since items are already sorted newest first, just take the first N items
                    items_to_process = min(len(filtered_items_list), max_items)
                    filtered_items = filtered_items_list[:items_to_process]  # Get first N items (newest)
                    
                except Exception as e:
                    if params.days > 7:  # Only log for longer operations
                        logger.warning(f"Restrict method failed: {e}, falling back to manual filtering")
                    # Fallback to manual filtering if Restrict fails
                    filtered_items = []
                    items_to_process = min(total_items, max_items)
                    
                    # Since items are sorted newest first, process from the beginning
                    for i in range(items_to_process):
                        try:
                            item_index = i + 1  # Outlook uses 1-based indexing
                            if item_index > total_items:
                                continue
                                
                            item = items_collection.Item(item_index)
                            if not item:
                                continue
                            
                            # Manual date check
                            if date_limit and hasattr(item, 'ReceivedTime') and item.ReceivedTime:
                                try:
                                    item_time = item.ReceivedTime
                                    if item_time.tzinfo is None:
                                        item_time = item_time.replace(tzinfo=timezone.utc)
                                    
                                    if item_time < date_limit:
                                        continue
                                except Exception:
                                    continue
                            
                            # Basic validation
                            if not hasattr(item, 'Class') or item.Class != 43:
                                continue
                            
                            if not item.ReceivedTime:
                                continue
                            
                            filtered_items.append(item)
                            
                        except Exception as e:
                            logger.debug(f"Error processing item {item_index}: {e}")
                            continue
            else:
                # No date filter - process recent items (already sorted newest first)
                items_to_process = min(total_items, max_items)
                filtered_items = []
                
                for i in range(items_to_process):
                    try:
                        item_index = i + 1  # Outlook uses 1-based indexing
                        if item_index > total_items:
                            continue
                            
                        item = items_collection.Item(item_index)
                        if not item:
                            continue
                        
                        # Basic validation
                        if not hasattr(item, 'Class') or item.Class != 43:
                            continue
                        
                        if not item.ReceivedTime:
                            continue
                        
                        filtered_items.append(item)
                        
                    except Exception as e:
                        logger.debug(f"Error processing item {item_index}: {e}")
                        continue
            
            # Minimal logging for performance
            if len(filtered_items) == 0:
                return [], f"No emails found in '{params.folder_name}' from last {params.days} days"
            
            # OPTIMIZATION 10: Skip sorting if already in correct order (newest first)
            # Since we process in reverse order, items should already be newest first
            
            # OPTIMIZATION 11: Enhanced batch processing with bulk timestamp handling - OPTIMIZED
            email_list = []
            cache_count = 0
            
            # Clear COM cache before processing to prevent memory growth
            from .search_common import clear_com_attribute_cache
            clear_com_attribute_cache()
            
            # MAJOR OPTIMIZATION: Use parallel extraction for list operations
            from .parallel_extractor import extract_emails_optimized
            
            email_list = extract_emails_optimized(filtered_items, use_parallel=True, max_workers=4)
            
            # Cache all extracted emails
            for email_data in email_list:
                if email_data and email_data.get("entry_id"):
                    add_email_to_cache(email_data["entry_id"], email_data)
                    cache_count += 1
            
            if not email_list:
                return [], f"No valid emails found in '{params.folder_name}' from last {params.days} days"
            
            return email_list, f"Found {len(email_list)} emails in '{params.folder_name}' from last {params.days} days"
            
    except Exception as e:
        logger.error(f"Error getting emails from folder: {e}")
        import traceback
        traceback.print_exc()
        return [], f"Error: Failed to get emails from folder '{folder_name}': {e}"


def get_emails_from_folder(folder_name: str = "Inbox", days: int = 7) -> Tuple[List[Dict[str, Any]], str]:
    """Backward compatibility wrapper - calls the optimized version."""
    return get_emails_from_folder_optimized(folder_name, days)