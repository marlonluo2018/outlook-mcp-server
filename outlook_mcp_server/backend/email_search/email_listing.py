"""
Email listing functionality for retrieving emails from folders.

This module provides functions for listing recent emails and getting emails
from specific folders with various filtering options.
"""

import logging
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Tuple

from ..outlook_session.session_manager import OutlookSessionManager
from ..shared import email_cache, email_cache_order, add_email_to_cache, clear_email_cache
from ..validators import EmailListParams
from .search_common import extract_email_info, get_folder_path_safe

# Set up logging
logger = logging.getLogger(__name__)


def list_recent_emails(folder_name: str = "Inbox", days: int = None) -> Tuple[List[Dict[str, Any]], str]:
    """Public interface for listing emails (used by CLI).
    Loads emails into cache and returns (emails, message) tuple.
    """
    try:
        # Default to 365 days if not specified to ensure we get results
        effective_days = days or 365
        params = EmailListParams(days=effective_days, folder_name=folder_name)
        
        logger.info(f"list_recent_emails called with folder={folder_name}, days={days}, effective_days={effective_days}")
    except Exception as e:
        logger.error(f"Validation error in list_recent_emails: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    # Clear cache before loading new emails to ensure fresh results
    logger.info(f"Cache before clearing: {len(email_cache)} emails, {len(email_cache_order)} order")
    clear_email_cache()
    logger.info(f"Cache after clearing: {len(email_cache)} emails, {len(email_cache_order)} order")

    emails, note = get_emails_from_folder_optimized(folder_name=params.folder_name, days=params.days)
    
    logger.info(f"get_emails_from_folder returned {len(emails)} emails, note: {note}")

    # UVX COMPATIBILITY: Ensure cache is saved immediately after loading emails
    # Only save to cache if we actually got emails successfully (no error)
    if emails and "Error:" not in note:
        from ..shared import immediate_save_cache
        immediate_save_cache()

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
            
            logger.info(f"Getting emails from folder '{params.folder_name}' for {params.days} days")
            
            # Get total item count first (needed for batch size calculation)
            total_items = folder.Items.Count if hasattr(folder.Items, 'Count') else 0
            logger.info(f"Total items in folder: {total_items}")
            
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
            logger.info(f"Adjusted max_items limit for {params.days} days: {max_items}")
            
            # Calculate date limit if needed
            date_limit = None
            if params.days:
                date_limit = datetime.now(timezone.utc) - timedelta(days=params.days)
                logger.info(f"Date filter: items from {date_limit.strftime('%Y-%m-%d')} onwards")
            
            # OPTIMIZATION 3: Process items in reverse order (newest first) for early termination
            items_to_process = min(total_items, max_items)
            logger.info(f"Processing {items_to_process} items in batches of {batch_size}")
            
            # OPTIMIZATION 4: Cache folder.Items to avoid repeated COM calls
            items_collection = folder.Items
            
            # Track if we should stop early (when emails are too old)
            early_termination = False
            consecutive_old_emails = 0  # OPTIMIZATION: Enhanced early termination counter
            total_recent_emails = 0  # Track total recent emails found
            
            for batch_start in range(0, items_to_process, batch_size):
                if early_termination:
                    break
                    
                batch_end = min(batch_start + batch_size, items_to_process)
                logger.debug(f"Processing batch {batch_start+1} to {batch_end}")
                
                batch_filtered = []
                
                for i in range(batch_start, batch_end):
                    try:
                        # OPTIMIZATION 5: Get item in reverse order (newest first)
                        item_index = total_items - i
                        if item_index <= 0:
                            continue
                            
                        # OPTIMIZATION 6: Single COM access call
                        item = items_collection.Item(item_index)
                        if not item:
                            continue
                        
                        # OPTIMIZATION 7: Early date check before other validations
                        if date_limit and hasattr(item, 'ReceivedTime') and item.ReceivedTime:
                            try:
                                item_time = item.ReceivedTime
                                if item_time.tzinfo is None:
                                    item_time = item_time.replace(tzinfo=timezone.utc)
                                
                                # OPTIMIZATION 8: Enhanced early termination if email is too old
                                # Since we process newest first, we can stop when we find consistently old emails
                                if item_time < date_limit:
                                    consecutive_old_emails += 1
                                    # OPTIMIZATION: Increased threshold from 3 to 10 and require some recent emails first
                                    if consecutive_old_emails >= 10 and total_recent_emails > 0:  # Only terminate after finding 10 consecutive old emails and some recent emails
                                        early_termination = True
                                        logger.info(f"Early termination: Found {consecutive_old_emails} consecutive emails older than {params.days} days after {total_recent_emails} recent emails at position {i}")
                                        break
                                    continue
                                else:
                                    consecutive_old_emails = 0  # Reset counter when we find a recent email
                            except Exception:
                                continue
                        
                        # OPTIMIZATION 9: Consolidated validation checks
                        if not hasattr(item, 'Class') or item.Class != 43:  # 43 = olMail
                            continue
                        
                        if not item.ReceivedTime:
                            continue
                        
                        batch_filtered.append(item)
                        
                    except Exception as e:
                        logger.debug(f"Error processing item {item_index}: {e}")
                        continue
                
                if early_termination:
                    break
                    
                filtered_items.extend(batch_filtered)
                logger.debug(f"Batch completed: {len(batch_filtered)} items added, total: {len(filtered_items)}")
            
            logger.info(f"Items after date filtering: {len(filtered_items)}")
            
            if not filtered_items:
                return [], f"No emails found in '{params.folder_name}' from last {params.days} days"
            
            # OPTIMIZATION 10: Skip sorting if already in correct order (newest first)
            # Since we process in reverse order, items should already be newest first
            logger.info(f"Processing {len(filtered_items)} emails for caching")
            
            # OPTIMIZATION 11: Enhanced batch processing with bulk timestamp handling
            email_list = []
            cache_count = 0
            
            # Pre-process timestamps in bulk for better cache performance
            timestamp_cache = {}
            for item in filtered_items:
                try:
                    received_time = getattr(item, 'ReceivedTime', None)
                    if received_time:
                        entry_id = getattr(item, 'EntryID', '')
                        timestamp_cache[entry_id] = received_time
                except Exception:
                    continue
            
            for item in filtered_items:
                try:
                    # OPTIMIZATION 12: Streamlined email extraction with cached timestamps
                    email_data = extract_email_info(item)
                    
                    # Use pre-cached timestamp for better performance
                    entry_id = email_data.get("entry_id", "")
                    if entry_id in timestamp_cache:
                        email_data["received_time"] = str(timestamp_cache[entry_id])
                    
                    add_email_to_cache(email_data["entry_id"], email_data)
                    email_list.append(email_data)
                    cache_count += 1
                    
                    # Log progress for large batches
                    if cache_count % 100 == 0:
                        logger.info(f"Cached {cache_count}/{len(filtered_items)} emails")
                        
                except Exception as e:
                    logger.warning(f"Failed to cache email: {e}")
                    continue
            
            logger.info(f"Successfully cached {len(email_list)} emails")
            
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