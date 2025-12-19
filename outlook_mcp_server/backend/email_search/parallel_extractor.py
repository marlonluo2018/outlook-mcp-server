"""
Parallel email extraction for performance optimization.
"""
import logging
from typing import List, Dict, Any
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import threading

logger = logging.getLogger(__name__)

# Thread-local storage for COM objects
_thread_local = threading.local()

def _extract_email_info_parallel(item_data: Dict[str, Any]) -> Dict[str, Any]:
    """Extract email info from item data in a thread-safe manner."""
    try:
        # Extract basic attributes only - this is the minimal version
        entry_id = item_data.get('EntryID', '')
        subject = item_data.get('Subject', 'No Subject')
        sender = item_data.get('SenderName', 'Unknown')
        received_time = item_data.get('ReceivedTime', None)
        
        return {
            "entry_id": entry_id,
            "subject": subject,
            "sender": sender,
            "received_time": str(received_time) if received_time else "Unknown"
        }
    except Exception as e:
        logger.debug(f"Error in parallel extraction: {e}")
        return {
            "entry_id": item_data.get('EntryID', ''),
            "subject": "No Subject",
            "sender": "Unknown",
            "received_time": "Unknown"
        }

def extract_emails_parallel(items: List[Any], max_workers: int = 4) -> List[Dict[str, Any]]:
    """
    Extract email information from a list of Outlook items using parallel processing.
    
    Args:
        items: List of Outlook MailItem objects
        max_workers: Maximum number of worker threads
        
    Returns:
        List of email dictionaries
    """
    if not items:
        return []
    
    try:
        # Convert items to dictionaries first to avoid COM threading issues
        logger.info(f"Converting {len(items)} items to dictionaries for parallel processing")
        
        item_dicts = []
        for item in items:
            try:
                item_dict = {
                    'EntryID': getattr(item, 'EntryID', ''),
                    'Subject': getattr(item, 'Subject', 'No Subject'),
                    'SenderName': getattr(item, 'SenderName', 'Unknown'),
                    'ReceivedTime': getattr(item, 'ReceivedTime', None)
                }
                item_dicts.append(item_dict)
            except Exception as e:
                logger.debug(f"Error converting item to dict: {e}")
                continue
        
        logger.info(f"Processing {len(item_dicts)} items in parallel with {max_workers} workers")
        
        # Process items in parallel
        email_list = []
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_item = {executor.submit(_extract_email_info_parallel, item_dict): item_dict 
                             for item_dict in item_dicts}
            
            # Collect results as they complete
            for future in as_completed(future_to_item):
                try:
                    email_data = future.result()
                    if email_data and email_data.get("entry_id"):
                        email_list.append(email_data)
                except Exception as e:
                    logger.debug(f"Error processing item in parallel: {e}")
                    continue
        
        logger.info(f"Parallel extraction completed: {len(email_list)} emails extracted")
        return email_list
        
    except Exception as e:
        logger.error(f"Error in parallel extraction: {e}")
        # Fallback to sequential processing
        return extract_emails_sequential_fallback(items)

def extract_emails_sequential_fallback(items: List[Any]) -> List[Dict[str, Any]]:
    """Optimized sequential extraction for small datasets with minimal overhead."""
    email_list = []
    
    # Pre-allocate list for better performance if size is known
    if hasattr(items, '__len__'):
        email_list = [None] * len(items)
        index = 0
    
    for item in items:
        try:
            # Minimal attribute access with error handling
            entry_id = getattr(item, 'EntryID', '')
            if not entry_id:
                continue
                
            subject = getattr(item, 'Subject', 'No Subject') or 'No Subject'
            sender = getattr(item, 'SenderName', 'Unknown') or 'Unknown'
            
            received_time = getattr(item, 'ReceivedTime', None)
            received_str = str(received_time) if received_time else "Unknown"
            
            email_data = {
                "entry_id": entry_id,
                "subject": subject,
                "sender": sender,
                "received_time": received_str
            }
            
            if hasattr(items, '__len__'):
                email_list[index] = email_data
                index += 1
            else:
                email_list.append(email_data)
                
        except Exception:
            # Silent fail for performance - skip problematic items
            continue
    
    # Remove None values if pre-allocation was used
    if hasattr(items, '__len__') and index < len(email_list):
        email_list = email_list[:index]
    
    return email_list

def extract_emails_optimized(items: List[Any], use_parallel: bool = True, max_workers: int = 4) -> List[Dict[str, Any]]:
    """
    Optimized email extraction with automatic fallback and improved small dataset handling.
    
    Args:
        items: List of Outlook MailItem objects
        use_parallel: Whether to use parallel processing
        max_workers: Maximum number of worker threads (if parallel)
        
    Returns:
        List of email dictionaries
    """
    if not items:
        return []
    
    item_count = len(items)
    
    # Optimized thresholds for better performance
    if item_count < 20:  # Very small datasets: sequential is definitely faster
        return extract_emails_sequential_fallback(items)
    elif item_count < 50:  # Small datasets: use sequential with minimal overhead
        return extract_emails_sequential_fallback(items)
    elif item_count < 100:  # Medium datasets: use sequential or light parallel
        return extract_emails_sequential_fallback(items)
    else:  # Large datasets: use parallel processing
        if use_parallel:
            return extract_emails_parallel(items, max_workers)
        else:
            return extract_emails_sequential_fallback(items)