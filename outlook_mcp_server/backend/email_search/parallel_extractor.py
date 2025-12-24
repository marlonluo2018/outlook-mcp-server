"""
Parallel email extraction for performance optimization.
"""

# Standard library imports
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Any, Dict, List

# Local application imports
from ..logging_config import get_logger

logger = get_logger(__name__)

# Thread-local storage for COM objects
_thread_local = threading.local()

def _extract_email_info_parallel(item_data: Dict[str, Any]) -> Dict[str, Any]:
    """Extract email info from item data in a thread-safe manner."""
    try:
        # Extract basic attributes
        entry_id = item_data.get('EntryID', '')
        subject = item_data.get('Subject', 'No Subject')
        sender = item_data.get('SenderName', 'Unknown')
        received_time = item_data.get('ReceivedTime', None)
        
        # Extract recipients - handle both formats
        to_recipients = item_data.get('to_recipients', [])
        cc_recipients = item_data.get('cc_recipients', [])
        
        # If recipients are not already extracted, try to extract from To/CC fields
        if not to_recipients and item_data.get('To'):
            to_field = str(item_data.get('To', ''))
            if to_field:
                to_list = to_field.split(';')
                to_recipients = [{"address": addr.strip(), "name": addr.strip()} for addr in to_list if addr.strip()]
        
        if not cc_recipients and item_data.get('CC'):
            cc_field = str(item_data.get('CC', ''))
            if cc_field:
                cc_list = cc_field.split(';')
                cc_recipients = [{"address": addr.strip(), "name": addr.strip()} for addr in cc_list if addr.strip()]
        
        # Extract attachment info
        has_attachments = item_data.get('has_attachments', False)
        attachments = item_data.get('attachments', [])
        embedded_images_count = item_data.get('embedded_images_count', 0)
        
        return {
            "entry_id": entry_id,
            "subject": subject,
            "sender": sender,
            "received_time": str(received_time) if received_time else "Unknown",
            "to_recipients": to_recipients,
            "cc_recipients": cc_recipients,
            "has_attachments": has_attachments,
            "attachments": attachments,
            "attachments_count": len(attachments),
            "embedded_images_count": embedded_images_count,
            "unread": item_data.get('UnRead', False)
        }
    except Exception as e:
        logger.debug(f"Error in parallel extraction: {e}")
        return {
            "entry_id": item_data.get('EntryID', ''),
            "subject": "No Subject",
            "sender": "Unknown",
            "received_time": "Unknown",
            "to_recipients": [],
            "cc_recipients": [],
            "has_attachments": False,
            "attachments": [],
            "attachments_count": 0,
            "unread": False
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
                    'ReceivedTime': getattr(item, 'ReceivedTime', None),
                    'To': getattr(item, 'To', ''),
                    'CC': getattr(item, 'CC', ''),
                    'UnRead': getattr(item, 'UnRead', False)
                }
                
                # Extract attachment info with embedded image detection
                try:
                    attachments = getattr(item, 'Attachments', None)
                    if attachments:
                        attachments_list = []
                        embedded_images_count = 0
                        
                        for att in attachments:
                            try:
                                file_name = getattr(att, 'FileName', '') or getattr(att, 'DisplayName', 'Unknown')
                                is_image = file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico'))
                                
                                # Check if it's an embedded image using multiple methods
                                is_embedded = False
                                
                                # Method 1: Check Content-ID property
                                try:
                                    content_id = getattr(att, 'PropertyAccessor', None)
                                    if content_id:
                                        cid = content_id.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                                        if cid and cid.strip():
                                            is_embedded = True
                                except Exception:
                                    pass
                                
                                # Method 2: Check if filename contains CID-like patterns
                                if not is_embedded and is_image:
                                    if 'cid:' in file_name.lower() or file_name.startswith('image'):
                                        is_embedded = True
                                
                                # Method 3: Check attachment type
                                try:
                                    att_type = getattr(att, 'Type', 1)
                                    if att_type == 6:  # Embedded message
                                        is_embedded = True
                                except Exception:
                                    pass
                                
                                # Count embedded images
                                if is_embedded and is_image:
                                    embedded_images_count += 1
                                else:
                                    # Only add non-embedded attachments to the list
                                    attachments_list.append({
                                        'filename': file_name,
                                        'size': getattr(att, 'Size', 0)
                                    })
                                
                            except Exception:
                                continue
                        
                        item_dict['has_attachments'] = len(attachments_list) > 0
                        item_dict['attachments'] = attachments_list
                        item_dict['embedded_images_count'] = embedded_images_count
                    else:
                        item_dict['has_attachments'] = False
                        item_dict['attachments'] = []
                        item_dict['embedded_images_count'] = 0
                except Exception:
                    item_dict['has_attachments'] = False
                    item_dict['attachments'] = []
                    item_dict['embedded_images_count'] = 0
                
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
            
            # Extract recipient information
            to_field = getattr(item, 'To', '')
            cc_field = getattr(item, 'CC', '')
            
            # Parse recipients from To field
            to_recipients = []
            if to_field:
                try:
                    to_list = str(to_field).split(';')
                    to_recipients = [{"address": addr.strip(), "name": addr.strip()} for addr in to_list if addr.strip()]
                except Exception:
                    to_recipients = []
            
            # Parse recipients from CC field
            cc_recipients = []
            if cc_field:
                try:
                    cc_list = str(cc_field).split(';')
                    cc_recipients = [{"address": addr.strip(), "name": addr.strip()} for addr in cc_list if addr.strip()]
                except Exception:
                    cc_recipients = []
            
            # Extract attachment info with embedded image detection
            has_attachments = False
            attachments = []
            embedded_images_count = 0
            try:
                attachments_obj = getattr(item, 'Attachments', None)
                if attachments_obj:
                    has_attachments = attachments_obj.Count > 0
                    attachments_list = []
                    
                    for att in attachments_obj:
                        try:
                            file_name = getattr(att, 'FileName', '') or getattr(att, 'DisplayName', 'Unknown')
                            is_image = file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico'))
                            
                            # Check if it's an embedded image using multiple methods
                            is_embedded = False
                            
                            # Method 1: Check Content-ID property
                            try:
                                content_id = getattr(att, 'PropertyAccessor', None)
                                if content_id:
                                    cid = content_id.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                                    if cid and cid.strip():
                                        is_embedded = True
                            except Exception:
                                pass
                            
                            # Method 2: Check if filename contains CID-like patterns
                            if not is_embedded and is_image:
                                if 'cid:' in file_name.lower() or file_name.startswith('image'):
                                    is_embedded = True
                            
                            # Method 3: Check attachment type
                            try:
                                att_type = getattr(att, 'Type', 1)
                                if att_type == 6:  # Embedded message
                                    is_embedded = True
                            except Exception:
                                pass
                            
                            # Count embedded images
                            if is_embedded and is_image:
                                embedded_images_count += 1
                            else:
                                # Only add non-embedded attachments to the list
                                attachments_list.append({
                                    'filename': file_name,
                                    'size': getattr(att, 'Size', 0)
                                })
                            
                        except Exception:
                            continue
                    
                    attachments = attachments_list
            except Exception:
                has_attachments = False
                attachments = []
                embedded_images_count = 0
            
            # Extract unread status
            unread = getattr(item, 'UnRead', False)
            
            email_data = {
                "entry_id": entry_id,
                "subject": subject,
                "sender": sender,
                "received_time": received_str,
                "to_recipients": to_recipients,
                "cc_recipients": cc_recipients,
                "has_attachments": has_attachments,
                "attachments": attachments,
                "attachments_count": len(attachments),
                "embedded_images_count": embedded_images_count,
                "unread": unread
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