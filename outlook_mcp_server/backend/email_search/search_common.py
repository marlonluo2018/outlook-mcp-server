"""
Common utilities for email search operations.

This module contains shared functions and utilities used across different
email search implementations.
"""

import logging
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

# Set up logging
logger = logging.getLogger(__name__)


def get_folder_path_safe(folder_name: Optional[str] = None) -> str:
    """Get safe folder path, defaulting to Inbox if not provided."""
    return folder_name if folder_name else "Inbox"


def get_date_limit(days: int) -> datetime:
    """Get the date limit for searching emails."""
    from datetime import timezone
    return datetime.now(timezone.utc) - timedelta(days=days)


def is_server_search_supported(search_type: str) -> bool:
    """Check if server-side search is supported for the given search type."""
    return search_type in ["subject", "sender", "recipient"]


# COM attribute cache to avoid repeated access
_com_attribute_cache = {}

def _get_cached_com_attribute(item, attr_name, default=None):
    """Get COM attribute with caching to avoid repeated access."""
    try:
        item_id = getattr(item, 'EntryID', '')
        if not item_id:
            return getattr(item, attr_name, default)
            
        cache_key = f"{item_id}:{attr_name}"
        if cache_key not in _com_attribute_cache:
            _com_attribute_cache[cache_key] = getattr(item, attr_name, default)
        return _com_attribute_cache[cache_key]
    except Exception:
        return default

def extract_email_info(item) -> Dict[str, Any]:
    """Extract basic email information from an Outlook item with optimized COM access."""
    # Single-pass COM attribute extraction with caching
    email_info = {
        "subject": _get_cached_com_attribute(item, 'Subject', 'No Subject'),
        "sender": _get_cached_com_attribute(item, 'SenderName', 'Unknown'),
        "received_time": _get_cached_com_attribute(item, 'ReceivedTime', None),
        "entry_id": _get_cached_com_attribute(item, 'EntryID', ''),
    }
    
    # Handle None received_time
    if email_info["received_time"] is None:
        email_info["received_time"] = "Unknown"
    else:
        email_info["received_time"] = str(email_info["received_time"])
    
    # Extract To recipients - optimized with single-pass COM access
    try:
        to_recipients = []
        
        # Use cached recipients collection access
        recipients = _get_cached_com_attribute(item, 'Recipients')
        if recipients:
            try:
                for recipient in recipients:
                    if _get_cached_com_attribute(recipient, 'Type') == 1:  # 1 = To recipient
                        recipient_info = {
                            "address": _get_cached_com_attribute(recipient, 'Address', ''),
                            "name": _get_cached_com_attribute(recipient, 'Name', '')
                        }
                        if recipient_info["address"] or recipient_info["name"]:
                            to_recipients.append(recipient_info)
            except Exception as e:
                logger.debug(f"Error extracting from Recipients collection: {e}")
        
        # Fallback to To field if Recipients collection didn't work
        if not to_recipients:
            to_field = _get_cached_com_attribute(item, 'To')
            if to_field:
                try:
                    # Parse To field which might be a semicolon-separated string
                    to_list = str(to_field).split(';')
                    for to_addr in to_list:
                        to_addr = to_addr.strip()
                        if to_addr:
                            to_recipients.append({"address": to_addr, "name": to_addr})
                except Exception as e:
                    logger.debug(f"Error extracting from To field: {e}")
        
        email_info["to_recipients"] = to_recipients
    except Exception as e:
        logger.debug(f"Error in To recipient extraction: {e}")
        email_info["to_recipients"] = []
    
    # Extract CC recipients - optimized with single-pass COM access
    try:
        cc_recipients = []
        
        # Use cached recipients collection access
        recipients = _get_cached_com_attribute(item, 'Recipients')
        if recipients:
            try:
                for recipient in recipients:
                    if _get_cached_com_attribute(recipient, 'Type') == 2:  # 2 = CC recipient
                        recipient_info = {
                            "address": _get_cached_com_attribute(recipient, 'Address', ''),
                            "name": _get_cached_com_attribute(recipient, 'Name', '')
                        }
                        if recipient_info["address"] or recipient_info["name"]:
                            cc_recipients.append(recipient_info)
            except Exception as e:
                logger.debug(f"Error extracting CC from Recipients collection: {e}")
        
        # Fallback to CC field if Recipients collection didn't work
        if not cc_recipients:
            cc_field = _get_cached_com_attribute(item, 'CC')
            if cc_field:
                try:
                    # Parse CC field which might be a semicolon-separated string
                    cc_list = str(cc_field).split(';')
                    for cc_addr in cc_list:
                        cc_addr = cc_addr.strip()
                        if cc_addr:
                            cc_recipients.append({"address": cc_addr, "name": cc_addr})
                except Exception as e:
                    logger.debug(f"Error extracting from CC field: {e}")
        
        email_info["cc_recipients"] = cc_recipients
    except Exception as e:
        logger.debug(f"Error in CC recipient extraction: {e}")
        email_info["cc_recipients"] = []
    
    # Extract additional useful information with optimized COM access
    try:
        email_info["unread"] = _get_cached_com_attribute(item, 'UnRead', False)
        attachments = _get_cached_com_attribute(item, 'Attachments')
        has_attachments = attachments and hasattr(attachments, 'Count') and attachments.Count > 0
        email_info["has_attachments"] = has_attachments
        
        # Extract attachment information if present
        if has_attachments:
            attachments_list = []
            try:
                for i in range(attachments.Count):
                    attachment = attachments.Item(i + 1)
                    file_name = _get_cached_com_attribute(attachment, 'FileName') or _get_cached_com_attribute(attachment, 'DisplayName', 'Unknown')
                    
                    # Check if it's an embedded image
                    is_embedded = False
                    try:
                        property_accessor = _get_cached_com_attribute(attachment, 'PropertyAccessor')
                        if property_accessor:
                            content_id = property_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                            is_embedded = content_id is not None and len(str(content_id)) > 0
                    except:
                        pass
                    
                    # Additional check for image files with common embedded naming pattern
                    if not is_embedded:
                        is_image = file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))
                        if is_image and file_name.lower().startswith('image'):
                            is_embedded = True
                    
                    # PDF files are always considered real attachments
                    is_pdf = file_name.lower().endswith('.pdf')
                    if is_pdf:
                        is_embedded = False
                    
                    # Only add non-embedded attachments to the list
                    if not is_embedded:
                        attachment_info = {
                            "name": file_name,
                            "size": _get_cached_com_attribute(attachment, 'Size', 0),
                            "type": _get_cached_com_attribute(attachment, 'Type', 1)  # 1 = ByValue, 2 = ByReference, 3 = Embedded, 4 = OLE
                        }
                        attachments_list.append(attachment_info)
                
                # Update has_attachments flag based on real attachments only
                email_info["has_attachments"] = len(attachments_list) > 0
                email_info["attachments"] = attachments_list
            except Exception as e:
                logger.debug(f"Error extracting attachment details: {e}")
                email_info["attachments"] = []
                email_info["has_attachments"] = False
        else:
            email_info["attachments"] = []
            email_info["has_attachments"] = False
    except Exception as e:
        logger.debug(f"Error extracting email metadata: {e}")
        email_info["unread"] = False
        email_info["has_attachments"] = False
        email_info["attachments"] = []
    
    return email_info