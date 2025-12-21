"""Simplified email data extraction with single comprehensive mode."""

import logging
from typing import Dict, Any, Optional

from .outlook_session.session_manager import OutlookSessionManager
from .utils import OutlookItemClass, safe_encode_text
from .email_utils import _format_recipient_for_display
from .shared import email_cache, email_cache_order

logger = logging.getLogger(__name__)


def extract_comprehensive_email_data(email: Dict[str, Any]) -> Dict[str, Any]:
    """Extract comprehensive email data with single mode - always return full text content."""
    
    # Start with basic email data
    sender = email.get("sender", "Unknown Sender")
    if isinstance(sender, dict):
        sender_name = sender.get("name", "Unknown Sender")
    else:
        sender_name = str(sender)
    
    result = {
        "id": email.get("id", email.get("entry_id", "")),
        "entry_id": email.get("id", email.get("entry_id", "")),
        "subject": email.get("subject", "No Subject"),
        "sender": sender_name,
        "from": sender_name,  # Alias for compatibility
        "received_time": email.get("received_time", ""),
        "received": email.get("received_time", ""),  # Alias for compatibility
        "unread": email.get("unread", False),
        "has_attachments": email.get("has_attachments", False),
        "size": email.get("size", 0),
        "to": (
            ", ".join([_format_recipient_for_display(r) for r in email.get("to_recipients", [])])
            if email.get("to_recipients")
            else ""
        ),
        "cc": (
            ", ".join([_format_recipient_for_display(r) for r in email.get("cc_recipients", [])])
            if email.get("cc_recipients")
            else ""
        ),
        "body": email.get("body", ""),  # Include cached body if available
        "attachments": email.get("attachments", []),  # Include cached attachments if available
        "attachments_count": len(email.get("attachments", [])),  # Count of real attachments
    }
    
    # Always attempt to get comprehensive content from Outlook
    try:
        with OutlookSessionManager() as session:
            if not session or not session.namespace:
                logger.error("Failed to establish Outlook session")
                return result
                
            if not hasattr(session.namespace, 'GetItemFromID'):
                logger.error("Namespace does not have GetItemFromID method")
                return result
                
            item = session.namespace.GetItemFromID(email.get("entry_id", email.get("id", "")))
            if not item or item.Class != OutlookItemClass.MAIL_ITEM:
                logger.warning(f"Email not found or not a mail item")
                return result

            # Extract all available text content
            result["body"] = safe_encode_text(getattr(item, "Body", ""), "body")
            result["html_body"] = safe_encode_text(getattr(item, "HTMLBody", ""), "html_body") if hasattr(item, "HTMLBody") else ""
            result["body_format"] = getattr(item, "BodyFormat", 1)  # 1=Plain, 2=HTML, 3=RichText
            
            # Extract attachment details if not already cached
            if hasattr(item, 'Attachments') and item.Attachments and item.Attachments.Count > 0:
                attachments = []
                try:
                    for i in range(item.Attachments.Count):
                        attachment = item.Attachments.Item(i + 1)
                        file_name = getattr(attachment, 'FileName', getattr(attachment, 'DisplayName', 'Unknown'))
                        
                        # Check if it's an embedded image
                        is_embedded = False
                        
                        # Method 1: Check Content-ID property (most reliable for embedded images)
                        try:
                            if hasattr(attachment, 'PropertyAccessor'):
                                content_id = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                                is_embedded = content_id is not None and len(str(content_id).strip()) > 0
                                
                                # Also check for Content-Location property (another indicator of embedded content)
                                if not is_embedded:
                                    try:
                                        content_location = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3713001F")
                                        is_embedded = content_location is not None and len(str(content_location).strip()) > 0
                                    except:
                                        pass
                        except:
                            pass
                        
                        # Method 2: Check attachment type - embedded attachments are usually Type 4 (OLE) or Type 3 (Embedded)
                        attachment_type = getattr(attachment, 'Type', 1)
                        if attachment_type in [3, 4]:  # 3 = Embedded, 4 = OLE
                            is_embedded = True
                        
                        # Method 3: Check if it's an image with suspicious naming patterns
                        is_image = file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico'))
                        if is_image and not is_embedded:
                            # Check for common embedded image naming patterns
                            lower_name = file_name.lower()
                            if any(pattern in lower_name for pattern in ['image', 'img', 'cid:', 'embedded']):
                                is_embedded = True
                            # Check if filename is just numbers with extension (common for embedded images)
                            elif '.' in lower_name:
                                name_without_ext = lower_name.rsplit('.', 1)[0]
                                if name_without_ext.isdigit() or (len(name_without_ext) <= 2 and name_without_ext.isalnum()):
                                    is_embedded = True
                        
                        # Method 4: Check if attachment size is suspiciously small for an image (embedded images are often smaller)
                        if is_image and not is_embedded:
                            try:
                                attachment_size = getattr(attachment, 'Size', 0)
                                # Embedded images are typically smaller than regular image attachments
                                if attachment_size > 0 and attachment_size < 50000:  # Less than 50KB
                                    # Additional check: if it's a very small image, more likely to be embedded
                                    if attachment_size < 10000:  # Less than 10KB
                                        is_embedded = True
                            except:
                                pass
                        
                        # PDF files and other documents are always considered real attachments
                        is_document = file_name.lower().endswith(('.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.zip', '.rar'))
                        if is_document:
                            is_embedded = False
                        
                        # Only add non-embedded attachments to the list
                        if not is_embedded:
                            attachment_info = {
                                "name": file_name,
                                "size": getattr(attachment, 'Size', 0),
                                "type": getattr(attachment, 'Type', 1)  # 1 = ByValue, 2 = ByReference, 3 = Embedded, 4 = OLE
                            }
                            attachments.append(attachment_info)
                    
                    # Update has_attachments flag and attachments list
                    result["attachments"] = attachments
                    result["has_attachments"] = len(attachments) > 0
                    result["attachments_count"] = len(attachments)
                except Exception as e:
                    logger.debug(f"Error extracting attachment details: {e}")
                    result["attachments"] = []
                    result["has_attachments"] = False
                    result["attachments_count"] = 0
            
            # Enhanced metadata
            result["importance"] = getattr(item, "Importance", 1)  # 0=Low, 1=Normal, 2=High
            result["sensitivity"] = getattr(item, "Sensitivity", 0)  # 0=Normal, 1=Personal, 2=Private, 3=Confidential
            result["conversation_topic"] = safe_encode_text(getattr(item, "ConversationTopic", ""), "conversation_topic")
            result["conversation_id"] = getattr(item, "ConversationID", "")
            result["categories"] = getattr(item, "Categories", "")
            result["flag_status"] = getattr(item, "FlagStatus", 0)  # 0=Unflagged, 1=Flagged, 2=Complete
            
    except Exception as e:
        logger.error(f"Error loading email details: {e}")
        # Return basic data on error
        pass
    
    return result


def extract_basic_email_data(email: Dict[str, Any]) -> Dict[str, Any]:
    """Extract email data with full text but without embedded images and attachments (renamed from text_only)."""
    # Start with comprehensive data but filter out attachments and embedded images
    comprehensive_data = extract_comprehensive_email_data(email)
    
    # Remove attachments and embedded images
    comprehensive_data["attachments"] = []
    comprehensive_data["has_attachments"] = False
    
    # Keep all text content but ensure no embedded images in HTML body
    if comprehensive_data.get("html_body"):
        # Simple regex to remove img tags (basic HTML cleaning)
        import re
        comprehensive_data["html_body"] = re.sub(r'<img[^>]*>', '', comprehensive_data["html_body"])
    
    return comprehensive_data


# Removed text_only function and get_conversation_thread function - now basic mode uses the text-only logic


def create_basic_email_response(email: Dict[str, Any]) -> Dict[str, Any]:
    """Create basic email response from cached data only."""
    sender = email.get("sender", "Unknown Sender")
    if isinstance(sender, dict):
        sender_name = sender.get("name", "Unknown Sender")
    else:
        sender_name = str(sender)

    return {
        "id": email.get("id", ""),
        "subject": email.get("subject", "No Subject"),
        "sender": sender_name,
        "received_time": email.get("received_time", ""),
        "unread": email.get("unread", False),
        "has_attachments": email.get("has_attachments", False),
        "size": email.get("size", 0),
        "body": email.get("body", ""),
        "to": (
            ", ".join([_format_recipient_for_display(r) for r in email.get("to_recipients", [])])
            if email.get("to_recipients")
            else ""
        ),
        "cc": (
            ", ".join([_format_recipient_for_display(r) for r in email.get("cc_recipients", [])])
            if email.get("cc_recipients")
            else ""
        ),
        "attachments": email.get("attachments", []),
    }


def get_email_by_number_unified(email_number: int, mode: str = "basic", include_attachments: bool = True, embed_images: bool = True) -> Optional[Dict[str, Any]]:
    """Get email by number from cache with unified interface.
    
    Args:
        email_number: The number of the email in the cache (1-based)
        mode: Retrieval mode - "basic", "enhanced", "lazy"
        include_attachments: Whether to include attachment content
        embed_images: Whether to embed inline images
        
    Returns:
        Email data dictionary or None if not found
    """
    if not isinstance(email_number, int) or email_number < 1:
        return None
        
    # Check if cache is loaded
    if not email_cache or not email_cache_order:
        return None
        
    # Validate email number
    if email_number > len(email_cache_order):
        return None
        
    # Get email ID from cache order
    email_id = email_cache_order[email_number - 1]
    
    # Get email from cache
    email_data = email_cache.get(email_id)
    if not email_data:
        return None
        
    # Extract email data based on mode
    if mode == "basic":
        return extract_basic_email_data(email_data)  # This is the new text-only mode (renamed)
    else:  # enhanced, lazy, or any other mode
        return extract_comprehensive_email_data(email_data)


def format_email_with_media(email_data: Dict[str, Any]) -> str:
    """Format email with media information for enhanced display."""
    formatted_text = f"Subject: {email_data.get('subject', 'N/A')}\n"
    formatted_text += f"From: {email_data.get('from', 'N/A')}\n"
    formatted_text += f"To: {email_data.get('to', 'N/A')}\n"
    formatted_text += f"Date: {email_data.get('received', 'N/A')}\n"
    
    # Add conversation topic if available
    if email_data.get("conversation_topic"):
        formatted_text += f"Conversation: {email_data.get('conversation_topic', 'N/A')}\n"
    
    formatted_text += f"Body: {email_data.get('body', 'N/A')}\n"
    
    # Add HTML body if available and different from plain body
    if email_data.get("html_body") and email_data.get("html_body") != email_data.get("body"):
        formatted_text += f"HTML Body: {email_data.get('html_body', 'N/A')}\n"
    
    # Add attachments if present and mode allows it
    if email_data.get("attachments") and email_data.get("has_attachments", False):
        formatted_text += f"\nAttachments: {len(email_data['attachments'])}\n"
        for attachment in email_data["attachments"]:
            formatted_text += f"  - {attachment.get('name', 'Unknown')}"
            if attachment.get('size'):
                formatted_text += f" ({attachment['size']} bytes)"
            if attachment.get('content_base64'):
                content_length = len(attachment['content_base64'])
                formatted_text += f" [Base64 content: {content_length} characters]"
            formatted_text += "\n"
    
    # Add metadata if available
    if email_data.get("importance") is not None:
        importance_map = {0: "Low", 1: "Normal", 2: "High"}
        formatted_text += f"Importance: {importance_map.get(email_data.get('importance'), 'Normal')}\n"
    
    if email_data.get("categories"):
        formatted_text += f"Categories: {email_data.get('categories', 'N/A')}\n"
    
    return formatted_text