"""Unified email retrieval module combining basic and enhanced functionality."""

import base64
import os
import tempfile
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple, Union

from .outlook_session import OutlookSessionManager
from .shared import email_cache, save_email_cache, LAZY_LOADING_ENABLED
from .utils import OutlookItemClass, safe_encode_text
from .validators import EmailSearchParams

logger = logging.getLogger(__name__)

# Configuration constants
MAX_ATTACHMENT_SIZE = 10 * 1024 * 1024  # 10MB limit for inline content
EMBEDDABLE_IMAGE_TYPES = {'image/jpeg', 'image/png', 'image/gif', 'image/bmp', 'image/x-icon'}
TEXT_MIME_TYPES = {'text/plain', 'text/html', 'text/css', 'application/json', 'application/xml', 'text/csv'}


class EmailRetrievalMode:
    """Email retrieval modes for different use cases."""
    BASIC = "basic"  # Original basic functionality
    ENHANCED = "enhanced"  # Full media support
    LAZY = "lazy"  # Lazy loading for performance


def get_mime_type(filename: str) -> str:
    """Determine MIME type from file extension."""
    ext = Path(filename).suffix.lower()
    mime_types = {
        '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
        '.png': 'image/png', '.gif': 'image/gif',
        '.bmp': 'image/bmp', '.ico': 'image/x-icon',
        '.txt': 'text/plain', '.html': 'text/html',
        '.htm': 'text/html', '.css': 'text/css',
        '.js': 'application/javascript', '.json': 'application/json',
        '.xml': 'application/xml', '.pdf': 'application/pdf',
        '.csv': 'text/csv', '.doc': 'application/msword', 
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.xls': 'application/vnd.ms-excel', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.ppt': 'application/vnd.ms-powerpoint', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    }
    return mime_types.get(ext, 'application/octet-stream')


def format_file_size(size_bytes: int) -> str:
    """Format file size in human-readable format."""
    if size_bytes == 0:
        return "0 B"
    
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            if unit == 'B' and size_bytes == int(size_bytes):
                return f"{int(size_bytes)} {unit}"
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    
    return f"{size_bytes:.1f} TB"


def extract_attachment_content(attachment, include_content: bool = True) -> Dict[str, Any]:
    """Extract comprehensive attachment data."""
    try:
        # Basic attachment info
        attachment_data = {
            "name": safe_encode_text(getattr(attachment, "FileName", "Unknown"), "attachment_name"),
            "size": getattr(attachment, "Size", 0),
            "type": getattr(attachment, "Type", 1),  # 1 = olByValue, 5 = olEmbeddeditem
            "position": getattr(attachment, "Position", 0),
            "content_id": getattr(attachment, "PropertyAccessor", None) and 
                         safe_encode_text(attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F"), "content_id") or None,
        }
        
        # Determine MIME type
        attachment_data["mime_type"] = get_mime_type(attachment_data["name"])
        
        # Determine if content should be extracted
        is_embeddable = (attachment_data["mime_type"] in EMBEDDABLE_IMAGE_TYPES or 
                        attachment_data["mime_type"] in TEXT_MIME_TYPES)
        attachment_data["is_embeddable"] = is_embeddable
        
        # Initialize content fields
        attachment_data["content_base64"] = None
        attachment_data["content_size"] = 0
        attachment_data["content_preview"] = None
        
        # Extract content if requested and embeddable
        if include_content and is_embeddable and attachment_data["size"] <= MAX_ATTACHMENT_SIZE:
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=Path(attachment_data["name"]).suffix) as tmp_file:
                    tmp_path = tmp_file.name
                
                attachment.SaveAsFile(tmp_path)
                
                with open(tmp_path, 'rb') as f:
                    content = f.read()
                    attachment_data["content_base64"] = base64.b64encode(content).decode('utf-8')
                    attachment_data["content_size"] = len(content)
                    
                    # For text files, also provide a preview
                    if attachment_data["mime_type"] in TEXT_MIME_TYPES and len(content) <= 1000:
                        try:
                            attachment_data["content_preview"] = content.decode('utf-8', errors='replace')
                        except Exception:
                            attachment_data["content_preview"] = "[Binary content]"
                
                # Clean up temp file
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass
                    
            except Exception as e:
                logger.warning(f"Failed to extract attachment content for {attachment_data['name']}: {e}")
                attachment_data["content_base64"] = None
                attachment_data["content_size"] = 0
        
        return attachment_data
        
    except Exception as e:
        logger.error(f"Error extracting attachment data: {e}")
        return {
            "name": "Unknown",
            "size": 0,
            "type": 1,
            "position": 0,
            "mime_type": "application/octet-stream",
            "is_embeddable": False,
            "content_base64": None,
            "content_size": 0,
            "content_id": None,
            "content_preview": None,
        }


def extract_inline_images_from_body(body_html: str, attachments: List[Dict[str, Any]]) -> Tuple[str, List[Dict[str, Any]]]:
    """Extract inline images from HTML body and replace with embeddable content."""
    try:
        if not body_html or not attachments:
            return body_html, []
        
        inline_images = []
        
        # Pattern to match CID references in HTML
        cid_pattern = r'cid:([^"\'\s>]+)'
        cid_matches = re.findall(cid_pattern, body_html, re.IGNORECASE)
        
        if not cid_matches:
            return body_html, []
        
        # Create a mapping of content IDs to attachments
        modified_body = body_html
        
        for cid in cid_matches:
            # Find attachment with matching content ID
            for attachment in attachments:
                if (attachment.get("content_id") == cid or 
                    attachment.get("name") == cid or
                    Path(attachment.get("name", "")).stem == cid):
                    
                    if attachment.get("content_base64") and attachment.get("is_embeddable"):
                        # Replace CID reference with base64 embedded image
                        mime_type = attachment.get("mime_type", "image/png")
                        base64_content = attachment.get("content_base64")
                        
                        # Create data URI
                        data_uri = f"data:{mime_type};base64,{base64_content}"
                        
                        # Replace all occurrences of this CID
                        modified_body = modified_body.replace(f"cid:{cid}", data_uri)
                        
                        inline_images.append({
                            "content_id": cid,
                            "attachment_name": attachment.get("name"),
                            "mime_type": mime_type,
                            "embedded": True,
                            "size": len(base64_content)
                        })
                        
                        logger.info(f"Replaced inline image CID {cid} with embedded content")
                        break
        
        return modified_body, inline_images
        
    except Exception as e:
        logger.error(f"Error extracting inline images: {e}")
        return body_html, []


def _format_recipient_for_display(recipient) -> str:
    """Format recipient for display using enhanced display name + email format."""
    if isinstance(recipient, dict):
        display_name = recipient.get("display_name", "").strip()
        email = recipient.get("email", "").strip()

        if display_name and email:
            return f"{display_name} <{email}>"
        elif email:
            return email
        elif display_name:
            return display_name
        else:
            return ""

    elif isinstance(recipient, str):
        return recipient.strip()

    return str(recipient) if recipient else ""


def get_email_by_number_unified(
    email_number: int, 
    mode: str = EmailRetrievalMode.BASIC,
    include_attachments: bool = True,
    embed_images: bool = True
) -> Optional[Dict[str, Any]]:
    """Unified email retrieval function with configurable modes.
    
    Args:
        email_number: The email position in cache (1-based)
        mode: Retrieval mode (basic, enhanced, lazy)
        include_attachments: Whether to include attachment content (enhanced mode)
        embed_images: Whether to embed inline images (enhanced mode)
        
    Returns:
        Email data based on requested mode
    """
    if not isinstance(email_number, int) or email_number < 1:
        logger.warning(f"Invalid email number: {email_number}")
        return None


def list_recent_emails(folder_name: str = "Inbox", days: int = None) -> str:
    """Public interface for listing emails (used by CLI).
    Loads emails into cache and returns count message.
    """
    try:
        params = EmailListParams(days=days or 7, folder_name=folder_name)
    except Exception as e:
        logger.error(f"Validation error in list_recent_emails: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    emails, note = get_emails_from_folder(folder_name=params.folder_name, days=params.days)

    days_str = f" from last {params.days} days" if params.days else ""
    logger.info(f"Listed {len(emails)} emails{days_str}")
    return f"Found {len(emails)} emails{days_str}. Use 'view_email_cache_tool' to view them.{note}"


def search_email_by_subject(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Tuple[List[Dict[str, Any]], str]:
    """Search emails by subject and return list of emails with note."""
    return _unified_search(search_term, days, folder_name, match_all, "subject")


def search_email_by_from(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Tuple[List[Dict[str, Any]], str]:
    """Search emails by sender name and return list of emails with note."""
    return _unified_search(search_term, days, folder_name, match_all, "sender")


def search_email_by_to(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Tuple[List[Dict[str, Any]], str]:
    """Search emails by recipient name and return list of emails with note."""
    return _unified_search(search_term, days, folder_name, match_all, "recipient")


def search_email_by_body(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Tuple[List[Dict[str, Any]], str]:
    """Search emails by body content and return list of emails with note."""
    return _unified_search(search_term, days, folder_name, match_all, "body")


def list_folders() -> List[str]:
    """List all available mail folders with recursive subfolder traversal."""
    with OutlookSessionManager() as session:
        folders = []

        def traverse_folders(folder, indent_level=0, parent_path=""):
            """Recursively traverse folders and subfolders.

            Args:
                folder: The folder object to traverse
                indent_level: Current indentation level (0 for top-level folders)
                parent_path: Full path of the parent folder for building complete paths
            """
            # Build the full path for this folder
            if parent_path:
                full_path = f"{parent_path}/{folder.Name}"
            else:
                full_path = folder.Name

            # Add current folder with appropriate indentation and full path info
            if indent_level == 0:
                # For top-level folders, show both name and full path
                folders.append(f"{folder.Name} (Full path: {full_path})")
            else:
                folders.append("  " * indent_level + f"{folder.Name} (Path: {full_path})")

            # Recursively process subfolders
            try:
                # Check if folder has Folders collection
                if folder and hasattr(folder, 'Folders') and folder.Folders:
                    for subfolder in folder.Folders:
                        traverse_folders(subfolder, indent_level + 1, full_path)
                else:
                    logger.debug(f"No subfolders collection available for folder: {getattr(folder, 'Name', 'Unknown')}")
            except Exception as e:
                logger.warning(f"Could not list subfolders for {getattr(folder, 'Name', 'Unknown')}: {e}")

        try:
            # Check if Folders collection exists
            if session.namespace and hasattr(session.namespace, 'Folders') and session.namespace.Folders:
                for root_folder in session.namespace.Folders:
                    traverse_folders(root_folder)
            else:
                logger.warning("No folders collection available in namespace")
        except Exception as e:
            logger.error(f"Error listing folders: {e}")
            raise
        return folders


def get_email_by_number(email_number: int) -> Optional[Dict[str, Any]]:
    """Get detailed information for a specific email by its position in cache (1-based index).
    Implements lazy loading for better performance."""
    if not isinstance(email_number, int) or email_number < 1:
        logger.warning(f"Invalid email number: {email_number}")
        return None

    cache_items = list(email_cache.values())
    if email_number > len(cache_items):
        logger.warning(f"Email number {email_number} out of range (cache size: {len(cache_items)})")
        return None

    email = cache_items[email_number - 1]

    # Validate cache item is a dictionary
    if not isinstance(email, dict):
        logger.error(f"Invalid cache item type: {type(email)}. Expected dict.")
        raise ValueError(f"Invalid cache item type: {type(email)}. Expected dict.")

    # Create filtered copy without sensitive fields
    sender = email.get("sender", "Unknown Sender")
    if isinstance(sender, dict):
        sender_name = sender.get("name", "Unknown Sender")
    else:
        sender_name = str(sender)

    # Check if we have all the details already cached
    if "body" in email and "attachments" in email:
        # Return cached full details
        filtered_email = {
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
        return filtered_email

    # Otherwise, create basic email entry and lazy load details
    filtered_email = {
        "id": email.get("id", ""),
        "subject": email.get("subject", "No Subject"),
        "sender": sender_name,
        "received_time": email.get("received_time", ""),
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
    }

    # Lazy load full details from Outlook only when requested
    try:
        with OutlookSessionManager() as session:
            # Check if namespace is available before using it
            if not session or not session.namespace:
                logger.error("Failed to establish Outlook session or namespace")
                return None
                
            if not hasattr(session.namespace, 'GetItemFromID'):
                logger.error("Namespace does not have GetItemFromID method")
                return None
                
            item = session.namespace.GetItemFromID(email["id"])
            if not item:
                logger.warning(f"Email with ID {email['id'][:20]}... not found")
                return None
                
            if item.Class != OutlookItemClass.MAIL_ITEM:
                logger.warning(f"Email {email_number} is not a mail item")
                return None

            # Extract detailed information
            body = safe_encode_text(getattr(item, "Body", ""), "body")
            attachments = []
            if hasattr(item, "Attachments"):
                try:
                    attachments = [
                        {
                            "name": safe_encode_text(attach.FileName, "attachment_name"),
                            "size": attach.Size,
                        }
                        for attach in item.Attachments
                    ]
                except Exception as e:
                    logger.warning(f"Failed to extract attachments: {e}")

            # Add detailed information
            filtered_email["body"] = body
            filtered_email["attachments"] = attachments

            # Optional: Cache the full details for future requests
            email["body"] = body
            email["attachments"] = attachments
            save_email_cache(force_save=False)  # Save updated cache with full details

            logger.info(f"Retrieved full details for email #{email_number}")
            return filtered_email

    except Exception as e:
        logger.error(f"Error fetching email details for #{email_number}: {e}")
        # Return basic information even if detailed fetch failed
        filtered_email["body"] = "Error loading body content"
        filtered_email["attachments"] = []
        return filtered_email


# Keep alias for backward compatibility
get_email_details = get_email_by_number


def view_email_cache(page: int = 1, per_page: int = 5) -> str:
    """View emails from cache with pagination and detailed info.

    Args:
        page: Page number (1-based)
        per_page: Items per page

    Returns:
        Formatted email previews as string
    """
    if not email_cache:
        return "Error: No emails in cache. Please use list_emails or search_emails first."

    try:
        params = PaginationParams(page=page, per_page=per_page)
    except Exception as e:
        logger.error(f"Validation error in view_email_cache: {e}")
        return f"Error: Invalid pagination parameters: {e}"

    cache_items = list(email_cache.values())
    pagination_info = get_pagination_info(len(cache_items), params.per_page)
    total_pages = pagination_info["total_pages"]
    total_emails = pagination_info["total_items"]

    if params.page > total_pages:
        return f"Error: Page {params.page} does not exist. There are only {total_pages} pages."

    start_idx = (params.page - 1) * params.per_page
    end_idx = min(params.page * params.per_page, total_emails)

    result = f"Showing emails {start_idx + 1}-{end_idx} of {total_emails} (Page {params.page}/{total_pages}):\n\n"

    for i in range(start_idx, end_idx):
        email = cache_items[i]
        result += f"Email #{i + 1}\n"
        result += f"Subject: {email['subject']}\n"
        result += f"From: {email['sender']}\n"

        # Display TO recipients if available
        if email.get("to_recipients"):
            to_names = [_format_recipient_for_display(r) for r in email["to_recipients"]]
            result += f"To: {', '.join(to_names)}\n"

        # Display CC recipients if available
        if email.get("cc_recipients"):
            cc_names = [_format_recipient_for_display(r) for r in email["cc_recipients"]]
            result += f"Cc: {', '.join(cc_names)}\n"

        result += f"Received: {email['received_time']}\n"
        result += f"Read Status: {'Read' if not email.get('unread', False) else 'Unread'}\n"
        result += f"Has Attachments: {'Yes' if email.get('has_attachments', False) else 'No'}\n\n"

    if params.page < total_pages:
        result += f"Use view_email_cache_tool(page={params.page + 1}) to view next page."
    else:
        result += "This is the last page."

    result += "\nCall get_email_details_tool() to get full content of the email."

    return result

    cache_items = list(email_cache.values())
    if email_number > len(cache_items):
        logger.warning(f"Email number {email_number} out of range (cache size: {len(cache_items)})")
        return None

    email = cache_items[email_number - 1]

    # Validate cache item is a dictionary
    if not isinstance(email, dict):
        logger.error(f"Invalid cache item type: {type(email)}. Expected dict.")
        raise ValueError(f"Invalid cache item type: {type(email)}. Expected dict.")

    # For basic mode or if already cached with full details, return cached version
    if mode == EmailRetrievalMode.BASIC or (mode == EmailRetrievalMode.LAZY and "body" in email):
        return _create_basic_email_response(email)

    # For enhanced mode or when lazy loading needs more data
    try:
        return _get_enhanced_email_data(email, include_attachments, embed_images)
    except Exception as e:
        logger.error(f"Failed to get enhanced email data: {e}")
        # Fallback to basic mode
        return _create_basic_email_response(email)


def _create_basic_email_response(email: Dict[str, Any]) -> Dict[str, Any]:
    """Create basic email response from cached data."""
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


def _get_enhanced_email_data(email: Dict[str, Any], include_attachments: bool, embed_images: bool) -> Dict[str, Any]:
    """Get enhanced email data with media support."""
    with OutlookSessionManager() as session:
        if not session or not session.namespace:
            logger.error("Failed to establish Outlook session")
            return _create_basic_email_response(email)
        
        # Get the full email item
        item = session.namespace.GetItemFromID(email["id"])
        if not item or item.Class != OutlookItemClass.MAIL_ITEM:
            logger.warning(f"Email not found or not a mail item")
            return _create_basic_email_response(email)
        
        # Start with basic email data
        enhanced_email = _create_basic_email_response(email)
        
        # Get both plain text and HTML body
        enhanced_email["body"] = safe_encode_text(getattr(item, "Body", ""), "body")
        enhanced_email["html_body"] = safe_encode_text(getattr(item, "HTMLBody", ""), "html_body") if hasattr(item, "HTMLBody") else ""
        enhanced_email["body_format"] = getattr(item, "BodyFormat", 1)  # 1=Plain, 2=HTML, 3=RichText
        
        # Enhanced metadata
        enhanced_email["importance"] = getattr(item, "Importance", 1)  # 0=Low, 1=Normal, 2=High
        enhanced_email["sensitivity"] = getattr(item, "Sensitivity", 0)  # 0=Normal, 1=Personal, 2=Private, 3=Confidential
        enhanced_email["conversation_topic"] = safe_encode_text(getattr(item, "ConversationTopic", ""), "conversation_topic")
        enhanced_email["conversation_id"] = getattr(item, "ConversationID", "")
        enhanced_email["categories"] = getattr(item, "Categories", "")
        enhanced_email["flag_status"] = getattr(item, "FlagStatus", 0)  # 0=Unflagged, 1=Flagged, 2=Complete
        
        # Process attachments if they exist and include_attachments is True
        if enhanced_email["has_attachments"] and include_attachments:
            try:
                attachments = []
                for attachment in item.Attachments:
                    try:
                        attachment_data = extract_attachment_content(attachment, include_attachments)
                        attachments.append(attachment_data)
                    except Exception as e:
                        logger.warning(f"Failed to process attachment: {e}")
                        continue
                
                enhanced_email["attachments"] = attachments
                
                # Process inline images if HTML body exists and embed_images is True
                if embed_images and enhanced_email["html_body"]:
                    try:
                        modified_html, inline_images = extract_inline_images_from_body(
                            enhanced_email["html_body"], attachments
                        )
                        if inline_images:
                            enhanced_email["html_body"] = modified_html
                            enhanced_email["inline_images"] = inline_images
                    except Exception as e:
                        logger.warning(f"Failed to process inline images: {e}")
                        
            except Exception as e:
                logger.error(f"Failed to process attachments: {e}")
                enhanced_email["attachments"] = []
        elif not include_attachments:
            # Clear attachments when include_attachments is False
            enhanced_email["attachments"] = []
        
        return enhanced_email


def format_email_with_media(email_data: Dict[str, Any]) -> str:
    """Format enhanced email data for display with media information."""
    if not email_data:
        return "Error: No email data provided."
    
    result = []
    
    # Basic email info
    result.append(f"Subject: {email_data.get('subject', 'No Subject')}")
    result.append(f"From: {email_data.get('sender', 'Unknown Sender')}")
    result.append(f"To: {email_data.get('to', 'Unknown')}")
    result.append(f"Received: {email_data.get('received_time', 'Unknown')}")
    result.append(f"Size: {email_data.get('size', 0)} bytes")
    
    # Body content
    if email_data.get('body'):
        result.append(f"\nBody:\n{email_data['body']}")
    
    # Attachments
    attachments = email_data.get('attachments', [])
    if attachments:
        result.append(f"\nAttachments ({len(attachments)}):")
        for i, attachment in enumerate(attachments, 1):
            name = attachment.get('name', f'attachment_{i}')
            size = attachment.get('size', 0)
            content_available = 'content' in attachment
            result.append(f"  {i}. {name} ({size} bytes) {'[Content Available]' if content_available else ''}")
    
    # Inline images
    inline_images = email_data.get('inline_images', [])
    if inline_images:
        result.append(f"\nInline Images ({len(inline_images)}):")
        for i, img in enumerate(inline_images, 1):
            name = img.get('name', f'image_{i}')
            size = img.get('size', 0)
            content_available = 'content' in img
            result.append(f"  {i}. {name} ({size} bytes) {'[Embedded]' if content_available else ''}")
    
    return "\n".join(result)


# Import required modules for the missing functions
from datetime import datetime, timedelta, timezone
import time
from .validators import EmailSearchParams, EmailListParams, PaginationParams
from .utils import build_dasl_filter, get_pagination_info


def _unified_search(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True,
    search_field: str = "subject",
) -> Tuple[List[Dict[str, Any]], str]:
    """Unified search function with field-specific filtering.

    Args:
        search_term: Search term to match
        days: Number of days to look back
        folder_name: Optional folder name
        match_all: If True, all terms must match (AND); if False, any term matches (OR)
        search_field: Which field to filter ('subject', 'sender', 'recipient', 'body')

    Returns:
        Tuple of (email list, note string)
    """
    # Validate using Pydantic
    try:
        params = EmailSearchParams(
            search_term=search_term, days=days, folder_name=folder_name, match_all=match_all
        )
    except Exception as e:
        logger.error(f"Validation error in _unified_search: {e}")
        raise ValueError(f"Invalid search parameters: {e}")

    # Map search field to filter parameter
    field_filters = {
        "subject": {"subject_filter_only": True},
        "sender": {"sender_filter_only": True},
        "recipient": {"recipient_filter_only": True},
        "body": {"body_filter_only": True},
    }

    filter_kwargs = field_filters.get(search_field, {"subject_filter_only": True})

    emails, note = get_emails_from_folder(
        search_term=params.search_term,
        days=params.days,
        folder_name=params.folder_name,
        match_all=params.match_all,
        **filter_kwargs,
    )

    # If no results found with server-side filtering, try extended search
    if not emails and search_term:
        search_terms = search_term.lower().split()
        if len(search_terms) == 1:
            extended_days = min(90, days * 4)
            logger.info(f"No results found, trying extended search for {extended_days} days")

            extended_emails, extended_note = get_emails_from_folder(
                search_term=search_term,
                days=extended_days,
                folder_name=folder_name,
                match_all=match_all,
                **filter_kwargs,
            )

            if extended_emails:
                note += f" (extended search in last {extended_days} days)"
                return extended_emails, note

    return emails, note


def get_emails_from_folder(
    search_term: Optional[str] = None,
    days: Optional[int] = None,
    folder_name: Optional[str] = None,
    match_all: bool = True,
    sender_filter_only: bool = False,
    recipient_filter_only: bool = False,
    subject_filter_only: bool = False,
    body_filter_only: bool = False,
) -> Tuple[List[Dict[str, Any]], str]:
    """Retrieve emails from specified folder with batch processing and timeout.

    Args:
        search_term: Optional search term to filter emails
        days: Optional number of days to look back
        folder_name: Optional folder name (defaults to Inbox)
        match_all: If True (default), all search terms must match (AND logic)
        sender_filter_only: If True, only search sender field
        recipient_filter_only: If True, only search recipient field
        subject_filter_only: If True, only search subject field
        body_filter_only: If True, only search body field

    Returns:
        Tuple of (email list, limit note string)
    """
    # Clear cache before loading new emails
    clear_email_cache()

    emails = []
    start_time = time.time()
    retry_count = 0
    limit_note = ""
    failed_count = 0

    with OutlookSessionManager() as session:
        while retry_count < 3:
            try:
                # Check timeout
                if time.time() - start_time > MAX_LOAD_TIME:
                    limit_note = " (MAX_LOAD_TIME reached)"
                    logger.warning(f"MAX_LOAD_TIME reached after processing {len(emails)} emails")
                    break

                folder = session.get_folder(folder_name)
                if not folder:
                    logger.error(f"Folder not found: {folder_name}")
                    return [], " (Folder not found)"

                folder_items = folder.Items

                # Get date threshold
                days_to_use = min(days or MAX_DAYS, MAX_DAYS)
                now = datetime.now()
                threshold_date = now.replace(tzinfo=timezone.utc) - timedelta(days=days_to_use)

                # Build optimized DASL filter if search term is provided
                if search_term:
                    search_terms = search_term.lower().split()

                    # Determine which field to search
                    if sender_filter_only:
                        field_filter = "sender"
                    elif recipient_filter_only:
                        field_filter = "recipient"
                    elif subject_filter_only:
                        field_filter = "subject"
                    elif body_filter_only:
                        field_filter = "body"
                    else:
                        field_filter = "subject"  # default

                    # Build optimized filter using utility function
                    filter_str = build_dasl_filter(search_terms, threshold_date, field_filter, match_all)
                    
                    # Apply filter if we have one
                    if filter_str:
                        logger.info(f"Applying DASL filter: {filter_str}")
                        try:
                            filtered_items = folder_items.Restrict(filter_str)
                            if filtered_items:
                                folder_items = filtered_items
                                logger.info(f"Filter applied, {filtered_items.Count} items match")
                            else:
                                logger.info("No items match the filter")
                                return [], " (No matching emails found)"
                        except Exception as filter_error:
                            logger.warning(f"DASL filter failed, falling back to client-side filtering: {filter_error}")
                            # Continue with unfiltered items and apply client-side filtering below

                # Sort by received time (newest first)
                try:
                    folder_items.Sort("[ReceivedTime]", True)
                except Exception as sort_error:
                    logger.warning(f"Failed to sort items: {sort_error}")

                # Process emails with timeout protection
                processed_count = 0
                
                # Use enumerator for better performance
                try:
                    items_enum = folder_items.GetEnumerator()
                except Exception as enum_error:
                    logger.warning(f"Failed to get enumerator, using direct iteration: {enum_error}")
                    items_enum = folder_items

                for item in items_enum:
                    # Check timeout every 50 emails
                    if processed_count % 50 == 0 and time.time() - start_time > MAX_LOAD_TIME:
                        limit_note = " (MAX_LOAD_TIME reached)"
                        logger.warning(f"MAX_LOAD_TIME reached after processing {processed_count} emails")
                        break

                    # Check email limit
                    if len(emails) >= MAX_EMAILS:
                        limit_note = f" (Limited to {MAX_EMAILS} emails)"
                        logger.info(f"Reached maximum email limit: {MAX_EMAILS}")
                        break

                    try:
                        # Skip non-mail items
                        if item.Class != OutlookItemClass.MAIL_ITEM:
                            continue

                        # Get received time and apply date filtering if no server-side filter was applied
                        received_time = getattr(item, "ReceivedTime", None)
                        if received_time is None:
                            continue

                        # Convert to UTC for comparison
                        try:
                            if isinstance(received_time, datetime):
                                # Ensure timezone-aware
                                if received_time.tzinfo is None:
                                    received_time = received_time.replace(tzinfo=timezone.utc)
                                else:
                                    received_time = received_time.astimezone(timezone.utc)
                            else:
                                # Handle non-datetime received_time
                                received_time = datetime.fromtimestamp(float(received_time), tz=timezone.utc)
                        except (ValueError, TypeError, AttributeError) as dt_error:
                            logger.warning(f"Invalid received time format, skipping: {dt_error}")
                            continue

                        # Apply client-side date filtering if no server-side filter was applied
                        if not search_term and received_time < threshold_date:
                            continue

                        # Client-side search term filtering (if server-side filtering failed or wasn't applied)
                        if search_term and not _client_side_filter(item, search_terms, match_all, sender_filter_only, recipient_filter_only, subject_filter_only, body_filter_only):
                            continue

                        # Extract email data
                        email_data = _extract_email_data(item)
                        if email_data:
                            # Add to cache and results
                            add_email_to_cache(email_data)
                            emails.append(email_data)

                        processed_count += 1

                    except Exception as item_error:
                        failed_count += 1
                        logger.warning(f"Error processing email item: {item_error}")
                        if failed_count > 10:  # Too many failures, stop processing
                            logger.error(f"Too many failed emails ({failed_count}), stopping processing")
                            break
                        continue

                # Force save cache at the end of processing to ensure all emails are saved
                save_email_cache(force_save=True)

                return emails[:MAX_EMAILS], limit_note

            except Exception as e:
                retry_count += 1
                logger.error(f"Error in get_emails_from_folder (attempt {retry_count}/3): {e}")
                if retry_count >= 3:
                    raise RuntimeError(f"Failed after {retry_count} retries: {str(e)}")
                time.sleep(1 * retry_count)  # Simple backoff

    return emails, limit_note


def _client_side_filter(
    item, search_terms: List[str], match_all: bool,
    sender_filter_only: bool, recipient_filter_only: bool,
    subject_filter_only: bool, body_filter_only: bool
) -> bool:
    """Apply client-side filtering to email item.
    
    Args:
        item: Outlook mail item
        search_terms: List of search terms to match
        match_all: If True, all terms must match (AND); if False, any term matches (OR)
        sender_filter_only: If True, only search sender field
        recipient_filter_only: If True, only search recipient field
        subject_filter_only: If True, only search subject field
        body_filter_only: If True, only search body field
        
    Returns:
        True if item matches search criteria, False otherwise
    """
    try:
        # Get field content based on filter type
        if sender_filter_only:
            content = getattr(item, "SenderName", "") + " " + getattr(item, "SenderEmailAddress", "")
        elif recipient_filter_only:
            content = ""
            if hasattr(item, "To"):
                content += getattr(item, "To", "")
            if hasattr(item, "CC"):
                content += " " + getattr(item, "CC", "")
        elif subject_filter_only:
            content = getattr(item, "Subject", "")
        elif body_filter_only:
            content = getattr(item, "Body", "")
        else:
            # Default: search subject
            content = getattr(item, "Subject", "")
            
        content = content.lower()
        
        # Apply search terms
        if match_all:
            # AND logic: all terms must be present
            return all(term in content for term in search_terms)
        else:
            # OR logic: any term must be present
            return any(term in content for term in search_terms)
            
    except Exception as e:
        logger.warning(f"Error in client-side filter: {e}")
        return False


def _extract_email_data(item) -> Optional[Dict[str, Any]]:
    """Extract email data from Outlook mail item.
    
    Args:
        item: Outlook mail item
        
    Returns:
        Dictionary with email data or None if extraction fails
    """
    try:
        # Basic email properties
        email_data = {
            "id": getattr(item, "EntryID", ""),
            "subject": safe_encode_text(getattr(item, "Subject", "No Subject"), "subject"),
            "sender": safe_encode_text(getattr(item, "SenderName", "Unknown Sender"), "sender_name"),
            "received_time": getattr(item, "ReceivedTime", datetime.now()),
            "unread": getattr(item, "UnRead", False),
            "has_attachments": getattr(item, "Attachments", None) and item.Attachments.Count > 0,
            "size": getattr(item, "Size", 0),
            "importance": getattr(item, "Importance", 1),
            "sensitivity": getattr(item, "Sensitivity", 0),
        }
        
        # Convert datetime to ISO format string
        if isinstance(email_data["received_time"], datetime):
            email_data["received_time"] = email_data["received_time"].isoformat()
        
        # Extract recipients
        to_recipients = []
        if hasattr(item, "To") and item.To:
            to_recipients.append({"display_name": item.To, "email": item.To})
        email_data["to_recipients"] = to_recipients
        
        cc_recipients = []
        if hasattr(item, "CC") and item.CC:
            cc_recipients.append({"display_name": item.CC, "email": item.CC})
        email_data["cc_recipients"] = cc_recipients
        
        # Extract categories if available
        if hasattr(item, "Categories") and item.Categories:
            email_data["categories"] = item.Categories
            
        # Extract conversation topic if available
        if hasattr(item, "ConversationTopic") and item.ConversationTopic:
            email_data["conversation_topic"] = item.ConversationTopic
        
        return email_data
        
    except Exception as e:
        logger.error(f"Error extracting email data: {e}")
        return None