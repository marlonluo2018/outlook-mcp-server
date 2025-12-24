"""Email viewing tools for Outlook MCP Server."""

# Type imports
from typing import Any, Dict, List, Optional, Union

# Local application imports
from ..backend.email_data_extractor import format_email_with_media, get_email_by_number_unified
from ..backend.outlook_session import OutlookSessionManager
from ..backend.shared import clear_email_cache, email_cache, email_cache_order
from ..backend.validation import (
    ValidationError,
    validate_cache_available,
    validate_days_parameter,
    validate_email_number,
    validate_folder_name,
    validate_page_parameter
)


def view_email_cache_tool(page: int = 1) -> Dict[str, Any]:
    """View comprehensive information of cached emails (5 emails per page).
    Shows Subject, From, To, CC, Received, Status, and Attachments.

    Args:
        page: Page number to view (1-based, each page contains 5 emails)

    Returns:
        dict: Response containing email previews in JSON format
        {
            "type": "json",
            "data": {
                "page": 1,
                "total_pages": 5,
                "total_emails": 23,
                "emails": [
                    {
                        "number": 1,
                        "subject": "Email Subject",
                        "from": "Sender Name",
                        "to": "Recipient Name",
                        "cc": "CC Recipient",
                        "received": "2023-12-21 10:30:00",
                        "status": "Read",
                        "attachments_count": 2,
                        "embedded_images_count": 2
                    }
                ]
            }
        }
    """
    try:
        validate_page_parameter(page, 0)
    except ValidationError as e:
        raise ValidationError(str(e))
    
    try:
        if not email_cache_order:
            return {
                "type": "json", 
                "data": {
                    "error": "No emails in cache",
                    "message": "Please load emails first using list_recent_emails or search functions."
                }
            }
        
        # Calculate pagination
        start_idx = (page - 1) * 5
        end_idx = start_idx + 5
        total_pages = (len(email_cache_order) + 4) // 5
        
        if start_idx >= len(email_cache_order):
            return {
                "type": "json", 
                "data": {
                    "error": "Page out of range",
                    "message": f"Page {page} is out of range. Available range: 1-{total_pages}"
                }
            }
        
        # Get emails for this page
        page_emails = []
        for i in range(start_idx, min(end_idx, len(email_cache_order))):
            email_id = email_cache_order[i]
            email_data = email_cache.get(email_id, {})
            if email_data:
                # Extract comprehensive information
                sender = email_data.get("sender", "Unknown")
                if isinstance(sender, dict):
                    sender_name = sender.get("name", "Unknown")
                else:
                    sender_name = str(sender)
                
                # Get recipients
                to_recipients = email_data.get("to_recipients", [])
                if to_recipients:
                    to_display = ", ".join([r.get("name", r.get("address", "Unknown")) for r in to_recipients[:3]])
                    if len(to_recipients) > 3:
                        to_display += f" and {len(to_recipients) - 3} more"
                else:
                    to_display = "N/A"
                
                # Get CC recipients
                cc_recipients = email_data.get("cc_recipients", [])
                if cc_recipients:
                    cc_display = ", ".join([r.get("name", r.get("address", "Unknown")) for r in cc_recipients[:3]])
                    if len(cc_recipients) > 3:
                        cc_display += f" and {len(cc_recipients) - 3} more"
                else:
                    cc_display = "N/A"
                
                # Determine status
                unread = email_data.get("unread", False)
                status = "Unread" if unread else "Read"
                
                # Check attachments and embedded images
                has_attachments = email_data.get("has_attachments", False)
                attachments_count = len(email_data.get("attachments", []))
                
                # Use cached embedded_images_count if available, otherwise count manually
                embedded_images_count = email_data.get("embedded_images_count", 0)
                
                # Only re-analyze if we don't have embedded_images_count in cache
                if embedded_images_count == 0 and not email_data.get("attachments_processed", False):
                    try:
                        # Try to get entry_id to check for embedded images
                        entry_id = email_data.get("id", email_data.get("entry_id", ""))
                        if entry_id:
                            from ..backend.outlook_session.session_manager import OutlookSessionManager
                            with OutlookSessionManager() as session:
                                if session and session.namespace and hasattr(session.namespace, 'GetItemFromID'):
                                    try:
                                        item = session.namespace.GetItemFromID(entry_id)
                                        if hasattr(item, 'Attachments') and item.Attachments:
                                            for attachment in item.Attachments:
                                                # Check if it's an embedded image using 4-method detection
                                                is_embedded = False
                                                
                                                # Method 1: Check Content-ID and Content-Location properties
                                                try:
                                                    if hasattr(attachment, 'PropertyAccessor'):
                                                        content_id = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                                                        is_embedded = content_id is not None and len(str(content_id).strip()) > 0
                                                        
                                                        if not is_embedded:
                                                            content_location = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3713001F")
                                                            is_embedded = content_location is not None and len(str(content_location).strip()) > 0
                                                except:
                                                    pass
                                                
                                                # Method 2: Check attachment type
                                                attachment_type = getattr(attachment, 'Type', 1)
                                                if attachment_type in [3, 4]:  # 3 = Embedded, 4 = OLE
                                                    is_embedded = True
                                                
                                                # Method 3: Check for embedded image naming patterns
                                                file_name = getattr(attachment, 'FileName', '') or getattr(attachment, 'DisplayName', 'Unknown')
                                                is_image = file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico'))
                                                if is_image and not is_embedded:
                                                    lower_name = file_name.lower()
                                                    if any(pattern in lower_name for pattern in ['image', 'img', 'cid:', 'embedded']):
                                                        is_embedded = True
                                                    elif '.' in lower_name:
                                                        name_without_ext = lower_name.rsplit('.', 1)[0]
                                                        if name_without_ext.isdigit() or (len(name_without_ext) <= 2 and name_without_ext.isalnum()):
                                                            is_embedded = True
                                                
                                                # Method 4: Check attachment size
                                                if is_image and not is_embedded:
                                                    try:
                                                        attachment_size = getattr(attachment, 'Size', 0)
                                                        if attachment_size > 0 and attachment_size < 10000:  # Less than 10KB
                                                            is_embedded = True
                                                    except:
                                                        pass
                                                
                                                # Document files are always real attachments
                                                is_document = file_name.lower().endswith(('.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.zip', '.rar'))
                                                if is_document:
                                                    is_embedded = False
                                                
                                                # Count embedded images that are not already in the attachments list
                                                if is_embedded:
                                                    # Check if this embedded image is not already counted as a real attachment
                                                    is_real_attachment = False
                                                    for real_attachment in email_data.get("attachments", []):
                                                        if real_attachment.get("name", "") == file_name:
                                                            is_real_attachment = True
                                                            break
                                                    
                                                    if not is_real_attachment:
                                                        embedded_images_count += 1
                                    except:
                                        pass
                    except:
                        pass
                
                # Store embedded images count directly
                page_emails.append({
                    "number": i + 1,
                    "subject": email_data.get("subject", "No Subject"),
                    "from": sender_name,
                    "to": to_display,
                    "cc": cc_display,
                    "received": email_data.get("received_time", "Unknown"),
                    "status": status,
                    "attachments_count": attachments_count,
                    "embedded_images_count": embedded_images_count
                })
        
        # Return JSON format
        return {
            "type": "json",
            "data": {
                "page": page,
                "total_pages": total_pages,
                "total_emails": len(email_cache_order),
                "emails": page_emails
            }
        }
        
    except Exception as e:
        return {
            "type": "json", 
            "data": {
                "error": "Error viewing email cache",
                "message": str(e)
            }
        }


def get_email_by_number_tool(email_number: int, mode: str = "basic", include_attachments: bool = True, embed_images: bool = True) -> Dict[str, Any]:
    """Get email content by cache number with 3 retrieval modes.
    
    Mode Selection Guide:
    - "basic": Full text content without embedded images and attachments - use for text-focused viewing
    - "enhanced": Full content + complete thread + HTML + attachments - use for complete analysis  
    - "lazy": Auto-adapts cached vs live data - use when unsure
    
    Email Thread Handling:
    - "basic": No conversation threads (focus on individual email content)
    - "enhanced": Shows complete conversation thread
    - "lazy": Auto-adaptive thread handling
    
    Requires emails to be loaded first via list_recent_emails or search_emails.
    
    Args:
        email_number: Position in cache (1-based)
        mode: "basic" (text-only), "enhanced" (complete), "lazy" (adaptive)
        include_attachments: Include file content (enhanced mode only)
        embed_images: Embed inline images as data URIs (enhanced mode only)

    Returns:
        dict: Response containing email details based on requested mode
        {
            "type": "text", 
            "text": "Formatted email content"
        }

    Raises:
        ValueError: If email number is invalid or no emails are loaded
        RuntimeError: If cache contains invalid data
    """
    try:
        validate_email_number(email_number, len(email_cache_order))
    except ValidationError as e:
        raise ValidationError(str(e))
    
    if mode not in ["basic", "enhanced", "lazy"]:
        raise ValidationError("Mode must be one of: basic, enhanced, lazy")
    
    try:
        email_data = get_email_by_number_unified(
            email_number, 
            mode=mode, 
            include_attachments=include_attachments, 
            embed_images=embed_images
        )
        
        if email_data is None:
            return {
                "type": "text", 
                "text": f"No email found at position {email_number}. Please load emails first using list_recent_emails or search_emails."
            }
        
        # Format the email with media
        formatted_email = format_email_with_media(email_data)
        return {"type": "text", "text": formatted_email}
        
    except ValidationError as e:
        return {"type": "text", "text": f"Validation error: {str(e)}"}
    except RuntimeError as e:
        return {"type": "text", "text": f"Runtime error: {str(e)}"}
    except Exception as e:
        return {"type": "text", "text": f"Error retrieving email: {str(e)}"}


def load_emails_by_folder_tool(folder_path: str, days: int = None, max_emails: int = None) -> Dict[str, Any]:
    """Load emails from a specific folder into cache.

    **LLM Note**: This function enforces strict mutual exclusion between 'days' and 'max_emails' parameters.
    You CANNOT use both parameters together. Choose either time-based loading (days) or number-based loading (max_emails).
    Attempting to use both parameters will raise a ValidationError.

    Args:
        folder_path: Path to the folder (supports nested paths like "user@company.com/Inbox/SubFolder1")
        days: Number of days to look back (max: 30) - mutually exclusive with max_emails
        max_emails: Maximum number of emails to load (mutually exclusive with days) - when specified, loads the most recent emails up to this count

    Returns:
        dict: Response containing email count message

    Note:
        Maximum 30 emails for Inbox folder
        Maximum 1000 emails for other folders
        Supports nested folder paths (e.g., "user@company.com/Inbox/SubFolder1/SubFolder2")
        
        IMPORTANT: Folder paths must include the email address as the root folder.
        Use format: "user@company.com/Inbox/SubFolder" not just "Inbox/SubFolder"
        
        Usage examples:
        - Time-based: load_emails_by_folder_tool("Inbox", days=7)
        - Number-based: load_emails_by_folder_tool("Inbox", max_emails=50)
        - Cannot use both: load_emails_by_folder_tool("Inbox", days=7, max_emails=50) - this will raise an error
    """
    try:
        validate_folder_name(folder_path)
    except ValidationError as e:
        return {"type": "text", "text": f"Validation error: {str(e)}"}
    
    validated_folder = folder_path
    
    # Enforce mutual exclusion: cannot use both days and max_emails together
    if days is not None and max_emails is not None:
        return {"type": "text", "text": "Cannot specify both 'days' and 'max_emails' parameters. Use either time-based (days) or number-based (max_emails) loading, not both."}
    
    # Set default behavior if neither parameter is specified
    if days is None and max_emails is None:
        days = 7  # Default to 7 days if neither parameter is specified
    
    # Validate parameters
    try:
        if days is not None:
            validate_days_parameter(days)
        
        if max_emails is not None and (not isinstance(max_emails, int) or max_emails < 1):
            raise ValidationError("max_emails must be a positive integer when specified")
    except ValidationError as e:
        return {"type": "text", "text": f"Validation error: {str(e)}"}
    
    try:
        # Determine max_emails based on parameters
        if max_emails is not None:
            # Number-based loading: use specified max_emails
            actual_max_emails = min(max_emails, 1000)  # Cap at 1000
        else:
            # Time-based loading: use a very high limit to get all emails in the date range
            # This ensures we respect the days parameter strictly without expanding the date range
            actual_max_emails = 10000  # High limit to capture all emails in the specified date range

        with OutlookSessionManager() as outlook_session:
            email_list, message = outlook_session.get_folder_emails(validated_folder, actual_max_emails, fast_mode=True, days_filter=days if max_emails is None else None)
            return {"type": "text", "text": message + ". Use 'view_email_cache_tool' to view them."}
    except Exception as e:
        return {"type": "text", "text": f"Error loading emails from folder: {str(e)}"}


def clear_email_cache_tool() -> dict:
    """Clear the email cache both in memory and on disk.

    This tool removes all cached emails from memory and deletes the persistent
    cache file from disk. Use this when you want to free up memory or ensure
    fresh data is loaded from Outlook.

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email cache cleared successfully"
        }
    """
    try:
        # Get current cache size for confirmation
        cache_size = len(email_cache_order)
        
        # Clear the cache
        clear_email_cache()
        
        return {
            "type": "text", 
            "text": f"Email cache cleared successfully. Removed {cache_size} cached emails."
        }
        
    except Exception as e:
        return {
            "type": "text", 
            "text": f"Error clearing email cache: {str(e)}"
        }