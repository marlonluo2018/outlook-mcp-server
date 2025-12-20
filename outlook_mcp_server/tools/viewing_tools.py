"""Email viewing tools for Outlook MCP Server."""

from ..backend.email_data_extractor import get_email_by_number_unified, format_email_with_media
from ..backend.shared import email_cache, email_cache_order, clear_email_cache
from ..backend.outlook_session import OutlookSessionManager


def view_email_cache_tool(page: int = 1) -> dict:
    """View comprehensive information of cached emails (5 emails per page).
    Shows Subject, From, To, CC, Received, Status, and Attachments.

    Args:
        page: Page number to view (1-based, each page contains 5 emails)

    Returns:
        dict: Response containing email previews
        {
            "type": "text",
            "text": "Formatted email previews here"
        }
    """
    if not isinstance(page, int) or page < 1:
        raise ValueError("Page must be a positive integer")
    
    try:
        if not email_cache_order:
            return {"type": "text", "text": "No emails in cache. Please load emails first using list_recent_emails or search functions."}
        
        # Calculate pagination
        start_idx = (page - 1) * 5
        end_idx = start_idx + 5
        
        if start_idx >= len(email_cache_order):
            return {"type": "text", "text": f"Page {page} is out of range. Available range: 1-{(len(email_cache_order) + 4) // 5}"}
        
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
                attachments_display = "Yes" if has_attachments else "No"
                
                # Count embedded images (if available in the email)
                embedded_images_count = 0
                real_attachments_count = len(email_data.get("attachments", []))
                
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
                                            file_name = getattr(attachment, 'FileName', getattr(attachment, 'DisplayName', 'Unknown'))
                                            is_image = file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))
                                            is_pdf = file_name.lower().endswith('.pdf')
                                            
                                            # Check if it's an embedded image
                                            is_embedded = False
                                            try:
                                                if hasattr(attachment, 'PropertyAccessor'):
                                                    content_id = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                                                    is_embedded = content_id is not None and len(str(content_id)) > 0
                                            except:
                                                pass
                                            
                                            # Additional check for image files with common embedded naming pattern
                                            if not is_embedded and is_image and file_name.lower().startswith('image'):
                                                is_embedded = True
                                            
                                            # PDF files are always considered real attachments, not embedded
                                            if is_pdf:
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
                
                # Create attachments info string
                attachments_info = attachments_display
                
                # Store embedded images count directly
                page_emails.append({
                    "number": i + 1,
                    "subject": email_data.get("subject", "No Subject"),
                    "from": sender_name,
                    "to": to_display,
                    "cc": cc_display,
                    "received": email_data.get("received_time", "Unknown"),
                    "status": status,
                    "attachments": attachments_info,
                    "embedded_images_count": embedded_images_count
                })
        
        # Format the output
        result = f"Cached Emails (Page {page} of {(len(email_cache_order) + 4) // 5}):\n"
        result += f"Total emails in cache: {len(email_cache_order)}\n\n"
        
        for email in page_emails:
            result += f"{email['number']}. {email['subject']}\n"
            result += f"   From: {email['from']}\n"
            result += f"   To: {email['to']}\n"
            result += f"   CC: {email['cc']}\n"
            result += f"   Received: {email['received']}\n"
            result += f"   Status: {email['status']}\n"
            embedded_images_display = str(email['embedded_images_count']) if email['embedded_images_count'] > 0 else "None"
            result += f"   Embedded Images: {embedded_images_display}\n"
            result += f"   Attachments: {email['attachments']}\n\n"
        
        return {"type": "text", "text": result}
        
    except Exception as e:
        return {"type": "text", "text": f"Error viewing email cache: {str(e)}"}


def get_email_by_number_tool(email_number: int, mode: str = "basic", include_attachments: bool = True, embed_images: bool = True) -> dict:
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
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    
    if mode not in ["basic", "enhanced", "lazy"]:
        raise ValueError("Mode must be one of: basic, enhanced, lazy")
    
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
        
    except ValueError as e:
        return {"type": "text", "text": f"Invalid input: {str(e)}"}
    except RuntimeError as e:
        return {"type": "text", "text": f"Runtime error: {str(e)}"}
    except Exception as e:
        return {"type": "text", "text": f"Error retrieving email: {str(e)}"}


def load_emails_by_folder_tool(folder_path: str, days: int = 7, max_emails: int = None) -> dict:
    """Load emails from a specific folder into cache.

    Args:
        folder_path: Path to the folder (supports nested paths like "user@company.com/Inbox/SubFolder1")
        days: Number of days to look back (default: 7, max: 30)
        max_emails: Maximum number of emails to load (optional). When specified, loads the most recent emails up to this count.

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
        - Combined: load_emails_by_folder_tool("Inbox", days=7, max_emails=50)
    """
    if not folder_path or not isinstance(folder_path, str):
        raise ValueError("Folder path must be a non-empty string")
    if not isinstance(days, int) or days < 1 or days > 30:
        raise ValueError("Days must be an integer between 1 and 30")
    if max_emails is not None and (not isinstance(max_emails, int) or max_emails < 1):
        raise ValueError("max_emails must be a positive integer when specified")
    
    try:
        # Determine max_emails based on parameters
        if max_emails is not None:
            # Number-based loading: use specified max_emails
            actual_max_emails = min(max_emails, 1000)  # Cap at 1000
        else:
            # Time-based loading: estimate based on days
            # If days is default (7) and no max_emails specified, use reasonable defaults
            if days == 7:
                actual_max_emails = 100  # Reasonable default for 7 days
            else:
                actual_max_emails = min(days * 50, 1000)  # Rough estimate: 50 emails per day, max 1000

        with OutlookSessionManager() as outlook_session:
            email_list, message = outlook_session.get_folder_emails(folder_path, actual_max_emails, fast_mode=True, days_filter=days if max_emails is None else None)
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