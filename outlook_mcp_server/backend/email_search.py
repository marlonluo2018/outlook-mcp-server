"""Email search and folder management functionality."""

import logging
from datetime import datetime
from typing import Dict, Any, Optional, List, Tuple

from .outlook_session import OutlookSessionManager
from .shared import email_cache, add_email_to_cache, email_cache_order
from .utils import OutlookItemClass, build_dasl_filter, get_pagination_info, safe_encode_text
from .validators import EmailSearchParams, EmailListParams, PaginationParams

logger = logging.getLogger(__name__)


def list_recent_emails(folder_name: str = "Inbox", days: int = None) -> str:
    """Public interface for listing emails (used by CLI).
    Loads emails into cache and returns count message.
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
    from .shared import clear_email_cache, email_cache, email_cache_order
    logger.info(f"Cache before clearing: {len(email_cache)} emails, {len(email_cache_order)} order")
    clear_email_cache()
    logger.info(f"Cache after clearing: {len(email_cache)} emails, {len(email_cache_order)} order")

    emails, note = get_emails_from_folder(folder_name=params.folder_name, days=params.days)
    
    logger.info(f"get_emails_from_folder returned {len(emails)} emails, note: {note}")

    # UVX COMPATIBILITY: Ensure cache is saved immediately after loading emails
    # Only save to cache if we actually got emails successfully (no error)
    if emails and "Error:" not in note:
        from .shared import immediate_save_cache
        immediate_save_cache()

    days_str = f" from last {params.days} days" if params.days else ""
    logger.info(f"Listed {len(emails)} emails{days_str}")
    
    # Always return the note from get_emails_from_folder to ensure consistency
    # The note already contains the correct email count or error message
    return note


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
    return _optimized_body_search(search_term, days, folder_name, match_all)


def list_folders() -> List[str]:
    """List all available mail folders with recursive subfolder traversal."""
    with OutlookSessionManager() as session:
        folders = []

        def traverse_folders(folder, indent_level=0, parent_path=""):
            """Recursively traverse folders and subfolders."""
            # Build the full path for this folder
            if parent_path:
                full_path = f"{parent_path}/{folder.Name}"
            else:
                full_path = folder.Name

            # Add current folder with appropriate indentation and full path info
            if indent_level == 0:
                folders.append(full_path)
            else:
                folders.append(f"{'  ' * indent_level}{full_path}")

            # Traverse subfolders
            try:
                if hasattr(folder, 'Folders') and folder.Folders.Count > 0:
                    for subfolder in folder.Folders:
                        traverse_folders(subfolder, indent_level + 1, full_path)
            except Exception as e:
                logger.warning(f"Error traversing subfolders of {folder.Name}: {e}")

        try:
            # Start with the root folders
            for folder in session.namespace.Folders:
                traverse_folders(folder)
        except Exception as e:
            logger.error(f"Error listing folders: {e}")
            return ["Error: Unable to list folders"]

    return folders


def view_email_cache(page: int = 1, per_page: int = 5) -> str:
    """View emails from cache with pagination and detailed info."""
    if not email_cache:
        return "Error: No emails in cache. Please use list_emails or search_emails first."

    try:
        params = PaginationParams(page=page, per_page=per_page)
    except Exception as e:
        logger.error(f"Validation error in view_email_cache: {e}")
        return f"Error: Invalid pagination parameters: {e}"

    # Use sorted order from email_cache_order to maintain chronological sorting
    total_emails = len(email_cache_order)
    
    if total_emails == 0:
        return "No emails in cache."
    
    # Calculate pagination
    start_idx = (params.page - 1) * params.per_page
    end_idx = start_idx + params.per_page
    
    if start_idx >= total_emails:
        return f"Page {params.page} is out of range. Total emails: {total_emails}"
    
    # Get emails for current page
    page_emails = []
    for i in range(start_idx, min(end_idx, total_emails)):
        email_id = email_cache_order[i]
        email = email_cache[email_id]
        page_emails.append(email)
    
    # Format output
    output = [f"Email Cache (Page {params.page}/{((total_emails - 1) // params.per_page) + 1}, Total: {total_emails})"]
    output.append("=" * 80)
    
    for i, email in enumerate(page_emails, start_idx + 1):
        sender = email.get("sender", "Unknown Sender")
        if isinstance(sender, dict):
            sender_name = sender.get("name", "Unknown Sender")
            sender_email = sender.get("email", "")
            sender_str = f"{sender_name} <{sender_email}>" if sender_email else sender_name
        else:
            sender_str = str(sender)
        
        subject = email.get("subject", "No Subject")
        received = email.get("received_time", "Unknown Date")
        unread = email.get("unread", False)
        has_attachments = email.get("has_attachments", False)
        
        # Truncate long subjects
        if len(subject) > 50:
            subject = subject[:47] + "..."
        
        status = "UNREAD" if unread else "READ"
        attachment_indicator = "📎" if has_attachments else ""
        
        output.append(f"{i:3d}. {status} | {received} | {sender_str}")
        output.append(f"     Subject: {subject} {attachment_indicator}")
        output.append("")
    
    return "\n".join(output)


def _unified_search(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True,
    search_field: str = "subject",
) -> Tuple[List[Dict[str, Any]], str]:
    """Unified search function for all email fields."""
    try:
        params = EmailSearchParams(
            search_term=search_term,
            days=days,
            folder_name=folder_name,
            match_all=match_all,
        )
    except Exception as e:
        logger.error(f"Validation error in unified search: {e}")
        raise ValueError(f"Invalid search parameters: {e}")

    logger.info(f"Searching for '{search_term}' in {search_field} (days={days}, folder={folder_name})")
    
    # Use optimized server-side filtering for subject, sender, and recipient
    if search_field in ["subject", "sender", "recipient"]:
        return _server_side_search(params.search_term, params.days, params.folder_name, params.match_all, search_field)
    
    # Fall back to client-side filtering for body search
    return _client_side_search(params.search_term, params.days, params.folder_name, params.match_all, search_field)


def _server_side_search(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True,
    search_field: str = "subject",
) -> Tuple[List[Dict[str, Any]], str]:
    """Optimized server-side search using Outlook's DASL filters."""
    from datetime import datetime, timedelta
    from .utils import build_dasl_filter, sanitize_search_term
    
    # Sanitize search term
    sanitized_term = sanitize_search_term(search_term)
    if not sanitized_term:
        return [], "Error: Invalid search term"
    
    # Build search terms list
    search_terms = [term.strip() for term in sanitized_term.split() if term.strip()]
    if not search_terms:
        return [], "Error: No valid search terms"
    
    # Calculate threshold date
    threshold_date = datetime.now() - timedelta(days=days)
    
    with OutlookSessionManager() as session:
        if not session or not session.namespace:
            logger.error("Failed to establish Outlook session")
            return [], "Error: Failed to establish Outlook session"
        
        try:
            # Find the target folder
            target_folder = None
            folder_parts = folder_name.split("/") if folder_name else ["Inbox"]
            
            # Try to get standard folders directly
            if len(folder_parts) == 1 and folder_parts[0] in ["Inbox", "Sent Items", "Drafts", "Outbox"]:
                try:
                    if folder_parts[0] == "Inbox":
                        target_folder = session.namespace.GetDefaultFolder(6)
                    elif folder_parts[0] == "Sent Items":
                        target_folder = session.namespace.GetDefaultFolder(5)
                    elif folder_parts[0] == "Drafts":
                        target_folder = session.namespace.GetDefaultFolder(16)
                    elif folder_parts[0] == "Outbox":
                        target_folder = session.namespace.GetDefaultFolder(4)
                except Exception as e:
                    logger.warning(f"Could not get default folder {folder_parts[0]}: {e}")
            
            if not target_folder:
                # Navigate through folder hierarchy
                target_folder = session.namespace
                for part in folder_parts:
                    found = False
                    for folder in target_folder.Folders:
                        if folder.Name.lower() == part.lower():
                            target_folder = folder
                            found = True
                            break
                    if not found:
                        logger.error(f"Folder '{part}' not found in path '{folder_name}'")
                        return [], f"Error: Folder '{folder_name}' not found"
            
            # Build DASL filter for server-side filtering
            dasl_filter = build_dasl_filter(search_terms, threshold_date, search_field, match_all)
            logger.info(f"Using DASL filter: {dasl_filter}")
            
            # Apply server-side filter
            try:
                items = target_folder.Items.Restrict(dasl_filter)
                logger.info(f"DASL filter applied, got {items.Count} items")
            except Exception as filter_error:
                logger.warning(f"DASL filter failed: {filter_error}. Falling back to client-side filtering.")
                # Fall back to client-side search if server-side filtering fails
                return _client_side_search(search_term, days, folder_name, match_all, search_field)
            
            # Sort by received time (newest first)
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception as e:
                logger.warning(f"Could not sort items: {e}")
            
            # Process filtered items
            emails = []
            count = 0
            max_emails = 1000
            
            for item in items:
                try:
                    if count >= max_emails:
                        break
                        
                    if item.Class == OutlookItemClass.MAIL_ITEM:
                        email_data = _extract_email_data(item)
                        if email_data:
                            emails.append(email_data)
                            count += 1
                            
                except Exception as e:
                    logger.warning(f"Error processing email item: {e}")
                    continue
            
            # Add to cache
            if emails:
                from .shared import add_email_to_cache
                logger.debug(f"Adding {len(emails)} emails to cache...")
                for email_data in emails:
                    add_email_to_cache(email_data["id"], email_data)
            
            logger.info(f"Server-side search found {len(emails)} matching emails")
            return emails, f"Found {len(emails)} matching emails from last {days} days"
            
        except Exception as e:
            logger.error(f"Error in server-side search: {e}")
            return [], f"Error: {e}"


def _client_side_search(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True,
    search_field: str = "subject",
) -> Tuple[List[Dict[str, Any]], str]:
    """Client-side search for cases where server-side filtering is not available."""
    logger.info(f"Using client-side search for '{search_term}' in {search_field}")
    
    # Get emails from folder
    emails, note = get_emails_from_folder(folder_name=folder_name, days=days)
    
    if not emails:
        return [], note
    
    # Filter emails based on search criteria
    filtered_emails = []
    for email in emails:
        if _client_side_filter(email, search_term, match_all, search_field):
            filtered_emails.append(email)
    
    # Sort by received time (newest first)
    filtered_emails.sort(key=lambda x: x.get("received_time", ""), reverse=True)
    
    logger.info(f"Client-side search found {len(filtered_emails)} matching emails")
    
    return filtered_emails, f"Found {len(filtered_emails)} matching emails from last {days} days"


def _optimized_body_search(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True,
) -> Tuple[List[Dict[str, Any]], str]:
    """Optimized body search that only loads email bodies when necessary."""
    from datetime import datetime, timedelta
    
    logger.info(f"Starting optimized body search for '{search_term}'")
    
    # Build search terms
    search_terms = [term.strip().lower() for term in search_term.split() if term.strip()]
    if not search_terms:
        return [], "Error: No valid search terms"
    
    with OutlookSessionManager() as session:
        if not session or not session.namespace:
            logger.error("Failed to establish Outlook session")
            return [], "Error: Failed to establish Outlook session"
        
        try:
            # Find the target folder
            target_folder = None
            folder_parts = folder_name.split("/") if folder_name else ["Inbox"]
            
            # Try to get standard folders directly
            if len(folder_parts) == 1 and folder_parts[0] in ["Inbox", "Sent Items", "Drafts", "Outbox"]:
                try:
                    if folder_parts[0] == "Inbox":
                        target_folder = session.namespace.GetDefaultFolder(6)
                    elif folder_parts[0] == "Sent Items":
                        target_folder = session.namespace.GetDefaultFolder(5)
                    elif folder_parts[0] == "Drafts":
                        target_folder = session.namespace.GetDefaultFolder(16)
                    elif folder_parts[0] == "Outbox":
                        target_folder = session.namespace.GetDefaultFolder(4)
                except Exception as e:
                    logger.warning(f"Could not get default folder {folder_parts[0]}: {e}")
            
            if not target_folder:
                # Navigate through folder hierarchy
                target_folder = session.namespace
                for part in folder_parts:
                    found = False
                    for folder in target_folder.Folders:
                        if folder.Name.lower() == part.lower():
                            target_folder = folder
                            found = True
                            break
                    if not found:
                        logger.error(f"Folder '{part}' not found in path '{folder_name}'")
                        return [], f"Error: Folder '{folder_name}' not found"
            
            # Build date filter
            if days and days > 0:
                start_date = datetime.now() - timedelta(days=days)
                filter_str = f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y')}'"
                logger.info(f"Using date filter: {filter_str}")
            else:
                filter_str = ""
            
            # Get items with date filter first (much faster than loading all)
            if filter_str:
                try:
                    items = target_folder.Items.Restrict(filter_str)
                    logger.info(f"Date filter applied, got {items.Count} items")
                except Exception as filter_error:
                    logger.warning(f"Date filter failed: {filter_error}. Using all items.")
                    items = target_folder.Items
            else:
                items = target_folder.Items
            
            # Sort by received time (newest first)
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception as e:
                logger.warning(f"Could not sort items: {e}")
            
            # Process items with lazy body loading
            emails = []
            count = 0
            max_emails = 1000
            
            for item in items:
                try:
                    if count >= max_emails:
                        break
                        
                    if item.Class == OutlookItemClass.MAIL_ITEM:
                        # First, extract basic email data without body
                        email_data = _extract_email_data_light(item)
                        if email_data:
                            # Only load and check body if basic data passes initial criteria
                            if _should_check_body(email_data, search_terms, match_all):
                                body_content = _get_email_body(item)
                                if body_content and _body_matches_search(body_content, search_terms, match_all):
                                    # Add body to email data and include in results
                                    email_data["body"] = body_content
                                    emails.append(email_data)
                                    count += 1
                            else:
                                # Skip body loading for emails that don't match basic criteria
                                logger.debug(f"Skipping body check for email: {email_data.get('subject', 'No Subject')}")
                                
                except Exception as e:
                    logger.warning(f"Error processing email item: {e}")
                    continue
            
            # Add to cache
            if emails:
                from .shared import add_email_to_cache
                logger.debug(f"Adding {len(emails)} emails to cache...")
                for email_data in emails:
                    add_email_to_cache(email_data["id"], email_data)
            
            logger.info(f"Body search found {len(emails)} matching emails")
            return emails, f"Found {len(emails)} matching emails from last {days} days"
            
        except Exception as e:
            logger.error(f"Error in optimized body search: {e}")
            return [], f"Error: {e}"


def _extract_email_data_light(item) -> Optional[Dict[str, Any]]:
    """Extract basic email data without loading the body (for performance)."""
    try:
        # Basic email properties without body
        received_time = getattr(item, "ReceivedTime", "")
        if received_time:
            received_time_str = str(received_time)
            if '+' in received_time_str:
                received_time_str = received_time_str.split('+')[0].strip()
        else:
            received_time_str = ""
        
        email_data = {
            "id": getattr(item, "EntryID", ""),
            "subject": safe_encode_text(getattr(item, "Subject", "No Subject"), "subject"),
            "received_time": received_time_str,
            "unread": getattr(item, "UnRead", False),
            "has_attachments": getattr(item, "Attachments", None) is not None and getattr(item, "Attachments").Count > 0,
            "size": getattr(item, "Size", 0),
        }
        
        # Sender information (lightweight)
        try:
            sender = getattr(item, "Sender", None)
            if sender:
                sender_name = getattr(sender, "Name", "")
                sender_email = getattr(sender, "Address", "")
                email_data["sender"] = {
                    "name": safe_encode_text(sender_name, "sender_name"),
                    "email": safe_encode_text(sender_email, "sender_email")
                }
            else:
                email_data["sender"] = {"name": "Unknown Sender", "email": ""}
        except Exception as e:
            logger.warning(f"Error extracting sender info: {e}")
            email_data["sender"] = {"name": "Unknown Sender", "email": ""}
        
        # Recipients (lightweight)
        try:
            to_recipients = getattr(item, "To", "")
            cc_recipients = getattr(item, "CC", "")
            email_data["to_recipients"] = _parse_recipients_light(to_recipients)
            email_data["cc_recipients"] = _parse_recipients_light(cc_recipients)
        except Exception as e:
            logger.warning(f"Error extracting recipients: {e}")
            email_data["to_recipients"] = []
            email_data["cc_recipients"] = []
        
        return email_data
        
    except Exception as e:
        logger.warning(f"Error extracting light email data: {e}")
        return None


def _parse_recipients_light(recipients_str: str) -> List[Dict[str, str]]:
    """Parse recipient string into lightweight format."""
    recipients = []
    if not recipients_str:
        return recipients
    
    # Simple split by semicolon
    for recipient in recipients_str.split(";"):
        recipient = recipient.strip()
        if recipient:
            # Extract name and email if in format "Name <email>"
            if "<" in recipient and ">" in recipient:
                name = recipient.split("<")[0].strip()
                email = recipient.split("<")[1].split(">")[0].strip()
                recipients.append({"name": name, "email": email})
            else:
                # Assume it's just an email
                recipients.append({"name": "", "email": recipient})
    
    return recipients


def _should_check_body(email_data: Dict[str, Any], search_terms: List[str], match_all: bool) -> bool:
    """Quick pre-filter to determine if we should load the body."""
    # For now, always check body - but this could be enhanced with heuristics
    # like checking if subject contains similar terms, etc.
    return True


def _get_email_body(item) -> str:
    """Get email body content."""
    try:
        body = getattr(item, "Body", "")
        return safe_encode_text(body, "body")
    except Exception as e:
        logger.warning(f"Error getting email body: {e}")
        return ""


def _body_matches_search(body_content: str, search_terms: List[str], match_all: bool) -> bool:
    """Check if body content matches search terms."""
    body_lower = body_content.lower()
    
    if match_all:
        return all(term in body_lower for term in search_terms)
    else:
        return any(term in body_lower for term in search_terms)


def get_emails_from_folder(folder_name: str = "Inbox", days: int = 7) -> Tuple[List[Dict[str, Any]], str]:
    """Get emails from specified folder within time range."""
    with OutlookSessionManager() as session:
        if not session or not session.namespace:
            logger.error("Failed to establish Outlook session")
            return [], "Error: Failed to establish Outlook session"
        
        try:
            # Find the target folder
            target_folder = None
            folder_parts = folder_name.split("/") if folder_name else ["Inbox"]
            
            # Start from root or namespace
            if len(folder_parts) == 1 and folder_parts[0] in ["Inbox", "Sent Items", "Drafts", "Outbox"]:
                # Try to get standard folders directly
                try:
                    if folder_parts[0] == "Inbox":
                        target_folder = session.namespace.GetDefaultFolder(6)  # olFolderInbox
                    elif folder_parts[0] == "Sent Items":
                        target_folder = session.namespace.GetDefaultFolder(5)  # olFolderSentMail
                    elif folder_parts[0] == "Drafts":
                        target_folder = session.namespace.GetDefaultFolder(16)  # olFolderDrafts
                    elif folder_parts[0] == "Outbox":
                        target_folder = session.namespace.GetDefaultFolder(4)  # olFolderOutbox
                except Exception as e:
                    logger.warning(f"Could not get default folder {folder_parts[0]}: {e}")
            
            if not target_folder:
                # Navigate through folder hierarchy
                target_folder = session.namespace
                for part in folder_parts:
                    found = False
                    for folder in target_folder.Folders:
                        if folder.Name.lower() == part.lower():
                            target_folder = folder
                            found = True
                            break
                    if not found:
                        logger.error(f"Folder '{part}' not found in path '{folder_name}'")
                        return [], f"Error: Folder '{folder_name}' not found"
            
            # Build date filter with format that works with Outlook
            if days and days > 0:
                from datetime import datetime, timedelta
                start_date = datetime.now() - timedelta(days=days)
                # Use the format that works with Outlook: mm/dd/yyyy
                filter_str = f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y')}'"
                logger.info(f"Using filter: {filter_str}")
            else:
                filter_str = ""
            
            # Get items from folder with fallback for date filter
            if filter_str:
                try:
                    items = target_folder.Items.Restrict(filter_str)
                    logger.info(f"Filter '{filter_str}' applied, got {items.Count} items")
                except Exception as filter_error:
                    logger.warning(f"Filter failed: {filter_error}. Trying without filter.")
                    items = target_folder.Items
                    # If filter failed, we'll need to apply date filtering manually
                    if days and days > 0:
                        from datetime import datetime, timedelta
                        start_date = datetime.now() - timedelta(days=days)
                        logger.info(f"Will apply manual date filtering for items since {start_date}")
            else:
                items = target_folder.Items
            
            # Sort by received time (newest first)
            try:
                items.Sort("[ReceivedTime]", True)
                logger.info(f"Sorted {items.Count} items by ReceivedTime")
            except Exception as e:
                logger.warning(f"Could not sort items: {e}")
            
            # Process items with manual date filtering if needed
            emails = []
            count = 0
            max_emails = 1000  # Limit to prevent memory issues
            
            # Determine if we need manual date filtering
            needs_manual_filter = False
            all_items = None  # Store reference to all items
            
            if days and days > 0 and filter_str and items.Count == 0:
                needs_manual_filter = True
                # Get all items without filter
                all_items = target_folder.Items
                logger.info(f"Filter returned 0 items, applying manual date filtering for {all_items.Count} items")
            else:
                all_items = items
            
            # Pre-calculate start date if manual filtering is needed
            start_date = None
            if needs_manual_filter:
                from datetime import datetime, timedelta
                start_date = datetime.now() - timedelta(days=days)
                logger.info(f"Will manually filter items since {start_date}")
            
            # Use simple for loop with COM collection
            # Note: Using a for loop with COM collections can be more reliable
            for item in all_items:
                try:
                    if count >= max_emails:
                        logger.info(f"Reached maximum email limit of {max_emails}")
                        break
                        
                    if item.Class == OutlookItemClass.MAIL_ITEM:
                        # Manual date filtering if needed (pre-calculated start_date)
                        if needs_manual_filter and start_date:
                            received_time = getattr(item, "ReceivedTime", None)
                            if received_time:
                                try:
                                    # Handle different COM datetime formats
                                    received_str = str(received_time)
                                    if 'T' in received_str and '+' in received_str:
                                        # Format: 2025-12-17 23:31:02.980000+00:00
                                        # Remove timezone info for comparison
                                        received_str = received_str.split('+')[0].strip()
                                        received_dt = datetime.strptime(received_str, "%Y-%m-%d %H:%M:%S.%f")
                                    else:
                                        # Try legacy format
                                        received_dt = datetime.strptime(received_str, "%m/%d/%y %H:%M:%S")
                                    
                                    if received_dt < start_date:
                                        continue  # Skip this item as it's too old
                                except Exception as date_error:
                                    logger.warning(f"Error parsing date {received_time}: {date_error}")
                                    # If we can't parse the date, include it to be safe
                                    continue
                        
                        email_data = _extract_email_data(item)
                        if email_data:
                            emails.append(email_data)
                            count += 1
                            
                            # Don't add to cache immediately - batch it for performance
                            # This will be done after all emails are collected
                            
                except Exception as e:
                    logger.warning(f"Error processing email item: {e}")
                    continue
            
            logger.info(f"Retrieved {len(emails)} emails from folder '{folder_name}' (days={days})")
            
            # Only add emails to cache if processing was successful
            # This ensures we don't add emails if there were processing errors
            if emails:
                from .shared import add_email_to_cache
                logger.debug(f"Batch adding {len(emails)} emails to cache...")
                for email_data in emails:
                    add_email_to_cache(email_data["id"], email_data)
                logger.debug(f"Finished adding emails to cache")
            
            result = (emails, f"Found {len(emails)} emails from last {days} days" if days else f"Found {len(emails)} emails")
            return result
            
        except Exception as e:
            logger.error(f"Error getting emails from folder '{folder_name}': {e}")
            # If an error occurred, return empty result with error message
            # Do NOT add any emails to cache in case of errors
            return [], f"Error: Failed to get emails from folder '{folder_name}': {e}"


def _client_side_filter(
    email: Dict[str, Any], 
    search_term: str, 
    match_all: bool = True, 
    search_field: str = "subject"
) -> bool:
    """Client-side filtering for email search."""
    if not search_term.strip():
        return True
    
    search_terms = [term.strip().lower() for term in search_term.split() if term.strip()]
    if not search_terms:
        return True
    
    # Get the field to search in
    if search_field == "subject":
        content = email.get("subject", "").lower()
    elif search_field == "sender":
        sender = email.get("sender", {})
        if isinstance(sender, dict):
            content = f"{sender.get('name', '')} {sender.get('email', '')}".lower()
        else:
            content = str(sender).lower()
    elif search_field == "recipient":
        to_recipients = email.get("to_recipients", [])
        cc_recipients = email.get("cc_recipients", [])
        all_recipients = to_recipients + cc_recipients
        content = " ".join([
            f"{r.get('name', '')} {r.get('email', '')}" if isinstance(r, dict) else str(r)
            for r in all_recipients
        ]).lower()
    elif search_field == "body":
        # For body search, we need to load the full content
        content = email.get("body", "").lower()
        # Also check cached body content if available
        if not content and "body" in email:
            content = email["body"].lower()
    else:
        content = ""
    
    # Check if all or any terms match
    if match_all:
        return all(term in content for term in search_terms)
    else:
        return any(term in content for term in search_terms)


def _extract_email_data(item) -> Optional[Dict[str, Any]]:
    """Extract email data from Outlook item."""
    try:
        # Basic email properties
        received_time = getattr(item, "ReceivedTime", "")
        # Format the received_time for consistent handling (optimized)
        if received_time:
            # Direct string conversion without unnecessary processing
            received_time_str = str(received_time)
            # Only process if it has timezone info
            if '+' in received_time_str:
                # Split off timezone info only if needed
                received_time_str = received_time_str.split('+')[0].strip()
        else:
            received_time_str = ""
        
        email_data = {
            "id": getattr(item, "EntryID", ""),
            "subject": safe_encode_text(getattr(item, "Subject", "No Subject"), "subject"),
            "received_time": received_time_str,
            "unread": getattr(item, "UnRead", False),
            "has_attachments": getattr(item, "Attachments", None) is not None and getattr(item, "Attachments").Count > 0,
            "size": getattr(item, "Size", 0),
        }
        
        # Sender information
        try:
            sender = getattr(item, "Sender", None)
            if sender:
                email_data["sender"] = {
                    "name": safe_encode_text(getattr(sender, "Name", ""), "sender_name"),
                    "email": safe_encode_text(getattr(sender, "Address", ""), "sender_email"),
                }
            else:
                email_data["sender"] = "Unknown Sender"
        except Exception as e:
            logger.warning(f"Error extracting sender: {e}")
            email_data["sender"] = "Unknown Sender"
        
        # Recipients
        try:
            to_recipients = []
            if hasattr(item, "To") and item.To:
                # Parse To recipients
                to_list = str(item.To).split(";")
                for recipient in to_list:
                    recipient = recipient.strip()
                    if recipient:
                        to_recipients.append({"name": recipient, "email": ""})
            
            cc_recipients = []
            if hasattr(item, "CC") and item.CC:
                # Parse CC recipients  
                cc_list = str(item.CC).split(";")
                for recipient in cc_list:
                    recipient = recipient.strip()
                    if recipient:
                        cc_recipients.append({"name": recipient, "email": ""})
            
            email_data["to_recipients"] = to_recipients
            email_data["cc_recipients"] = cc_recipients
            
        except Exception as e:
            logger.warning(f"Error extracting recipients: {e}")
            email_data["to_recipients"] = []
            email_data["cc_recipients"] = []
        
        return email_data
        
    except Exception as e:
        logger.error(f"Error extracting email data: {e}")
        return None