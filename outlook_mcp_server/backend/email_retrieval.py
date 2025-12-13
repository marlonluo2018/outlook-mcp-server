"""Email retrieval functions with consolidated search and improved error handling"""
from datetime import datetime, timedelta, timezone
import time
import logging
from typing import List, Dict, Optional, Tuple

from .outlook_session import OutlookSessionManager
from .shared import MAX_DAYS, MAX_EMAILS, MAX_LOAD_TIME, email_cache
from .utils import OutlookItemClass, build_dasl_filter, get_pagination_info, safe_encode_text
from .validators import EmailSearchParams, EmailListParams, PaginationParams

logger = logging.getLogger(__name__)


def _unified_search(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True,
    search_field: str = "subject"
) -> Tuple[List[Dict], str]:
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
            search_term=search_term,
            days=days,
            folder_name=folder_name,
            match_all=match_all
        )
    except Exception as e:
        logger.error(f"Validation error in _unified_search: {e}")
        raise ValueError(f"Invalid search parameters: {e}")
    
    # Map search field to filter parameter
    field_filters = {
        "subject": {"subject_filter_only": True},
        "sender": {"sender_filter_only": True},
        "recipient": {"recipient_filter_only": True},
        "body": {"body_filter_only": True}
    }
    
    filter_kwargs = field_filters.get(search_field, {"subject_filter_only": True})
    
    emails, note = get_emails_from_folder(
        search_term=params.search_term,
        days=params.days,
        folder_name=params.folder_name,
        match_all=params.match_all,
        **filter_kwargs
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
                **filter_kwargs
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
    body_filter_only: bool = False
) -> Tuple[List[Dict], str]:
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
    email_cache.clear()
    
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
                        field_filter = 'sender'
                    elif recipient_filter_only:
                        field_filter = 'recipient'
                    elif subject_filter_only:
                        field_filter = 'subject'
                    elif body_filter_only:
                        field_filter = 'body'
                    else:
                        field_filter = 'subject'  # default
                    
                    # Build optimized filter using utility function
                    filter_str = build_dasl_filter(search_terms, threshold_date, field_filter, match_all)
                    folder_items = folder_items.Restrict(filter_str)
                    logger.info(f"Applied DASL filter for {field_filter}: {len(search_terms)} terms")
                
                # Sort by received time (newest first)
                folder_items.Sort("[ReceivedTime]", True)
                
                # Process emails with improved error handling
                processed_count = 0
                total_items = folder_items.Count
                logger.info(f"Processing up to {min(total_items, MAX_EMAILS)} emails from {total_items} total")
                
                for i in range(1, total_items + 1):
                    # Check limits
                    if len(emails) >= MAX_EMAILS:
                        limit_note = f" (MAX_EMAILS={MAX_EMAILS} reached)"
                        logger.info(f"Reached MAX_EMAILS limit: {MAX_EMAILS}")
                        break
                    
                    if time.time() - start_time > MAX_LOAD_TIME:
                        limit_note = f" (MAX_LOAD_TIME reached after {processed_count} emails)"
                        logger.warning(f"MAX_LOAD_TIME reached after {processed_count} emails")
                        break
                    
                    try:
                        item = folder_items.Item(i)
                        
                        # Skip non-mail items
                        if item.Class != OutlookItemClass.MAIL_ITEM:
                            continue
                            
                        processed_count += 1
                        received_time = item.ReceivedTime
                        
                        # Convert COM datetime to Python datetime
                        try:
                            if hasattr(received_time, 'year'):
                                received_datetime = received_time
                            else:
                                received_datetime = datetime.strptime(
                                    received_time.strftime('%Y-%m-%d %H:%M:%S'),
                                    '%Y-%m-%d %H:%M:%S'
                                )
                        except Exception as date_error:
                            logger.warning(f"Failed to parse date for email at index {i}: {date_error}")
                            continue
                        
                        # Stop if we've gone past the date threshold
                        if received_datetime < threshold_date:
                            logger.debug(f"Reached date threshold at email {i}")
                            break
                            
                        # Extract email data with safe encoding
                        raw_sender_name = safe_encode_text(
                            getattr(item, 'SenderName', 'Unknown Sender'),
                            'sender_name'
                        )
                        # Clean sender name
                        sender_name = raw_sender_name.split('/')[0].strip()
                        sender_name = sender_name.split('[')[0].strip()
                        sender_name = sender_name.split(',')[0].strip()
                        sender_name = ' '.join(sender_name.split())
                        
                        email_data = {
                            'id': getattr(item, 'EntryID', ''),
                            'subject': safe_encode_text(getattr(item, 'Subject', 'No Subject'), 'subject'),
                            'sender': sender_name,
                            'sender_email': safe_encode_text(getattr(item, 'SenderEmailAddress', ''), 'sender_email'),
                            'received_time': str(received_datetime),
                            'unread': getattr(item, 'UnRead', False),
                            'to_recipients': [{'name': safe_encode_text(getattr(item, 'To', ''), 'to_recipients')}],
                            'cc_recipients': [{'name': safe_encode_text(getattr(item, 'CC', ''), 'cc_recipients')}]
                        }
                        
                        emails.append(email_data)
                        email_cache[email_data['id']] = email_data
                        
                    except Exception as e:
                        failed_count += 1
                        logger.warning(f"Failed to process email at index {i}: {type(e).__name__}: {str(e)}")
                        continue
                
                if failed_count > 0:
                    logger.info(f"Completed with {failed_count} failed emails out of {processed_count} processed")
                
                return emails[:MAX_EMAILS], limit_note
                
            except Exception as e:
                retry_count += 1
                logger.error(f"Error in get_emails_from_folder (attempt {retry_count}/3): {e}")
                if retry_count >= 3:
                    raise RuntimeError(f"Failed after {retry_count} retries: {str(e)}")
                time.sleep(1 * retry_count)  # Simple backoff
    
    return emails, limit_note


def list_recent_emails(folder_name: str = "Inbox", days: int = None) -> str:
    """Public interface for listing emails (used by CLI).
    Loads emails into cache and returns count message.
    """
    try:
        params = EmailListParams(days=days or 7, folder_name=folder_name)
    except Exception as e:
        logger.error(f"Validation error in list_recent_emails: {e}")
        raise ValueError(f"Invalid parameters: {e}")
    
    emails, note = get_emails_from_folder(
        folder_name=params.folder_name,
        days=params.days
    )
    
    days_str = f" from last {params.days} days" if params.days else ""
    logger.info(f"Listed {len(emails)} emails{days_str}")
    return f"Found {len(emails)} emails{days_str}. Use 'view_email_cache_tool' to view them.{note}"


def search_email_by_subject(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> Tuple[List[Dict], str]:
    """Search emails by subject and return list of emails with note."""
    return _unified_search(search_term, days, folder_name, match_all, "subject")


def search_email_by_from(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> Tuple[List[Dict], str]:
    """Search emails by sender name and return list of emails with note."""
    return _unified_search(search_term, days, folder_name, match_all, "sender")


def search_email_by_to(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> Tuple[List[Dict], str]:
    """Search emails by recipient name and return list of emails with note."""
    return _unified_search(search_term, days, folder_name, match_all, "recipient")


def search_email_by_body(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> Tuple[List[Dict], str]:
    """Search emails by body content and return list of emails with note."""
    return _unified_search(search_term, days, folder_name, match_all, "body")


def list_folders() -> List[str]:
    """List all available mail folders."""
    with OutlookSessionManager() as session:
        folders = []
        try:
            for folder in session.namespace.Folders:
                folders.append(folder.Name)
                try:
                    for subfolder in folder.Folders:
                        folders.append(f"  {subfolder.Name}")
                except Exception as e:
                    logger.warning(f"Could not list subfolders for {folder.Name}: {e}")
        except Exception as e:
            logger.error(f"Error listing folders: {e}")
            raise
        return folders


def get_email_by_number(email_number: int) -> Optional[Dict]:
    """Get detailed information for a specific email by its position in cache (1-based index)."""
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
    sender = email.get('sender', 'Unknown Sender')
    if isinstance(sender, dict):
        sender_name = sender.get('name', 'Unknown Sender')
    else:
        sender_name = str(sender)
    
    filtered_email = {
        'subject': email.get('subject', 'No Subject'),
        'sender': sender_name,
        'received_time': email.get('received_time', ''),
        'unread': email.get('unread', False),
        'has_attachments': email.get('has_attachments', False),
        'size': email.get('size', 0),
        'body': email.get('body', ''),
        'to': ', '.join([
            r.get('name', '') if isinstance(r, dict) else str(r)
            for r in email.get('to_recipients', [])
        ]) if email.get('to_recipients') else '',
        'cc': ', '.join([
            r.get('name', '') if isinstance(r, dict) else str(r)
            for r in email.get('cc_recipients', [])
        ]) if email.get('cc_recipients') else '',
        'attachments': email.get('attachments', [])
    }
    
    # If email has full details already, return filtered version
    if 'body' in email and 'attachments' in email:
        return filtered_email
        
    # Otherwise fetch full details from Outlook
    with OutlookSessionManager() as session:
        try:
            item = session.namespace.GetItemFromID(email['id'])
            if item.Class != OutlookItemClass.MAIL_ITEM:
                logger.warning(f"Email {email_number} is not a mail item")
                return None
                
            filtered_email.update({
                'body': safe_encode_text(getattr(item, 'Body', ''), 'body'),
                'attachments': [
                    {
                        'name': safe_encode_text(attach.FileName, 'attachment_name'),
                        'size': attach.Size
                    }
                    for attach in item.Attachments
                ] if hasattr(item, 'Attachments') else []
            })
            
            logger.info(f"Retrieved full details for email #{email_number}")
            return filtered_email
            
        except Exception as e:
            logger.error(f"Error fetching email details for #{email_number}: {e}")
            return None


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
    total_pages = pagination_info['total_pages']
    total_emails = pagination_info['total_items']
    
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
        if email.get('to_recipients'):
            to_names = [r.get('name', '') for r in email['to_recipients']]
            result += f"To: {', '.join(to_names)}\n"
        
        # Display CC recipients if available
        if email.get('cc_recipients'):
            cc_names = [r.get('name', '') for r in email['cc_recipients']]
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

