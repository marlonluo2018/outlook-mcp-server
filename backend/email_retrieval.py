from datetime import datetime, timedelta
import time
import pythoncom
from typing import List, Dict, Optional
from backend.outlook_session import OutlookSessionManager
from backend.shared import MAX_DAYS, MAX_EMAILS, MAX_LOAD_TIME, email_cache

def get_emails_from_folder(
    search_term: Optional[str] = None,
    days: Optional[int] = None,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> tuple[List[Dict], str]:
    """Retrieve emails from specified folder with batch processing and timeout
    
    Args:
        search_term: Optional search term to filter emails
        days: Optional number of days to look back
        folder_name: Optional folder name (defaults to Inbox)
        match_all: If True (default), all search terms must match (AND logic)
                  If False, any search term can match (OR logic)
    Returns:
        Tuple of (email list, limit note string)
    """
    emails = []
    start_time = time.time()
    retry_count = 0
    limit_note = ""
    
    with OutlookSessionManager() as session:
        while retry_count <= 2:  # MAX_RETRIES from shared.py
            try:
                # Check timeout
                if time.time() - start_time > MAX_LOAD_TIME:
                    limit_note = " (MAX_LOAD_TIME reached)"
                    break
                    
                folder = session.get_folder(folder_name)
                if not folder:
                    return [], " (Folder not found)"
                    
                folder_items = folder.Items
                folder_items.Sort("[ReceivedTime]", True)
                
                # Filter by date range (exact match to original implementation)
                days_to_use = min(days or MAX_DAYS, MAX_DAYS)  # Apply MAX_DAYS limit
                threshold_date = datetime.now() - timedelta(days=days_to_use)
                
                # Apply date filter using Restrict() with original format
                filter_str = f"[ReceivedTime] >= '{threshold_date.strftime('%m/%d/%Y %H:%M %p')}'"
                filtered_items = folder_items.Restrict(filter_str)
                
                # Process emails with MAX_EMAILS limit
                for i, item in enumerate(filtered_items):
                    if len(emails) >= MAX_EMAILS:
                        limit_note = f" (MAX_EMAILS={MAX_EMAILS} reached)"
                        break
                folder_items = folder_items.Restrict(
                    f"[ReceivedTime] >= '{threshold_date.strftime('%m/%d/%Y %H:%M %p')}'"
                )
                
                if search_term:
                    folder_items = _apply_search_filter(
                        folder_items,
                        search_term,
                        match_all=match_all
                    )
                
                # Process emails in batches
                pythoncom.CoInitialize()
                try:
                    total_items = min(folder_items.Count, 10000)  # Safety limit
                    
                    for batch_start in range(1, total_items + 1, MAX_EMAILS):
                        limit_reached = ""
                        if time.time() - start_time > MAX_LOAD_TIME:
                            limit_reached = " (MAX_LOAD_TIME reached)"
                        elif len(emails) >= MAX_EMAILS:
                            limit_reached = " (MAX_EMAILS reached)"
                        
                        if limit_reached:
                            return emails[:MAX_EMAILS], limit_reached
                            
                        batch_end = min(batch_start + MAX_DAYS- 1, total_items)
                        batch_emails = _process_email_batch(folder_items, batch_start, batch_end)
                        emails.extend(batch_emails)
                finally:
                    pythoncom.CoUninitialize()
                    
                if len(emails) >= MAX_EMAILS:
                    limit_note = " (MAX_EMAILS reached)"
                return emails[:MAX_EMAILS], limit_note
                
            except Exception as e:
                retry_count += 1
                if retry_count > 2:
                    raise RuntimeError(f"Failed after {retry_count} retries: {str(e)}")

def _apply_search_filter(folder_items, search_term: str, match_all: bool = True):
    """Apply search filter to folder items
    
    Args:
        folder_items: Outlook folder items to filter
        search_term: Term(s) to search for
        match_all: If True (default), all terms must match (AND logic)
                  If False, any term can match (OR logic)
    """
    terms = search_term.split()
    if not terms:
        return folder_items
        
    if len(terms) == 1:
        filter_term = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{terms[0]}%'"
    else:
        conditions = []
        for term in terms:
            conditions.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{term}%'")
            
        join_op = " AND " if match_all else " OR "
        filter_term = f"@SQL={join_op.join(conditions)}"
        
    return folder_items.Restrict(filter_term)

def _process_email_batch(folder_items, start: int, end: int) -> List[Dict]:
    """Process a batch of emails and update cache"""
    batch_emails = []
    for i in range(start, end + 1):
        try:
            item = folder_items.Item(i)
            if item.Class != 43:  # Skip non-mail items
                continue
                
            email_data = {
                'id': getattr(item, 'EntryID', ''),
                'subject': getattr(item, 'Subject', 'No Subject'),
                'sender': getattr(item, 'SenderName', 'Unknown Sender'),
                'sender_email': getattr(item, 'SenderEmailAddress', ''),
                'received_time': str(getattr(item, 'ReceivedTime', '')),
                'date': str(getattr(item, 'ReceivedTime', '')),  # Backward compat
                'unread': getattr(item, 'UnRead', False),
                'has_attachments': getattr(item, 'Attachments', False).Count > 0,
                'size': getattr(item, 'Size', 0),
                'body': getattr(item, 'Body', '')[:1000],
                'to_recipients': [
                    {
                        'name': getattr(r, 'Name', ''),
                        'address': getattr(r, 'Address', '').split('/')[-1],
                        'type': getattr(r, 'Type', 1)  # 1=To, 2=CC
                    }
                    for r in getattr(item, 'Recipients', [])
                    if hasattr(r, 'Address') or hasattr(r, 'Name')
                ],
                'cc_recipients': [
                    {
                        'name': getattr(r, 'Name', ''),
                        'address': getattr(r, 'Address', '').split('/')[-1],
                        'type': getattr(r, 'Type', 2)  # 1=To, 2=CC
                    }
                    for r in getattr(item, 'Recipients', [])
                    if hasattr(r, 'Address') or hasattr(r, 'Name')
                ]
            }
            batch_emails.append(email_data)
            
            # Update cache
            if email_data['id']:
                email_cache[email_data['id']] = email_data
                
        except Exception as e:
            continue  # Skip problematic emails
            
    return batch_emails

def list_recent_emails(folder_name: str = "Inbox", days: int = None) -> str:
    """Public interface for listing emails (used by CLI)
    Loads emails into cache and returns count message"""
    emails, note = get_emails_from_folder(
        folder_name=folder_name,
        days=days)
    return f"Found {len(emails)} emails from last {days} days. Use 'view_email_cache_tool' to view them.{note}"

def search_emails(
    query: str,
    days: Optional[int] = None,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> str:
    """Public interface for searching emails (used by CLI)
    
    Args:
        query: Search term to match in email subjects
        days: Optional number of days to filter by
        folder_name: Optional folder name to search in
        match_all: If True (default), all terms must match (AND logic)
                  If False, any term can match (OR logic)
    Returns:
        Formatted string with count and note
    """
    emails, note = get_emails_from_folder(
        search_term=query,
        days=days,
        folder_name=folder_name,
        match_all=match_all)
    days_str = f" from last {days} days" if days else ""
    return f"Found {len(emails)} matching emails{days_str}. Use 'view_email_cache_tool' to view them.{note}"

def list_folders() -> List[str]:
    """List all available mail folders"""
    with OutlookSessionManager() as session:
        folders = []
        for folder in session.namespace.Folders:
            folders.append(folder.Name)
            for subfolder in folder.Folders:
                folders.append(f"  {subfolder.Name}")
        return folders

def get_email_by_number(email_number: int) -> Optional[Dict]:
    """Get detailed information for a specific email by its position in cache (1-based index)"""
    if not isinstance(email_number, int) or email_number < 1:
        return None
        
    cache_items = list(email_cache.values())
    if email_number > len(cache_items):
        return None
        
    email = cache_items[email_number - 1]
    
    # Validate cache item is a dictionary
    if not isinstance(email, dict):
        raise ValueError(f"Invalid cache item type: {type(email)}. Expected dict.")
    
    # Create filtered copy without sensitive fields
    filtered_email = {
        'subject': email.get('subject', 'No Subject'),
        'sender': {'name': email.get('sender', 'Unknown Sender')},
        'received_time': email.get('received_time', ''),
        'unread': email.get('unread', False),
        'has_attachments': email.get('has_attachments', False),
        'size': email.get('size', 0),
        'body': email.get('body', ''),
        'to_recipients': [
            {'name': r.get('name', '')}
            for r in email.get('to_recipients', [])
        ],
        'attachments': email.get('attachments', [])
    }
    
    # If email has full details already, return filtered version
    if 'body' in email and 'attachments' in email:
        return filtered_email
        
    # Otherwise fetch full details from Outlook
    with OutlookSessionManager() as session:
        try:
            item = session.namespace.GetItemFromID(email['id'])
            if item.Class != 43:  # Skip non-mail items
                return None
                
            filtered_email.update({
                'body': getattr(item, 'Body', ''),
                'attachments': [
                    {
                        'name': attach.FileName,
                        'size': attach.Size
                    }
                    for attach in item.Attachments
                ] if hasattr(item, 'Attachments') else []
            })
            
            return filtered_email
            
        except Exception as e:
            return None

# Keep alias for backward compatibility
get_email_details = get_email_by_number


def view_email_cache(page: int = 1, per_page: int = 5) -> str:
    """View emails from cache with pagination and detailed info
    
    Args:
        page: Page number (1-based)
        per_page: Items per page
        
    Returns:
        Formatted email previews as string
    """
    if not email_cache:
        return "Error: No emails in cache. Please use list_emails or search_emails first."
    if not isinstance(page, int) or page < 1:
        return "Error: 'page' must be a positive integer"
    
    cache_items = list(email_cache.values())
    total_emails = len(cache_items)
    total_pages = (total_emails + per_page - 1) // per_page
    
    if page > total_pages:
        return f"Error: Page {page} does not exist. There are only {total_pages} pages."
    
    start_idx = (page - 1) * per_page
    end_idx = min(page * per_page, total_emails)
    
    result = f"Showing emails {start_idx + 1}-{end_idx} of {total_emails} (Page {page}/{total_pages}):\n\n"
    for i in range(start_idx, end_idx):
        email = cache_items[i]
        result += f"Email #{i + 1}\n"
        result += f"Subject: {email['subject']}\n"
        
        # Display sender name only (like CC field)
        sender_name = email['sender'].split('/')[0].strip()
        result += f"From: {sender_name}\n"
        
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
    
    result += f"Use view_email_cache_tool(page={page + 1}) to view next page." if page < total_pages else "This is the last page."
    result += "\nCall get_email_details_tool() to get full content of the email."
    
    return result
