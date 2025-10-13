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
    # Clear cache before loading new emails
    email_cache.clear()
    
    emails = []
    start_time = time.time()
    retry_count = 0
    limit_note = ""
    
    with OutlookSessionManager() as session:
        while retry_count < 3:  # MAX_RETRIES from shared.py
            try:
                # Check timeout
                if time.time() - start_time > MAX_LOAD_TIME:
                    limit_note = " (MAX_LOAD_TIME reached)"
                    break
                    
                folder = session.get_folder(folder_name)
                if not folder:
                    return [], " (Folder not found)"
                    
                folder_items = folder.Items
                folder_items.Sort("[ReceivedTime]", True)  # Newest first
                
                # Get date threshold
                days_to_use = min(days or MAX_DAYS, MAX_DAYS)
                threshold_date = datetime.now().astimezone() - timedelta(days=days_to_use)
                
                    
                # Process emails from newest to oldest
                emails = []
                processed_count = 0
                total_items = folder_items.Count
                if search_term:
                    print(f"Searching for emails containing: '{search_term}'...")
                else:
                    print(f"Loading emails from the last {days_to_use} days...")
                pythoncom.CoInitialize()
                try:
                    for i in range(1, min(total_items + 1, MAX_EMAILS + 1)):
                        current_time = time.time()
                        if current_time - start_time > MAX_LOAD_TIME:
                            limit_note = f" (MAX_LOAD_TIME reached after {processed_count} emails)"
                            break
                        
                            
                        if len(emails) >= MAX_EMAILS:
                            limit_note = f" (MAX_EMAILS={MAX_EMAILS} reached)"
                            break
                            
                        try:
                            item = folder_items.Item(i)
                            if item.Class != 43:  # Skip non-mail items
                                continue
                                
                            processed_count += 1
                            received_time = item.ReceivedTime
                            
                            # Stop if we've gone past the date threshold
                            if received_time < threshold_date:
                                break
                                
                            # Only process if within date range
                            if received_time >= threshold_date:
                                # Simple progress indicator for large folders
                                if processed_count % 50 == 0:
                                    if search_term:
                                        print(f"Found {len(emails)} matching emails so far...")
                                    else:
                                        print(f"Loaded {processed_count} emails...")
                                # Extract just the display name from SenderName (before first '/')
                                sender_name = getattr(item, 'SenderName', 'Unknown Sender').split('/')[0].strip()
                                
                                
                                email_data = {
                                    'id': getattr(item, 'EntryID', ''),
                                    'subject': getattr(item, 'Subject', 'No Subject'),
                                    'sender': sender_name,
                                    'sender_email': getattr(item, 'SenderEmailAddress', ''),
                                    'received_time': str(received_time),
                                    'unread': getattr(item, 'UnRead', False),
                                    'to_recipients': [{'name': getattr(item, 'To', '')}],
                                    'cc_recipients': [{'name': getattr(item, 'CC', '')}]
                                }
                                
                                # Apply search filtering if search term provided
                                if search_term:
                                    search_lower = search_term.lower()
                                    subject_lower = email_data['subject'].lower()
                                    if search_lower in subject_lower:
                                        emails.append(email_data)
                                        email_cache[email_data['id']] = email_data
                                else:
                                    emails.append(email_data)
                                    email_cache[email_data['id']] = email_data
                                
                                
                        except Exception as e:
                            continue  # Skip problematic emails
                finally:
                    pythoncom.CoUninitialize()
                    
                
                return emails[:MAX_EMAILS], limit_note
                
            except Exception as e:
                retry_count += 1
                if retry_count > 2:
                    raise RuntimeError(f"Failed after {retry_count} retries: {str(e)}")


def _are_terms_close(text: str, terms: List[str], max_distance: int = 50) -> bool:
    """
    Check if all terms appear close to each other in the text.
    
    Args:
        text: The text to search in
        terms: List of terms to search for
        max_distance: Maximum distance between terms (in characters)
        
    Returns:
        True if all terms appear close to each other, False otherwise
    """
    if len(terms) <= 1:
        return True
    
    # Find all positions of each term in the text
    term_positions = {}
    for term in terms:
        positions = []
        start = 0
        while True:
            pos = text.find(term, start)
            if pos == -1:
                break
            positions.append(pos)
            start = pos + 1
        term_positions[term] = positions
    
    # Check if there's a combination of positions where all terms are close
    # We'll use a simple approach: check if any term's position is close to any other term's position
    for i, term1 in enumerate(terms):
        for term2 in terms[i+1:]:
            positions1 = term_positions.get(term1, [])
            positions2 = term_positions.get(term2, [])
            
            # Check if any position of term1 is close to any position of term2
            found_close = False
            for pos1 in positions1:
                for pos2 in positions2:
                    if abs(pos1 - pos2) <= max_distance:
                        found_close = True
                        break
                if found_close:
                    break
            
            if not found_close:
                return False
    
    return True


def _process_email_item(item) -> Optional[Dict]:
    """Process a single email item and update cache"""
    try:
        if item.Class != 43:  # Skip non-mail items
            return None
            
        
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
        
        # Update cache
        if email_data['id']:
            email_cache[email_data['id']] = email_data
            
        return email_data
            
    except Exception as e:
        return None  # Skip problematic emails

def list_recent_emails(folder_name: str = "Inbox", days: int = None) -> str:
    """Public interface for listing emails (used by CLI)
    Loads emails into cache and returns count message"""
    if days is not None and not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    if days is not None and (days < 1 or days > 30):
        raise ValueError("Days parameter must be between 1 and 30")
    
    emails, note = get_emails_from_folder(
        folder_name=folder_name,
        days=days)
    days_str = f" from last {days} days" if days else ""
    return f"Found {len(emails)} emails{days_str}. Use 'view_email_cache_tool' to view them.{note}"

def search_emails(
    query: str,
    days: Optional[int] = None,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> str:
    """Public interface for searching emails (used by CLI)
    
    Args:
        query: Search term to match in email subjects (colons are allowed as part of regular text)
        days: Optional number of days to filter by
        folder_name: Optional folder name to search in
        match_all: If True (default), all terms must match (AND logic)
                  If False, any term can match (OR logic)
    Returns:
        Formatted string with count and note
    """
    if not query or not isinstance(query, str):
        raise ValueError("Search term must be a non-empty string")
    
    if days is not None and not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    
    if days is not None and (days < 1 or days > 30):
        raise ValueError("Days parameter must be between 1 and 30")
    
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
        
    # Validate cache structure
    first_email = next(iter(email_cache.values()), None)
    
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
        
        # Display sender name as-is
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
    
    result += f"Use view_email_cache_tool(page={page + 1}) to view next page." if page < total_pages else "This is the last page."
    result += "\nCall get_email_details_tool() to get full content of the email."
    
    return result

# Add specialized search functions for MCP server compatibility
def search_email_by_subject(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> tuple[List[Dict], str]:
    """Search emails by subject and return list of emails with note
    
    Args:
        search_term: Search term to match in email subjects
        days: Number of days to look back (default: 7)
        folder_name: Optional folder name to search (default: Inbox)
        match_all: If True (default), all terms must match (AND logic)
                  If False, any term can match (OR logic)
    
    Returns:
        Tuple of (email list, note string)
    """
    return get_emails_from_folder(
        search_term=search_term,
        days=days,
        folder_name=folder_name,
        match_all=match_all
    )

def search_email_by_from(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> tuple[List[Dict], str]:
    """Search emails by sender name and return list of emails with note
    
    Args:
        search_term: Search term to match in sender name
        days: Number of days to look back (default: 7)
        folder_name: Optional folder name to search (default: Inbox)
        match_all: If True (default), all terms must match (AND logic)
                  If False, any term can match (OR logic)
    
    Returns:
        Tuple of (email list, note string)
    """
    # Get all emails from the specified folder and time period
    emails, note = get_emails_from_folder(
        search_term=None,  # Don't filter by subject
        days=days,
        folder_name=folder_name
    )
    
    # Filter by sender name
    if search_term:
        search_lower = search_term.lower()
        filtered_emails = []
        
        for email in emails:
            sender_name = email.get('sender', '').lower()
            if search_lower in sender_name:
                filtered_emails.append(email)
        
        return filtered_emails, note
    
    return emails, note

def search_email_by_to(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> tuple[List[Dict], str]:
    """Search emails by recipient name and return list of emails with note
    
    Args:
        search_term: Search term to match in recipient name
        days: Number of days to look back (default: 7)
        folder_name: Optional folder name to search (default: Inbox)
        match_all: If True (default), all terms must match (AND logic)
                  If False, any term can match (OR logic)
    
    Returns:
        Tuple of (email list, note string)
    """
    # Get all emails from the specified folder and time period
    emails, note = get_emails_from_folder(
        search_term=None,  # Don't filter by subject
        days=days,
        folder_name=folder_name
    )
    
    # Filter by recipient name
    if search_term:
        search_lower = search_term.lower()
        filtered_emails = []
        
        for email in emails:
            # Check TO recipients
            to_recipients = email.get('to_recipients', [])
            for recipient in to_recipients:
                recipient_name = recipient.get('name', '').lower()
                if search_lower in recipient_name:
                    filtered_emails.append(email)
                    break  # Found match, no need to check other recipients
        
        return filtered_emails, note
    
    return emails, note

def search_email_by_body(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> tuple[List[Dict], str]:
    """Search emails by body content and return list of emails with note
    
    Args:
        search_term: Search term to match in email body
        days: Number of days to look back (default: 7)
        folder_name: Optional folder name to search (default: Inbox)
        match_all: If True (default), all terms must match (AND logic)
                  If False, any term can match (OR logic)
    
    Returns:
        Tuple of (email list, note string)
    """
    # Get all emails from the specified folder and time period (without subject filtering)
    emails, note = get_emails_from_folder(
        search_term=None,  # Don't filter by subject
        days=days,
        folder_name=folder_name
    )
    
    # Filter by body content
    if search_term:
        # Check if search term is enclosed in quotes (exact phrase search)
        is_exact_phrase = (search_term.startswith('"') and search_term.endswith('"')) or \
                          (search_term.startswith("'") and search_term.endswith("'"))
        
        if is_exact_phrase:
            # Remove quotes and search for exact phrase
            search_phrase = search_term[1:-1].lower()
            filtered_emails = []
            
            with OutlookSessionManager() as session:
                pythoncom.CoInitialize()
                try:
                    for email in emails:
                        try:
                            # Get the full email item to access the body
                            item = session.namespace.GetItemFromID(email['id'])
                            if item.Class != 43:  # Skip non-mail items
                                continue
                                
                            body_text = getattr(item, 'Body', '').lower()
                            
                            # Check if the exact phrase is found in the body
                            if search_phrase in body_text:
                                filtered_emails.append(email)
                        except Exception as e:
                            # Skip problematic emails
                            continue
                finally:
                    pythoncom.CoUninitialize()
        else:
            # Original logic for word-based search
            search_terms = search_term.lower().split()
            filtered_emails = []
            
            with OutlookSessionManager() as session:
                pythoncom.CoInitialize()
                try:
                    for email in emails:
                        try:
                            # Get the full email item to access the body
                            item = session.namespace.GetItemFromID(email['id'])
                            if item.Class != 43:  # Skip non-mail items
                                continue
                                
                            body_text = getattr(item, 'Body', '').lower()
                            
                            # Check if all terms match (AND logic) or any term matches (OR logic)
                            if match_all:
                                # All terms must be found in the body
                                if all(term in body_text for term in search_terms):
                                    # Additional check: ensure terms appear close to each other
                                    # This helps filter out emails where terms are scattered throughout
                                    if _are_terms_close(body_text, search_terms):
                                        filtered_emails.append(email)
                            else:
                                # Any term can be found in the body
                                if any(term in body_text for term in search_terms):
                                    filtered_emails.append(email)
                        except Exception as e:
                            # Skip problematic emails
                            continue
                finally:
                    pythoncom.CoUninitialize()
        
        # Update the email cache with only the filtered emails
        email_cache.clear()
        for email in filtered_emails:
            # Get full email details including body for the cache
            try:
                with OutlookSessionManager() as session:
                    pythoncom.CoInitialize()
                    try:
                        item = session.namespace.GetItemFromID(email['id'])
                        if item.Class == 43:  # Only process mail items
                            full_email = _process_email_item(item)
                            if full_email:
                                email_cache[email['id']] = full_email
                    finally:
                        pythoncom.CoUninitialize()
            except Exception:
                # If we can't get full details, use the basic email data
                email_cache[email['id']] = email
        
        return filtered_emails, note
    
    # If no search term provided, return all emails and update cache
    email_cache.clear()
    for email in emails:
        # Get full email details including body for the cache
        try:
            with OutlookSessionManager() as session:
                pythoncom.CoInitialize()
                try:
                    item = session.namespace.GetItemFromID(email['id'])
                    if item.Class == 43:  # Only process mail items
                        full_email = _process_email_item(item)
                        if full_email:
                            email_cache[email['id']] = full_email
                finally:
                    pythoncom.CoUninitialize()
        except Exception:
            # If we can't get full details, use the basic email data
            email_cache[email['id']] = email
    
    return emails, note
