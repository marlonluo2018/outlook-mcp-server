import csv
from datetime import datetime, timedelta
import time
import pythoncom
import win32com.client
from typing import List, Dict, Optional, Any

# Global configuration constants
MAX_DAYS = 7
MAX_EMAILS = 1000
MAX_LOAD_TIME = 58  # seconds

# Global email cache dictionary (key=EntryID, value=formatted email)
email_cache = {}

class OutlookSessionManager:
    """Context manager for Outlook COM session handling"""
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.folder = None
        self._connected = False
        
    def __enter__(self):
        """Initialize Outlook COM objects"""
        self._connect()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up COM objects"""
        self._disconnect()
            
    def _connect(self):
        """Establish COM connection with proper threading"""
        try:
            # Ensure we're in STA mode for Outlook COM
            if pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED) != 0:
                pythoncom.CoUninitialize()
                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            
            # Create Outlook instance with retry
            for attempt in range(3):
                try:
                    self.outlook = win32com.client.Dispatch("Outlook.Application")
                    self.namespace = self.outlook.GetNamespace("MAPI")
                    self._connected = True
                    return
                except Exception as e:
                    if attempt == 2:
                        raise
                    time.sleep(1)
                    
        except Exception as e:
            self._connected = False
            raise RuntimeError(f"Failed to connect to Outlook (HRESULT: {hex(e.hresult) if hasattr(e, 'hresult') else str(e)})")
            
    def _disconnect(self):
        """Clean up COM objects"""
        if self.folder:
            del self.folder
            self.folder = None
        if self.namespace:
            del self.namespace
            self.namespace = None
        if self.outlook:
            del self.outlook
            self.outlook = None
        pythoncom.CoUninitialize()
        self._connected = False
        
    def _ensure_connected(self):
        """Verify connection or reconnect with retry"""
        if not self._connected:
            for attempt in range(3):
                try:
                    self._connect()
                    return
                except Exception as e:
                    if attempt == 2:
                        raise RuntimeError(f"Failed to reconnect after 3 attempts: {str(e)}")
                    time.sleep(2 * (attempt + 1))  # Exponential backoff
            
    def get_folder(self, folder_name: Optional[str] = None):
        """Get specified folder or default inbox"""
        self._ensure_connected()
        try:
            if folder_name:
                return self._get_folder_by_name(folder_name)
            return self.namespace.GetDefaultFolder(6)  # Inbox
        except Exception as e:
            self._connected = False
            raise RuntimeError(f"Failed to access folder: {str(e)}")
        
    def _get_folder_by_name(self, folder_name: str):
        """Find folder by name in folder hierarchy"""
        self._ensure_connected()
        try:
            inbox = self.namespace.GetDefaultFolder(6)
            
            # Check inbox subfolders first
            for folder in inbox.Folders:
                if folder.Name.lower() == folder_name.lower():
                    return folder
                    
            # Check all folders at root level
            for folder in self.namespace.Folders:
                if folder.Name.lower() == folder_name.lower():
                    return folder
                    
                # Check subfolders
                for subfolder in folder.Folders:
                    if subfolder.Name.lower() == folder_name.lower():
                        return subfolder
        except Exception as e:
            self._connected = False
            raise RuntimeError(f"Failed to find folder: {str(e)}")
        return None

def format_email(mail_item) -> Dict[str, Any]:
    """Format an Outlook mail item into a structured dictionary"""
    try:
        if isinstance(mail_item, dict):
            # Handle already formatted emails from cache
            if 'to_recipients' in mail_item and 'cc_recipients' in mail_item:
                # Already properly formatted - return as-is
                return mail_item
            elif 'recipients' in mail_item:
                # Convert recipients format to match our expected structure
                to_recipients = []
                cc_recipients = []
                
                for recipient in mail_item['recipients']:
                    if 'name' in recipient:
                        name = recipient['name'].strip()
                        if 'type' in recipient and recipient['type'] == 2:  # CC recipient
                            cc_recipients.append({'name': name})
                        else:  # Default to To recipient
                            to_recipients.append({'name': name})
                
                return {
                    **mail_item,
                    'to_recipients': to_recipients,
                    'cc_recipients': cc_recipients
                }
            return mail_item
            
        # Get To and CC recipients separately (store emails but don't display)
        to_recipients = []
        cc_recipients = []
        if hasattr(mail_item, 'Recipients') and mail_item.Recipients:
            for i in range(1, mail_item.Recipients.Count + 1):
                recipient = mail_item.Recipients.Item(i)
                if recipient.Type == 1:  # olTo
                    to_recipients.append({'name': recipient.Name})
                elif recipient.Type == 2:  # olCC
                    cc_recipients.append({'name': recipient.Name})
        
        # Fallback for direct To recipients if none found
        if not to_recipients and hasattr(mail_item, 'To'):
            to_recipients = [{'name': name.strip()} for name in mail_item.To.split(';') if name.strip()]
            
        return {
            "id": mail_item.EntryID,
            "subject": getattr(mail_item, 'Subject', 'No Subject'),
            "sender": getattr(mail_item, 'SenderName', 'Unknown Sender'),
            "received_time": mail_item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S") if hasattr(mail_item, 'ReceivedTime') and mail_item.ReceivedTime else None,
            "to_recipients": to_recipients,
            "cc_recipients": cc_recipients,
            "has_attachments": getattr(mail_item, 'Attachments', 0).Count > 0
        }
    except Exception as e:
        raise Exception(f"Failed to format email: {str(e)}")


def get_emails_from_folder(
    folder_name: Optional[str] = None,
    days: int = 7,
    search_term: Optional[str] = None,
    batch_size: int = 50
) -> List[Dict]:
    """
    Retrieve emails from specified folder with batch processing and timeout
    
    Args:
        folder_name: Name of folder to search (None for inbox)
        days: Number of days to look back
        search_term: Optional search filter
        batch_size: Number of emails to process per batch (50-200 recommended)
        max_runtime: Maximum processing time in seconds
        max_emails: Maximum number of emails to return (0 for unlimited)
        
    Returns:
        List of formatted email dictionaries
    """
    global email_cache
    emails = []
    start_time = time.time()
    max_retries = 2
    retry_count = 0
    
    with OutlookSessionManager() as session:
        while retry_count <= max_retries:
            try:
                try:
                    folder = session.get_folder(folder_name)
                    if not folder:
                        print(f"Folder '{folder_name or 'Inbox'}' not found")
                        return []
                except RuntimeError as e:
                    print(f"Connection error (attempt {retry_count + 1}/{max_retries}): {str(e)}")
                    if retry_count == max_retries:
                        raise RuntimeError(f"Failed after {max_retries} retries")
                    time.sleep(3 * (retry_count + 1))
                    retry_count += 1
                    continue
                    
                folder_items = folder.Items
                folder_items.Sort("[ReceivedTime]", True)
                
                # Filter by date range first to reduce processing
                threshold_date = datetime.now() - timedelta(days=days)
                folder_items = folder_items.Restrict(
                    f"[ReceivedTime] >= '{threshold_date.strftime('%m/%d/%Y %H:%M %p')}'"
                )
                
                if search_term:
                    folder_items = _apply_search_filter(folder_items, search_term)
                
                total_items = min(folder_items.Count, 10000)  # Safety limit
                
                # Process in batches with timeout and max emails check
                pythoncom.CoInitialize()
                try:
                    for batch_start in range(1, total_items + 1, batch_size):
                        limit_reached = ""
                        if time.time() - start_time > MAX_LOAD_TIME:
                            limit_reached = " (MAX_LOAD_TIME reached)"
                        elif len(emails) >= MAX_EMAILS:
                            limit_reached = " (MAX_EMAILS reached)"
                        
                        if limit_reached:
                            print(f"Processing completed{limit_reached}")
                            return emails[:MAX_EMAILS], limit_reached
                            
                        batch_end = min(batch_start + batch_size - 1, total_items)
                        batch_emails = []
                        for i in range(batch_start, batch_end + 1):
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
                                    'recipients': [
                                        {
                                            'name': getattr(recipient, 'Name', ''),
                                            'address': getattr(recipient, 'Address', '').split('/')[-1],
                                            'type': getattr(recipient, 'Type', 1)  # 1=To, 2=CC
                                        }
                                        for recipient in getattr(item, 'Recipients', [])
                                        if hasattr(recipient, 'Address') or hasattr(recipient, 'Name')
                                    ]
                                }
                                
                                batch_emails.append(email_data)
                                
                            except Exception as e:
                                print(f"Error processing email {i}: {str(e)}")
                                continue
                                
                        formatted_batch = [format_email(email) for email in batch_emails]
                        emails.extend(formatted_batch)
                        # Add formatted emails to cache
                        for email in formatted_batch:
                            if 'id' in email and email['id']:
                                email_cache[email['id']] = email
                finally:
                    pythoncom.CoUninitialize()
                
                limit_note = " (MAX_EMAILS reached)" if len(emails) > MAX_EMAILS else ""
                return emails[:MAX_EMAILS], limit_note
                
            except Exception as e:
                print(f"Error processing batch (attempt {retry_count + 1}): {str(e)}")
                retry_count += 1
                if retry_count > max_retries:
                    raise

def _apply_search_filter(folder_items, search_term: str):
    """Apply search filter to folder items"""
    try:
        filter_term = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{search_term}%'"
        return folder_items.Restrict(filter_term)
    except Exception as e:
        print(f"Warning: Could not apply search filter - {str(e)}")
        return folder_items

def list_folders() -> List[str]:
    """List all available mail folders with selection numbers"""
    with OutlookSessionManager() as session:
        folders = []
        folder_num = 1
        for folder in session.namespace.Folders:
            folders.append(f"{folder_num}. {folder.Name}")
            folder_num += 1
            for subfolder in folder.Folders:
                folders.append(f"  {folder_num}. {subfolder.Name}")
                folder_num += 1
        return folders

def list_recent_emails(days: int = MAX_DAYS, folder_name: Optional[str] = None) -> str:
    """Get count of recent emails and cache formatted emails"""
    if days > MAX_DAYS:
        actual_days = MAX_DAYS
        note = " (limited to MAX_DAYS for performance)"
    else:
        actual_days = days
        note = ""
    
    email_cache.clear()
    emails, limit_note = get_emails_from_folder(folder_name, actual_days)
    if not emails:
        return f"No emails found in the last {actual_days} days{note}{limit_note}"
    
    for mail in emails:
        if mail.get('id'):
            # Store original mail data without reformatting
            email_cache[mail['id']] = mail
    return f"Found {len(emails)} emails from last {actual_days} days{note}{limit_note}. Use 'view_email_cache' to view them."

def search_emails(search_term: str, days: int = MAX_DAYS, folder_name: Optional[str] = None,
                 match_all: bool = True) -> str:
    """Search emails by term and cache formatted emails"""
    if days > MAX_DAYS:
        actual_days = MAX_DAYS
        note = " (limited to MAX_DAYS for performance)"
    else:
        actual_days = days
        note = ""
    
    email_cache.clear()
    emails, limit_note = get_emails_from_folder(folder_name, actual_days, search_term)
    if not emails:
        return f"No emails found matching '{search_term}' in the last {actual_days} days{note}{limit_note}"
    
    # Store emails with both ID and sequential index for reference
    for idx, mail in enumerate(emails, 1):
        if mail.get('id'):
            mail['cache_index'] = idx  # Add sequential position
            email_cache[mail['id']] = mail
            # Store original mail data without reformatting
            email_cache[mail['id']] = mail
    return f"Found {len(emails)} matching emails from last {actual_days} days{note}{limit_note}. Use 'view_email_cache' to view them."

def view_email_cache(page: int = 1, per_page: int = 5) -> str:
    """
    View emails from cache with pagination and detailed info
    
    Returns:
        str: Formatted email previews as string
    """
    if not email_cache:
        return "Error: No emails in cache. Please use list_recent_emails or search_emails first."
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
    
    result += f"Use view_email_cache(page={page + 1}) to view next page." if page < total_pages else "This is the last page."
    result += "\nCall get_email_by_number() to get full content of the email."
    
    return result

def get_email_by_number(email_number: int) -> str:
    """Get detailed content of a specific email by its number from the last listing"""
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."
            
        # Get email by sequential index (just for validation)
        cache_items = list(email_cache.values())
        if not 1 <= email_number <= len(cache_items):
            return f"Error: Email #{email_number} not found in the current listing."
            
        with OutlookSessionManager() as session:
            email = session.namespace.GetItemFromID(cache_items[email_number - 1]["id"])
            if not email:
                return f"Error: Email #{email_number} could not be retrieved from Outlook."
            
            # Get fresh data directly from Outlook
            subject = getattr(email, 'Subject', 'No Subject')
            sender = getattr(email, 'SenderName', 'Unknown Sender')
            sender_email = getattr(email, 'SenderEmailAddress', 'Unknown').split('/')[-1]
            if sender_email.startswith('CN='):
                sender_email = sender_email.split('=')[-1]
            received_time = getattr(email, 'ReceivedTime', '').strftime("%Y-%m-%d %H:%M:%S") if hasattr(email, 'ReceivedTime') else ''
            
            result = f"Email #{email_number} Details:\n\n"
            result += f"Subject: {subject}\n"
            result += f"From: {sender} <{sender_email}>\n"
            result += f"Received: {received_time}\n"
            
            # Get clean To recipients
            to_recipients = []
            for i in range(1, email.Recipients.Count + 1):
                recipient = email.Recipients.Item(i)
                if recipient.Type == 1:  # 1 = olTo
                    to_recipients.append(f"{recipient.Name} <{recipient.Address.split('/')[-1]}>")
            
            # Get clean CC recipients
            cc_recipients = []
            for i in range(1, email.Recipients.Count + 1):
                recipient = email.Recipients.Item(i)
                if recipient.Type == 2:  # 2 = olCC
                    cc_recipients.append(f"{recipient.Name} <{recipient.Address.split('/')[-1]}>")
            
            result += f"To: {', '.join(to_recipients)}\n"
            if cc_recipients:
                result += f"Cc: {', '.join(cc_recipients)}\n"
        
        result += f"Has Attachments: {'Yes' if getattr(email, 'Attachments', False).Count > 0 else 'No'}\n"
        
        if getattr(email, 'Attachments', False).Count > 0:
            result += "Attachments:\n"
            for i in range(1, email.Attachments.Count + 1):
                attachment = email.Attachments(i)
                result += f"  - {attachment.FileName}\n"
        
        result += "\nBody:\n"
        result += getattr(email, 'Body', 'No body content')
        result += f"\n\nTo reply to this email, first confirm with the user. If approved, call: reply_to_email_by_number(email_number={email_number}, reply_text='your reply text')"
        return result
    except Exception as e:
        return f"Error retrieving email details: {str(e)}"

def reply_to_email_by_number(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[List[str]] = None,
    cc_recipients: Optional[List[str]] = None
) -> str:
    """Reply to an email with custom recipients if provided"""
    try:
        if not email_cache:
            return "No emails available - please list emails first."
            
        cache_items = list(email_cache.values())
        if not 1 <= email_number <= len(cache_items):
            return f"Email #{email_number} not found in current listing."
        
        email_id = cache_items[email_number - 1]["id"]
        with OutlookSessionManager() as session:
            email = session.namespace.GetItemFromID(email_id)
            if not email:
                return "Could not retrieve the email from Outlook."
            
            if to_recipients is None and cc_recipients is None:
                reply = email.ReplyAll()
            else:
                reply = email.Reply()
                if to_recipients:
                    reply.To = "; ".join(to_recipients)
                if cc_recipients:
                    reply.CC = "; ".join(cc_recipients)

            if hasattr(email, 'HTMLBody') and email.HTMLBody:
                # Convert newlines to HTML line breaks
                formatted_reply = reply_text.replace('\n\n', '<br><br>').replace('\n', '<br>')
                html_reply = f"<div>{formatted_reply}</div>"
                html_reply += "<div style='border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0in 0in 0in'>"
                html_reply += f"<div><b>From:</b> {email.SenderName} <{email.SenderEmailAddress}></div>"
                html_reply += f"<div><b>Sent:</b> {email.ReceivedTime.strftime('%A, %B %d, %Y %I:%M %p')}</div>"
                html_reply += f"<div><b>To:</b> {email.To}</div>"
                if email.CC:
                    html_reply += f"<div><b>Cc:</b> {email.CC}</div>"
                html_reply += f"<div><b>Subject:</b> {email.Subject}</div><br>"
                
                html_body = email.HTMLBody
                if ("This Message Is From an Untrusted Sender" in html_body or
                    any("BannerStart" in part for part in html_body.split("<p>")) or
                    "proofpoint.com/EWT" in html_body):
                    banner_start = -1
                    p_tags = html_body.split("<p>")
                    for i in range(1, len(p_tags)):
                        if "BannerStart" in p_tags[i].split(">")[0]:
                            banner_start = sum(len(p_tags[j]) + 3 for j in range(i))
                            break
                    
                    if banner_start != -1:
                        html_body = (html_body[:banner_start+2] +
                                    " style='font-size:0.8em; background-color:#f5f5f5; padding:5px; margin-bottom:10px;'" +
                                    html_body[banner_start+2:])
                
                html_reply += html_body
                html_reply += "</div>"
                reply.HTMLBody = html_reply
            else:
                reply.Body = f"{reply_text}\n\n" + \
                            f"From: {email.SenderName}\n" + \
                            f"Sent: {email.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')}\n" + \
                            f"To: {email.To}\n" + \
                            (f"Cc: {email.CC}\n" if email.CC else "") + \
                            f"Subject: {email.Subject}\n\n" + \
                            email.Body
            
            reply.Send()
            return "Reply sent successfully."
    except Exception as e:
        return f"Failed to send reply: {str(e)}"

def compose_email(to_recipients: List[str], subject: str, body: str,
                 cc_recipients: Optional[List[str]] = None) -> str:
    """Compose and send a new email using Outlook COM API
    
    Args:
        to_recipients: List of email addresses for To field
        subject: Email subject
        body: Email body content
        cc_recipients: Optional list of email addresses for CC field
        
    Returns:
        str: Success/error message
    """
    try:
        with OutlookSessionManager() as session:
            mail = session.outlook.CreateItem(0)  # 0 = olMailItem
            
            mail.To = "; ".join(to_recipients)
            mail.Subject = subject
            
            # Detect HTML content and handle appropriately
            if any(tag in body.lower() for tag in ['<div>', '<p>', '<br>', '<html>']):
                # For HTML content, replace newlines with <br> and set HTMLBody
                html_body = body.replace('\n\n', '<br>').replace('\n', '<br>')
                mail.HTMLBody = html_body
            else:
                # Plain text content
                mail.Body = body
            
            if cc_recipients:
                mail.CC = "; ".join(cc_recipients)
                
            mail.Send()
            return "Email sent successfully"
    except Exception as e:
        return f"Failed to send email: {str(e)}"

def send_batch_emails(email_number: int, csv_path: str, custom_text: str = "") -> str:
    """Send email to recipients in batches of 500 (Outlook BCC limit)
    
    Args:
        email_number: Email number from cache to use as template
        csv_path: Path to CSV file containing recipient emails (one per line)
        custom_text: Additional text to prepend to email body
        
    Returns:
        str: Status message with batch sending results
    """
    try:
        # Find email in cache by sequential number
        template = None
        for mail in email_cache.values():
            if mail.get('cache_index') == email_number:
                template = mail
                break
                
        if not template:
            return f"Error: Email #{email_number} not found in cache"
            
        # Validate template has required fields
        if not all(key in template for key in ['subject', 'body']):
            return "Error: Cached email missing required fields (subject/body)"
        
        # Read and validate recipient emails from CSV
        try:
            # Strip quotes from path if present
            clean_path = csv_path.strip('"\'')
            with open(clean_path, 'r', newline='', encoding='utf-8-sig') as csvfile:
                reader = csv.reader(csvfile)
                recipients = []
                for row_num, row in enumerate(reader, 1):
                    if not row:
                        continue
                    email = row[0].strip()
                    if not email:
                        continue
                    if '@' not in email:
                        return f"Invalid email format in CSV row {row_num}: {email}"
                    recipients.append(email)
                
                if not recipients:
                    return "Error: No valid email addresses found in CSV"
        except Exception as e:
            return f"Error reading CSV file: {str(e)}"
            
        if not recipients:
            return "Error: No valid email addresses found in CSV"
            
        # Process in batches of 500
        batch_size = 500
        total_recipients = len(recipients)
        batches = [recipients[i:i + batch_size] for i in range(0, total_recipients, batch_size)]
        
        results = []
        with OutlookSessionManager() as session:
            for i, batch in enumerate(batches, 1):
                try:
                    # Forward original email with all recipients in BCC
                    email_id = template['id']
                    original = session.namespace.GetItemFromID(email_id)
                    if not original:
                        results.append(f"Error sending batch {i}: Could not retrieve original email")
                        continue
                        
                    fwd = original.Forward()
                    
                    # Remove any automatically added recipients
                    while fwd.Recipients.Count > 0:
                        fwd.Recipients(1).Delete()
                    
                    # Add all recipients to BCC only
                    for recipient in batch:
                        new_recip = fwd.Recipients.Add(recipient)
                        new_recip.Type = 3  # 3 = olBCC
                    
                    # Prepend custom text if provided
                    if custom_text:
                        if fwd.BodyFormat == 2:  # 2 = olFormatHTML
                            fwd.HTMLBody = f"<div>{custom_text}</div><br><br>{fwd.HTMLBody}"
                        else:
                            fwd.Body = f"{custom_text}\n\n{fwd.Body}"
                    
                    fwd.Send()
                    results.append(f"Batch {i} ({len(batch)} emails) sent successfully")
                except Exception as e:
                    results.append(f"Error sending batch {i}: {str(e)}")
                    
        return "\n".join([
            f"Batch sending completed for {total_recipients} recipients in {len(batches)} batches:",
            *results
        ])
    except Exception as e:
        return f"Error in batch sending process: {str(e)}"

if __name__ == "__main__":
    """Command line interface for Outlook operations"""
    print("Outlook Operations CLI")
    print("----------------------")
    
    while True:
        print("\nAvailable commands:")
        print("1. List folders")
        print("2. List recent emails")
        print("3. Search emails")
        print("4. View email cache")
        print("5. Get email details")
        print("6. Reply to email")
        print("7. Compose new email")
        print("8. Send batch emails")
        print("0. Exit")
        
        choice = input("\nEnter command number: ").strip()
        
        if choice == "0":
            print("Exiting...")
            break
            
        elif choice == "1":
            print(list_folders())
            
        elif choice == "2":
            days = input("Enter number of days (1-30): ").strip()
            folder = input("Enter folder name (leave blank for Inbox): ").strip() or None
            try:
                print(list_recent_emails(int(days), folder))
            except ValueError:
                print("Invalid days input - must be a number")
                
        elif choice == "3":
            term = input("Enter search term: ").strip()
            days = input("Enter number of days (1-30): ").strip()
            folder = input("Enter folder name (leave blank for Inbox): ").strip() or None
            match_all = input("Match all terms? (y/n, default=y): ").strip().lower() != 'n'
            try:
                print(search_emails(term, int(days), folder, match_all))
            except ValueError:
                print("Invalid days input - must be a number")

        elif choice == "8":
            # Send batch emails
            try:
                email_num = input("Enter email number from cache: ").strip()
                if not email_num.isdigit():
                    print("Error: Email number must be numeric")
                    continue
                    
                csv_path = input("Enter path to CSV file: ").strip()
                custom_text = input("Enter custom text to prepend (optional): ").strip()
                
                result = send_batch_emails(int(email_num), csv_path, custom_text)
                print(result)
            except Exception as e:
                print(f"Error: {str(e)}")
                
        elif choice == "4":
            page = input("Enter page number (default 1): ").strip() or "1"
            try:
                print(view_email_cache(int(page)))
            except ValueError:
                print("Invalid page number - must be a number")
                
        elif choice == "5":
            num = input("Enter email number: ").strip()
            try:
                print(get_email_by_number(int(num)))
            except ValueError:
                print("Invalid email number - must be a number")
                
        elif choice == "6":
            num = input("Enter email number to reply to: ").strip()
            text = input("Enter reply text: ").strip()
            to = input("Enter To recipients (comma separated, blank for reply-all): ").strip()
            cc = input("Enter CC recipients (comma separated, blank for none): ").strip()
            try:
                to_list = [x.strip() for x in to.split(",")] if to else None
                cc_list = [x.strip() for x in cc.split(",")] if cc else None
                print(reply_to_email_by_number(int(num), text, to_list, cc_list))
            except ValueError:
                print("Invalid email number - must be a number")
                
        elif choice == "7":
            to = input("Enter To recipients (comma separated): ").strip()
            subject = input("Enter subject: ").strip()
            body = input("Enter email body: ").strip()
            cc = input("Enter CC recipients (comma separated, blank for none): ").strip()
            try:
                to_list = [x.strip() for x in to.split(",")] if to else []
                cc_list = [x.strip() for x in cc.split(",")] if cc else []
                print(compose_email(to_list, subject, body, cc_list))
            except Exception as e:
                print(f"Error composing email: {str(e)}")
                
        else:
            print("Invalid command - please try again")