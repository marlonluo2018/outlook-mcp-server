import datetime
import win32com.client
from typing import List, Optional, Dict, Any

# Constants
MAX_DAYS = 30

# Email cache for storing retrieved emails by number
email_cache = {}

def connect_to_outlook():
    """Connect to Outlook application using COM"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as e:
        raise Exception(f"Failed to connect to Outlook: {str(e)}")

def get_folder_by_name(namespace, folder_name: str):
    """Get a specific Outlook folder by name"""
    try:
        # First check inbox subfolder
        inbox = namespace.GetDefaultFolder(6)  # 6 is the index for inbox folder
        
        # Check inbox subfolders first (most common)
        for folder in inbox.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
                
        # Then check all folders at root level
        for folder in namespace.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
            
            # Also check subfolders
            for subfolder in folder.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    return subfolder
                    
        # If not found
        return None
    except Exception as e:
        raise Exception(f"Failed to access folder {folder_name}: {str(e)}")

def format_email(mail_item) -> Dict[str, Any]:
    """Format an Outlook mail item into a structured dictionary"""
    try:
        # Extract recipients
        recipients = []
        if mail_item.Recipients:
            for i in range(1, mail_item.Recipients.Count + 1):
                recipient = mail_item.Recipients(i)
                try:
                    recipients.append(f"{recipient.Name} <{recipient.Address}>")
                except:
                    recipients.append(f"{recipient.Name}")
        
        # Format the email data
        email_data = {
            "id": mail_item.EntryID,
            "conversation_id": mail_item.ConversationID if hasattr(mail_item, 'ConversationID') else None,
            "subject": mail_item.Subject,
            "sender": mail_item.SenderName,
            "sender_email": mail_item.SenderEmailAddress,
            "received_time": mail_item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S") if mail_item.ReceivedTime else None,
            "recipients": recipients,
            "body": mail_item.Body,
            "has_attachments": mail_item.Attachments.Count > 0,
            "attachment_count": mail_item.Attachments.Count if hasattr(mail_item, 'Attachments') else 0,
            "unread": mail_item.UnRead if hasattr(mail_item, 'UnRead') else False,
            "importance": mail_item.Importance if hasattr(mail_item, 'Importance') else 1,
            "categories": mail_item.Categories if hasattr(mail_item, 'Categories') else ""
        }
        return email_data
    except Exception as e:
        raise Exception(f"Failed to format email: {str(e)}")

def clear_email_cache():
    """Clear the email cache"""
    global email_cache
    email_cache = {}

def get_emails_from_folder(folder, days: int, search_term: Optional[str] = None, match_all: bool = False):
    """Get emails from a folder with optional search filter"""
    emails_list = []
    
    # Calculate the date threshold
    now = datetime.datetime.now()
    threshold_date = now - datetime.timedelta(days=days)
    
    try:
        # Set up filtering
        folder_items = folder.Items
        folder_items.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
        
        # If we have a search term, apply it
        if search_term:
            print(f"\nDEBUG: Applying search filter for term: {search_term}")
            print(f"DEBUG: match_all mode: {match_all}")
            
            # Parse search terms preserving quoted phrases
            search_terms = []
            in_quote = False
            current_term = ""
            
            for char in search_term:
                if char == '"':
                    if in_quote:
                        # End of quoted term
                        if current_term:
                            search_terms.append(current_term)
                        current_term = ""
                    in_quote = not in_quote
                elif char == " " and not in_quote:
                    # Space outside quote - split term
                    if current_term:
                        search_terms.append(current_term)
                    current_term = ""
                else:
                    current_term += char
            
            # Add any remaining term
            if current_term:
                search_terms.append(current_term)
            
            # Remove empty terms and clean up
            search_terms = [term.strip() for term in search_terms if term.strip()]
            
            # Try to create a filter for subject, sender name or body
            try:
                print("DEBUG: Attempting SQL filter...")
                # Build SQL filter based on match_all mode
                sql_conditions = []
                for term in search_terms:
                    # Escape single quotes for SQL
                    safe_term = term.replace("'", "''")
                    term_conditions = [
                        f"\"urn:schemas:httpmail:subject\" LIKE '%{safe_term}%'",
                        f"\"urn:schemas:httpmail:fromname\" LIKE '%{safe_term}%'",
                        f"\"urn:schemas:httpmail:textdescription\" LIKE '%{safe_term}%'"
                    ]
                    sql_conditions.append("(" + " OR ".join(term_conditions) + ")")
                
                if match_all:
                    filter_term = f"@SQL=" + " AND ".join(sql_conditions)
                else:
                    filter_term = f"@SQL=" + " OR ".join(sql_conditions)
                print(f"DEBUG: SQL filter term: {filter_term}")
                print(f"DEBUG: Parsed search terms: {search_terms}")
                folder_items = folder_items.Restrict(filter_term)
                print("DEBUG: SQL filter applied successfully")
            except Exception as e:
                print(f"DEBUG: SQL filter failed, falling back to manual filter: {str(e)}")
                # If filtering fails, we'll do manual filtering later
        
        # Process emails
        count = 0
        for item in folder_items:
            try:
                if hasattr(item, 'ReceivedTime') and item.ReceivedTime:
                    # Convert to naive datetime for comparison
                    received_time = item.ReceivedTime.replace(tzinfo=None)
                    
                    # Skip emails older than our threshold
                    if received_time < threshold_date:
                        continue
                    
                    # Manual search filter if needed
                    if search_term and folder_items == folder.Items:  # If we didn't apply filter earlier
                        # Split search terms (support both space and OR separator)
                        search_terms = []
                        for part in search_term.split(" OR "):
                            search_terms.extend(part.strip().lower().split())
                        search_terms = [term for term in search_terms if term]
                        
                        # Check matches based on mode
                        if match_all:
                            # All terms must match somewhere
                            found_match = all(
                                any(term in field.lower() for field in [
                                    item.Subject,
                                    item.SenderName,
                                    item.Body
                                ])
                                for term in search_terms
                            )
                        else:
                            # Any term can match anywhere
                            found_match = any(
                                term in field.lower()
                                for term in search_terms
                                for field in [item.Subject, item.SenderName, item.Body]
                            )
                        
                        if not found_match:
                            continue
                    
                    # Format and add the email
                    email_data = format_email(item)
                    emails_list.append(email_data)
                    count += 1
            except Exception as e:
                print(f"Warning: Error processing email: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Error retrieving emails: {str(e)}")
        
    return emails_list
def list_folders() -> str:
    """List all available mail folders in Outlook"""
    try:
        _, namespace = connect_to_outlook()
        result = "Available mail folders:\n\n"
        for folder in namespace.Folders:
            result += f"- {folder.Name}\n"
            for subfolder in folder.Folders:
                result += f"  - {subfolder.Name}\n"
                try:
                    for subsubfolder in subfolder.Folders:
                        result += f"    - {subsubfolder.Name}\n"
                except:
                    pass
        return result
    except Exception as e:
        return f"Error listing mail folders: {str(e)}"

def list_recent_emails(days: int = 7, folder_name: Optional[str] = None) -> str:
    """List email titles from the specified number of days"""
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"
    
    try:
        _, namespace = connect_to_outlook()
        folder = get_folder_by_name(namespace, folder_name) if folder_name else namespace.GetDefaultFolder(6)
        if not folder:
            return f"Error: Folder '{folder_name}' not found"
        
        clear_email_cache()
        emails = get_emails_from_folder(folder, days)
        for i, email in enumerate(emails, 1):
            email_cache[i] = email
        
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails found in {folder_display} from the last {days} days."
        
        return f"Found {len(emails)} emails in {folder_display} from the last {days} days. Use view_email_cache(page=1) to view page #1 of first 5 emails."
    except Exception as e:
        return f"Error retrieving email titles: {str(e)}"

def search_emails(search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = False) -> str:
    """Search emails by contact name or keyword within a time period
    
    Args:
        search_term: Keywords to search for
        days: Number of days to search back (1-30)
        folder_name: Optional folder name (default: Inbox)
        match_all: If True, requires all keywords to match (AND logic)
                   If False, matches any keyword (OR logic, default)
    """
    if not search_term:
        return "Error: Please provide a search term"
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"
    
    try:
        _, namespace = connect_to_outlook()
        folder = get_folder_by_name(namespace, folder_name) if folder_name else namespace.GetDefaultFolder(6)
        if not folder:
            return f"Error: Folder '{folder_name}' not found"
        
        clear_email_cache()
        emails = get_emails_from_folder(folder, days, search_term, match_all)
        for i, email in enumerate(emails, 1):
            email_cache[i] = email
        
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails matching '{search_term}' found in {folder_display} from the last {days} days."
        
        return f"Found {len(emails)} emails matching '{search_term}' in {folder_display} from the last {days} days. Use view_email_cache(page=1) to view page #1 of first 5 emails."
    except Exception as e:
        return f"Error searching emails: {str(e)}"

def view_email_cache(page: int = 1) -> str:
    """View emails from cache in pages of 5"""
    if not email_cache:
        return "Error: No emails in cache. Please use list_recent_emails or search_emails first."
    if not isinstance(page, int) or page < 1:
        return "Error: 'page' must be a positive integer"
    
    total_emails = len(email_cache)
    total_pages = (total_emails + 4) // 5
    if page > total_pages:
        return f"Error: Page {page} does not exist. There are only {total_pages} pages."
    
    start_idx = (page - 1) * 5 + 1
    end_idx = min(page * 5, total_emails)
    
    result = f"Showing emails {start_idx}-{end_idx} of {total_emails} (Page {page}/{total_pages}):\n\n"
    for i in range(start_idx, end_idx + 1):
        email = email_cache[i]
        result += f"Email #{i}\n"
        result += f"Subject: {email['subject']}\n"
        result += f"From: {email['sender']} <{email['sender_email']}>\n"
        result += f"Received: {email['received_time']}\n"
        result += f"Read Status: {'Read' if not email['unread'] else 'Unread'}\n"
        result += f"Has Attachments: {'Yes' if email['has_attachments'] else 'No'}\n\n"
    
    result += f"Use view_email_cache(page={page + 1}) to view next page." if page < total_pages else "This is the last page."
    result += "\nTo view the full content of an email, use the get_email_by_number tool with the email number."
    return result

def get_email_by_number(email_number: int) -> str:
    """Get detailed content of a specific email by its number from the last listing"""
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."
        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."
        
        email_data = email_cache[email_number]
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_data["id"])
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."
        
        result = f"Email #{email_number} Details:\n\n"
        result += f"Subject: {email_data['subject']}\n"
        result += f"From: {email_data['sender']} <{email_data['sender_email']}>\n"
        result += f"Received: {email_data['received_time']}\n"
        result += f"Recipients: {', '.join(email_data['recipients'])}\n"
        result += f"Has Attachments: {'Yes' if email_data['has_attachments'] else 'No'}\n"
        
        if email_data['has_attachments']:
            result += "Attachments:\n"
            for i in range(1, email.Attachments.Count + 1):
                attachment = email.Attachments(i)
                result += f"  - {attachment.FileName}\n"
        
        result += "\nBody:\n"
        result += email_data['body']
        result += "\n\nTo reply to this email, use the reply_to_email_by_number tool with this email number."
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
        if email_number not in email_cache:
            return f"Email #{email_number} not found in current listing."
        
        email_id = email_cache[email_number]["id"]
        outlook, namespace = connect_to_outlook()
        try:
            email = namespace.GetItemFromID(email_id)
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
                html_reply = f"<div>{reply_text}</div>"
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
            return "Failed to send reply - please try again."
        finally:
            email = None
            namespace = None
            outlook = None
    except Exception as e:
        return "An error occurred while processing your request."

def compose_email(to_recipients: List[str], subject: str, body: str, cc_recipients: Optional[List[str]] = None) -> str:
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
        outlook, _ = connect_to_outlook()
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        
        mail.To = "; ".join(to_recipients)
        mail.Subject = subject
        mail.Body = body
        
        if cc_recipients:
            mail.CC = "; ".join(cc_recipients)
            
        mail.Send()
        return "Email sent successfully"
    except Exception as e:
        return f"Failed to send email: {str(e)}"

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
            match_all = input("Match all terms? (y/n, default=n): ").strip().lower() == 'y'
            try:
                print(search_emails(term, int(days), folder, match_all))
            except ValueError:
                print("Invalid days input - must be a number")
                
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