import datetime
import os
import win32com.client
from typing import List, Optional, Dict, Any
from mcp.server.fastmcp import FastMCP, Context

# Initialize FastMCP server
mcp = FastMCP("outlook-assistant")

# Constants
MAX_DAYS = 30
# Email cache for storing retrieved emails by number
email_cache = {}

# Helper functions
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

def get_emails_from_folder(folder, days: int, search_term: Optional[str] = None):
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
            # Handle OR operators in search term
            search_terms = [term.strip() for term in search_term.split(" OR ")]
            
            # Try to create a filter for subject, sender name or body
            try:
                # Build SQL filter with OR conditions for each search term
                sql_conditions = []
                for term in search_terms:
                    sql_conditions.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{term}%'")
                    sql_conditions.append(f"\"urn:schemas:httpmail:fromname\" LIKE '%{term}%'")
                    sql_conditions.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{term}%'")
                
                filter_term = f"@SQL=" + " OR ".join(sql_conditions)
                folder_items = folder_items.Restrict(filter_term)
            except:
                # If filtering fails, we'll do manual filtering later
                pass
        
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
                        # Handle OR operators in search term for manual filtering
                        search_terms = [term.strip().lower() for term in search_term.split(" OR ")]
                        
                        # Check if any of the search terms match
                        found_match = False
                        for term in search_terms:
                            if (term in item.Subject.lower() or 
                                term in item.SenderName.lower() or 
                                term in item.Body.lower()):
                                found_match = True
                                break
                        
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

# MCP Tools
@mcp.tool()
def list_folders() -> str:
    """
    List all available mail folders in Outlook
    
    Returns:
        A list of available mail folders
    """
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        result = "Available mail folders:\n\n"
        
        # List all root folders and their subfolders
        for folder in namespace.Folders:
            result += f"- {folder.Name}\n"
            
            # List subfolders
            for subfolder in folder.Folders:
                result += f"  - {subfolder.Name}\n"
                
                # List subfolders (one more level)
                try:
                    for subsubfolder in subfolder.Folders:
                        result += f"    - {subsubfolder.Name}\n"
                except:
                    pass
        
        return result
    except Exception as e:
        return f"Error listing mail folders: {str(e)}"

@mcp.tool()
def list_recent_emails(days: int = 7, folder_name: Optional[str] = None) -> str:
    """
    List email titles from the specified number of days
    
    Args:
        days: Number of days to look back for emails (max 30)
        folder_name: Name of the folder to check (if not specified, checks the Inbox)
        
    Returns:
        Number of emails found and instructions to view them
    """
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the appropriate folder
        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder '{folder_name}' not found"
        else:
            folder = namespace.GetDefaultFolder(6)  # Default inbox
        
        # Clear previous cache
        clear_email_cache()
        
        # Get emails from folder
        emails = get_emails_from_folder(folder, days)
        
        # Cache emails
        for i, email in enumerate(emails, 1):
            email_cache[i] = email
        
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails found in {folder_display} from the last {days} days."
        
        return f"Found {len(emails)} emails in {folder_display} from the last {days} days. Use view_email_cache(page=1) to view page #1 of first 5 emails."
    
    except Exception as e:
        return f"Error retrieving email titles: {str(e)}"

@mcp.tool()
def search_emails(search_term: str, days: int = 7, folder_name: Optional[str] = None) -> str:
    """
    Search emails by contact name or keyword within a time period
    
    Args:
        search_term: Name or keyword to search for
        days: Number of days to look back (max 30)
        folder_name: Name of the folder to search (if not specified, searches the Inbox)
        
    Returns:
        Number of matching emails found and instructions to view them
    """
    if not search_term:
        return "Error: Please provide a search term"
        
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the appropriate folder
        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder '{folder_name}' not found"
        else:
            folder = namespace.GetDefaultFolder(6)  # Default inbox
        
        # Clear previous cache
        clear_email_cache()
        
        # Get emails matching search term
        emails = get_emails_from_folder(folder, days, search_term)
        
        # Cache emails
        for i, email in enumerate(emails, 1):
            email_cache[i] = email
        
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails matching '{search_term}' found in {folder_display} from the last {days} days."
        
        return f"Found {len(emails)} emails matching '{search_term}' in {folder_display} from the last {days} days. Use view_email_cache(page=1) to view page #1 of first 5 emails."
    
    except Exception as e:
        return f"Error searching emails: {str(e)}"

@mcp.tool()
def view_email_cache(page: int = 1) -> str:
    """
    View emails from cache in pages of 5
    
    Args:
        page: Page number to view (1-based)
        
    Returns:
        List of 5 emails from the specified page
    """
    if not email_cache:
        return "Error: No emails in cache. Please use list_recent_emails or search_emails first."
    
    if not isinstance(page, int) or page < 1:
        return "Error: 'page' must be a positive integer"
    
    total_emails = len(email_cache)
    total_pages = (total_emails + 4) // 5  # Ceiling division
    
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


@mcp.tool()
def get_email_by_number(email_number: int) -> str:
    """
    Get detailed content of a specific email by its number from the last listing
    
    Args:
        email_number: The number of the email from the list results
        
    Returns:
        Full details of the specified email
    """
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."
        
        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."
        
        email_data = email_cache[email_number]
        
        # Connect to Outlook to get the full email content
        _, namespace = connect_to_outlook()
        
        # Retrieve the specific email
        email = namespace.GetItemFromID(email_data["id"])
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."
        
        # Format the output
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

@mcp.tool()
def reply_to_email_by_number(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[List[str]] = None,
    cc_recipients: Optional[List[str]] = None
) -> str:
    """
    Reply to an email. Uses custom recipients if provided; otherwise replies to all.
    Matches Outlook's native reply behavior including message history format.
    
    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: List of "To" emails (if provided, replaces original recipients)
        cc_recipients: List of "CC" emails (if provided, replaces original recipients)
    """
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
            
            # Create reply based on recipient preference
            if to_recipients is None and cc_recipients is None:
                reply = email.ReplyAll()  # Default reply to all
            else:
                reply = email.Reply()     # Start fresh for custom recipients
                if to_recipients:
                    reply.To = "; ".join(to_recipients)
                if cc_recipients:
                    reply.CC = "; ".join(cc_recipients)

            # Preserve HTML formatting if available
            if hasattr(email, 'HTMLBody') and email.HTMLBody:
                # Format HTML reply with original message history
                html_reply = f"<div>{reply_text}</div>"
                html_reply += "<div style='border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0in 0in 0in'>"
                html_reply += f"<div><b>From:</b> {email.SenderName} <{email.SenderEmailAddress}></div>"
                html_reply += f"<div><b>Sent:</b> {email.ReceivedTime.strftime('%A, %B %d, %Y %I:%M %p')}</div>"
                html_reply += f"<div><b>To:</b> {email.To}</div>"
                if email.CC:
                    html_reply += f"<div><b>Cc:</b> {email.CC}</div>"
                html_reply += f"<div><b>Subject:</b> {email.Subject}</div><br>"
                
                # Process HTML body to modify security banner styling
                html_body = email.HTMLBody
                # Look for multiple indicators of security banner
                if ("This Message Is From an Untrusted Sender" in html_body or
                    any("BannerStart" in part for part in html_body.split("<p>")) or
                    "proofpoint.com/EWT" in html_body):
                    # Find the opening <p> tag containing BannerStart
                    banner_start = -1
                    p_tags = html_body.split("<p>")
                    for i in range(1, len(p_tags)):
                        if "BannerStart" in p_tags[i].split(">")[0]:
                            banner_start = sum(len(p_tags[j]) + 3 for j in range(i))  # +3 for "<p>"
                            break
                    
                    if banner_start != -1:
                        # Insert style into the <p> tag
                        html_body = (html_body[:banner_start+2] +
                                    " style='font-size:0.8em; background-color:#f5f5f5; padding:5px; margin-bottom:10px;'" +
                                    html_body[banner_start+2:])
                
                html_reply += html_body
                html_reply += "</div>"
                
                reply.HTMLBody = html_reply
                # Plain text version with proper formatting
                # Create HTML version with proper page breaker
                html_reply = f"<div>{reply_text}</div>"
                html_reply += "<div style='border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0in 0in 0in'>"
                html_reply += f"<div><b>From:</b> {email.SenderName}</div>"
                html_reply += f"<div><b>Sent:</b> {email.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')}</div>"
                html_reply += f"<div><b>To:</b> {email.To}</div>"
                if email.CC:
                    html_reply += f"<div><b>Cc:</b> {email.CC}</div>"
                html_reply += f"<div><b>Subject:</b> {email.Subject}</div><br>"
                html_reply += email.HTMLBody if hasattr(email, 'HTMLBody') and email.HTMLBody else email.Body
                html_reply += "</div>"
                
                reply.HTMLBody = html_reply
                
                # Plain text version with minimal formatting
                reply.Body = f"{reply_text}\n\n" + \
                            f"From: {email.SenderName}\n" + \
                            f"Sent: {email.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')}\n" + \
                            f"To: {email.To}\n" + \
                            (f"Cc: {email.CC}\n" if email.CC else "") + \
                            f"Subject: {email.Subject}\n\n" + \
                            email.Body
            else:
                # Fallback to plain text formatting
                reply_body = f"{reply_text}\n\n"
                reply_body += f"From: {email.SenderName}\n"
                reply_body += f"Sent: {email.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')}\n"
                reply_body += f"To: {email.To}\n"
                if email.CC:
                    reply_body += f"Cc: {email.CC}\n"
                reply_body += f"Subject: {email.Subject}\n\n"
                reply_body += email.Body
                reply.Body = reply_body
            reply.Send()
            
            return "Reply sent successfully."
            
        except Exception as e:
            return "Failed to send reply - please try again."
            
        finally:
            # Clean up COM objects
            email = None
            namespace = None
            outlook = None
            
    except Exception as e:
        return "An error occurred while processing your request."

@mcp.tool()
def compose_email(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> str:
    """
    Compose and send a new email
    
    Args:
        recipient_email: Email address of the recipient
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Email address for CC (optional)
        
    Returns:
        Status message indicating success or failure
    """
    try:
        # Connect to Outlook
        outlook, _ = connect_to_outlook()
        
        # Create a new email
        mail = outlook.CreateItem(0)  # 0 is the value for a mail item
        mail.Subject = subject
        mail.To = recipient_email
        
        if cc_email:
            mail.CC = cc_email
        
        # Add signature to the body
        mail.Body = body
        
        # Send the email
        mail.Send()
        
        return f"Email sent successfully to: {recipient_email}"
    
    except Exception as e:
        return f"Error sending email: {str(e)}"

# Run the server
if __name__ == "__main__":
    print("Starting Outlook MCP Server...")
    print("Connecting to Outlook...")
    
    try:
        # Test Outlook connection
        outlook, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)  # 6 is inbox
        print(f"Successfully connected to Outlook. Inbox has {inbox.Items.Count} items.")
        
        # Run the MCP server
        print("Starting MCP server. Press Ctrl+C to stop.")
        mcp.run()
    except Exception as e:
        print(f"Error starting server: {str(e)}")
