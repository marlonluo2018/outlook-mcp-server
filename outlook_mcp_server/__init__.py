from typing import List, Optional, Union
from fastmcp import FastMCP
from .backend.outlook_session import OutlookSessionManager
# Import shared module to ensure cache is loaded on startup
from .backend import shared
from .backend.email_retrieval import (
    list_folders,
    search_email_by_subject,
    search_email_by_from,
    search_email_by_to,
    search_email_by_body,
    view_email_cache,
    get_email_by_number,
    list_recent_emails
)
from .backend.email_composition import (
    reply_to_email_by_number,
    compose_email
)
from .backend.batch_operations import batch_forward_emails

# Initialize FastMCP server
mcp = FastMCP("outlook-assistant")

# MCP Tools - Imported from outlook_operations
@mcp.tool
def get_folder_list_tool() -> dict:
    """Lists all Outlook mail folders as a string representation of a list.
    
    Returns:
        dict: MCP response with format {"type": "text", "text": "['Inbox', 'Sent', ...]"}
        
    """
    return {
        "type": "text",
        "text": list_folders()
    }

@mcp.tool
def list_recent_emails_tool(days: int = 7, folder_name: Optional[str] = None) -> dict:
    """Load emails into cache and return count message.
    
    Args:
        days: Days to look back (1-30, default:7)
        folder_name: Folder to search (default:Inbox)
        
    Returns:
        dict: Response containing email count message:
        {
            "type": "text",
            "text": "Found X emails from last Y days. Use 'view_email_cache_tool' to view them."
        }
    """
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    count_result = list_recent_emails(folder_name=folder_name, days=days)
    return {
        "type": "text",
        "text": count_result
    }
@mcp.tool
def search_email_by_subject_tool(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> dict:
    """Search email subjects and load matching emails into cache.
    
    This function only searches the email subject field. It does not search in the email body,
    sender name, recipients, or other fields.
    
    Args:
        search_term: Plain text search term (colons are allowed as part of regular text)
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox)
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last Y days. Use 'view_email_cache_tool' to view them."
        }

    """
    if not search_term or not isinstance(search_term, str):
        raise ValueError("Search term must be a non-empty string")
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    emails, note = search_email_by_subject(search_term, days, folder_name, match_all=match_all)
    result = f"Found {len(emails)} matching emails{note}. Use 'view_email_cache_tool' to view them."
    return {
        "type": "text",
        "text": result
    }

@mcp.tool
def search_email_by_sender_name_tool(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> dict:
    """Search emails by sender name and load matching emails into cache.
    
    This function only searches the sender name field. It does not search in the email body,
    subject, recipients, or other fields.
    
    Search by name only, not email address.
    
    Args:
        search_term: Plain text search term for sender name (colons are allowed as part of regular text)
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox)
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last Y days. Use 'view_email_cache_tool' to view them."
        }

    """
    if not search_term or not isinstance(search_term, str):
        raise ValueError("Search term must be a non-empty string")
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    emails, note = search_email_by_from(search_term, days, folder_name, match_all=match_all)
    result = f"Found {len(emails)} matching emails{note}. Use 'view_email_cache_tool' to view them."
    return {
        "type": "text",
        "text": result
    }

@mcp.tool
def search_email_by_recipient_name_tool(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> dict:
    """Search emails by recipient name and load matching emails into cache.
    
    This function only searches the recipient (To) field. It does not search in the email body,
    subject, sender, or other fields.
    
    Search by name only, not email address.
    
    Args:
        search_term: Plain text search term for recipient name (colons are allowed as part of regular text)
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox)
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last Y days. Use 'view_email_cache_tool' to view them."
        }

    """
    if not search_term or not isinstance(search_term, str):
        raise ValueError("Search term must be a non-empty string")
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    emails, note = search_email_by_to(search_term, days, folder_name, match_all=match_all)
    result = f"Found {len(emails)} matching emails{note}. Use 'view_email_cache_tool' to view them."
    return {
        "type": "text",
        "text": result
    }

@mcp.tool
def search_email_by_body_tool(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> dict:
    """Search emails by body content and load matching emails into cache.
    
    This function searches the email body content. It does not search in the subject,
    sender name, recipients, or other fields.
    
    Note: Searching email body is slower than searching other fields as it requires
    loading the full content of each email.
    
    Args:
        search_term: Plain text search term (colons are allowed as part of regular text)
                    For exact phrase matching, enclose the term in quotes (e.g., "red hat partner day")
                    For word-based matching, use the term without quotes (e.g., red hat partner day)
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox)
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last Y days. Use 'view_email_cache_tool' to view them."
        }

    """
    if not search_term or not isinstance(search_term, str):
        raise ValueError("Search term must be a non-empty string")
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    emails, note = search_email_by_body(search_term, days, folder_name, match_all=match_all)
    result = f"Found {len(emails)} matching emails{note}. Use 'view_email_cache_tool' to view them."
    return {
        "type": "text",
        "text": result
    }

@mcp.tool
def view_email_cache_tool(page: int = 1) -> dict:
    """
    View basic information of cached emails (5 emails per page).
    Shows sender, subject, date.
   
    Args:
        page: Page number to view (1-based, each page contains 5 emails)
        
    Returns:
        dict: Response containing email previews
        {
            "type": "text",
            "text": "Formatted email previews here"
        }
    """
    result = view_email_cache(page)
    return {
        "type": "text",
        "text": result
    }

@mcp.tool
def get_email_by_number_tool(email_number: int) -> dict:
    """
    Get full email content including body and attachments by its cache number.
    Requires emails to be loaded first via list_recent_emails or search_emails.
    
    Args:
        email_number: The number of the email in the cache (1-based)
        
    Returns:
        dict: Response containing full email details
        {
            "type": "text",
            "text": "Full email details here"
        }
        
    Raises:
        ValueError: If email number is invalid or no emails are loaded
        RuntimeError: If cache contains invalid data
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
        
    try:
        result = get_email_by_number(email_number)
        if result is None:
            raise ValueError("No email found at that position. Please load emails first using list_recent_emails or search_emails.")
        
        # Format the email details in a specific order
        formatted_text = f"""Subject: {result.get('subject', 'No Subject')}
From: {result.get('sender', 'Unknown Sender')}"""
        
        if result.get('to'):
            formatted_text += f"\nTo: {result.get('to')}"
        
        if result.get('cc'):
            formatted_text += f"\nCC: {result.get('cc')}"
        
        formatted_text += f"""
Date: {result.get('received_time', 'Unknown Date')}

Body:
{result.get('body', 'No body content')}"""
        
        if result.get('attachments'):
            formatted_text += "\n\nAttachments:"
            for attach in result['attachments']:
                formatted_text += f"\n- {attach.get('name', 'Unknown')} ({attach.get('size', 0)} bytes)"
            
        return {
            "type": "text",
            "text": formatted_text
        }
    except ValueError as e:
        return {
            "type": "text",
            "text": f"Error: {str(e)}"
        }

@mcp.tool
def reply_to_email_by_number_tool(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[Union[str, List[str]]] = None,
    cc_recipients: Optional[Union[str, List[str]]] = None
) -> dict:
    """
    IMPORTANT: You MUST get explicit user confirmation before calling this tool.
    Never reply to an email without the user's direct approval.

    Reply to an email with custom recipients if provided
    
    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
                      Examples: "user@company.com" OR ["user@company.com", "boss@company.com"]
        cc_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
                      Examples: "user@company.com" OR ["user@company.com", "boss@company.com"]
        
    Behavior:
        - When both to_recipients and cc_recipients are None:
          * Uses ReplyAll() to maintain original recipients
        - When either parameter is provided:
          * Uses Reply() with specified recipients
          * Any None parameters will result in empty recipient fields
        - Single email strings and lists of email strings are both accepted

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Confirmation message here"
        }
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    if not reply_text or not isinstance(reply_text, str):
        raise ValueError("Reply text must be a non-empty string")
    
    # Handle single email string to list conversion
    if to_recipients is not None and not isinstance(to_recipients, list):
        if isinstance(to_recipients, str):
            to_recipients = [to_recipients]
        else:
            raise ValueError("To recipients must be a string or list of strings")
    if cc_recipients is not None and not isinstance(cc_recipients, list):
        if isinstance(cc_recipients, str):
            cc_recipients = [cc_recipients]
        else:
            raise ValueError("CC recipients must be a string or list of strings")
    result = reply_to_email_by_number(email_number, reply_text, to_recipients, cc_recipients)
    return {
        "type": "text",
        "text": result
    }

@mcp.tool
def compose_email_tool(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> dict:
    """
    IMPORTANT: You MUST get explicit user confirmation before calling this tool.
    Never send an email without the user's direct approval.

    Compose and send a new email
    
    Args:
        recipient_email: Email address of the recipient
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Optional CC email address
        
    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Confirmation message here"
        }
    """
    if not recipient_email or not isinstance(recipient_email, str):
        raise ValueError("Recipient email must be a non-empty string")
    if not subject or not isinstance(subject, str):
        raise ValueError("Subject must be a non-empty string")
    if not body or not isinstance(body, str):
        raise ValueError("Body must be a non-empty string")
    result = compose_email([recipient_email], subject, body, [cc_email] if cc_email else None)
    return {
        "type": "text",
        "text": result
    }

@mcp.tool
def batch_forward_email_tool(
    email_number: int,
    csv_path: str,
    custom_text: str = ""
) -> dict:
    """
    IMPORTANT: You MUST get explicit user confirmation before calling this tool.
    Never forward batch emails without the user's direct approval.

    Forward an email to recipients listed in a CSV file in batches of 500 (Outlook BCC limit).
    
    This function uses an email from your cache as a template and forwards it to multiple recipients
    from a CSV file. The email is sent via BCC to protect recipient privacy.
    
    Args:
        email_number: The number of the email in the cache to use as template (1-based)
        csv_path: Path to CSV file containing recipient email addresses in 'email' column
        custom_text: Optional custom text to prepend to the email body
        
    CSV Format:
        The CSV file must contain a column named 'email' with recipient email addresses.
        Example:
        ```
        email
        user1@example.com
        user2@example.com
        user3@example.com
        ```
        
    Returns:
        dict: Response containing batch sending results
        {
            "type": "text",
            "text": "Batch sending completed for X recipients in Y batches: [detailed results]"
        }
        
    Note:
        - Maximum 500 recipients per batch due to Outlook BCC limitations
        - Invalid email addresses in the CSV will be skipped with warnings
        - The email is sent as BCC to protect recipient privacy
        - Recipients will see it as a forwarded email with "FW:" prefix
        - This function forwards existing emails, it does not compose new ones
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    if not csv_path or not isinstance(csv_path, str):
        raise ValueError("CSV path must be a non-empty string")
    if not isinstance(custom_text, str):
        raise ValueError("Custom text must be a string")
    
    result = batch_forward_emails(email_number, csv_path, custom_text)
    return {
        "type": "text",
        "text": result
    }


@mcp.tool
def create_folder_tool(folder_name: str, parent_folder_name: Optional[str] = None) -> dict:
    """Create a new folder in the specified parent folder.
    
    Args:
        folder_name: Name of the folder to create
        parent_folder_name: Name of the parent folder (optional, defaults to Inbox)
    
    Returns:
        dict: Response containing confirmation message
        {
            "type": "text", 
            "text": "Folder created successfully: folder_path"
        }
    """
    if not folder_name or not isinstance(folder_name, str):
        raise ValueError("Folder name must be a non-empty string")
    if parent_folder_name and not isinstance(parent_folder_name, str):
        raise ValueError("Parent folder name must be a string")
    
    try:
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.create_folder(folder_name, parent_folder_name)
            return {
                "type": "text",
                "text": result
            }
    except Exception as e:
        return {
            "type": "text",
            "text": f"Error creating folder: {str(e)}"
        }

@mcp.tool
def remove_folder_tool(folder_name: str) -> dict:
    """Remove an existing folder.
    
    Args:
        folder_name: Name or path of the folder to remove
    
    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Folder removed successfully"
        }
    """
    if not folder_name or not isinstance(folder_name, str):
        raise ValueError("Folder name must be a non-empty string")
    
    try:
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.remove_folder(folder_name)
            return {
                "type": "text",
                "text": result
            }
    except Exception as e:
        return {
            "type": "text",
            "text": f"Error removing folder: {str(e)}"
        }

@mcp.tool
def move_email_tool(email_number: int, target_folder_name: str) -> dict:
    """Move an email to the specified folder.
    
    Args:
        email_number: The number of the email in the cache to move (1-based)
        target_folder_name: Name or path of the target folder
    
    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved successfully to target_folder"
        }
    
    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
        After moving, the cache will be cleared to reflect the new email positions.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    if not target_folder_name or not isinstance(target_folder_name, str):
        raise ValueError("Target folder name must be a non-empty string")
    
    try:
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.move_email(email_number, target_folder_name)
            return {
                "type": "text",
                "text": result
            }
    except Exception as e:
        return {
            "type": "text",
            "text": f"Error moving email: {str(e)}"
        }

@mcp.tool
def delete_email_tool(email_number: int) -> dict:
    """Delete an email.
    
    Args:
        email_number: The number of the email in the cache to delete (1-based)
    
    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email deleted successfully"
        }
    
    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    
    try:
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.delete_email(email_number)
            return {
                "type": "text",
                "text": result
            }
    except Exception as e:
        return {
            "type": "text",
            "text": f"Error deleting email: {str(e)}"
        }

@mcp.tool
def get_email_policies_tool(email_number: int) -> dict:
    """Get enterprise policies assigned to an email.
    
    Args:
        email_number: The number of the email in the cache to check (1-based)
        
    Returns:
        dict: Response containing assigned policies
        {
            "type": "text",
            "text": "Assigned policies for email:\n- Policy1\n- Policy2\n\nor\nNo policies assigned"
        }
    
    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    
    try:
        # First get the email from cache to get the email ID
        email_data = get_email_by_number(email_number)
        if email_data is None:
            return {
                "type": "text",
                "text": f"No email found at position {email_number}. Please load emails first using list_recent_emails or search_emails."
            }
        
        # Get the email ID (EntryID) from the cached data
        email_id = email_data.get('id', '')
        if not email_id:
            return {
                "type": "text",
                "text": "Email data is missing the required ID for policy retrieval."
            }
        
        # Now get the policies using the email ID
        with OutlookSessionManager() as outlook_session:
            policies = outlook_session.get_email_policies(email_id)
            
            if policies:
                result = "Assigned policies for email:\n" + "\n".join(f"- {policy}" for policy in policies)
            else:
                result = "No policies assigned to this email."
                
            return {
                "type": "text",
                "text": result
            }
    except Exception as e:
        return {
            "type": "text",
            "text": f"Error getting email policies: {str(e)}"
        }






# Main function for UVX entry point
def main():
    """Main function to start the Outlook MCP Server."""
    try:
        # Test Outlook connection using context manager
        with OutlookSessionManager() as session:
            inbox = session.get_folder()
            
            # Run the MCP server
            mcp.run()
    except Exception as e:
        # Use stderr for error messages to avoid interfering with JSON-RPC
        import sys
        print(f"Error starting server: {str(e)}", file=sys.stderr)

# Run the server
if __name__ == "__main__":
    main()