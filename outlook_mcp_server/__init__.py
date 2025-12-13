from typing import List, Optional, Union
from fastmcp import FastMCP
from .backend.outlook_session import OutlookSessionManager
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