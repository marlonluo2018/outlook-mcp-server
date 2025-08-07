from typing import List, Optional
from mcp.server.fastmcp import FastMCP, Context
from backend.outlook_session import OutlookSessionManager
from backend.email_retrieval import (
    list_folders,
    search_emails,
    view_email_cache,
    get_email_by_number,
    list_recent_emails
)
from backend.email_composition import (
    reply_to_email_by_number,
    compose_email
)

# Initialize FastMCP server
mcp = FastMCP("outlook-assistant")

# MCP Tools - Imported from outlook_operations
@mcp.tool()
def get_folder_list_tool() -> dict:
    """Lists all Outlook mail folders as a string representation of a list.
    
    Returns:
        dict: MCP response with format {"content": [{"type": "text", "text": "['Inbox', 'Sent', ...]"}]}
        
    """
    return {
        "content": [{
            "type": "text",
            "text": list_folders()
        }]
    }

@mcp.tool()
def list_recent_emails_tool(days: int = 7, folder_name: Optional[str] = None) -> dict:
   
    """Gets email count and page 1 preview (5 emails).
    
    Args:
        days: Days to look back (1-30, default:7)
        folder_name: Folder to search (default:Inbox)
        
    Returns:
        dict: Combined results in format:
        {
            "content": [{
                "type": "text",
                "text": "Count: X\n\nPreview: Y"
            }]
        }
    """
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    count_result = list_recent_emails(folder_name=folder_name, days=days)
    preview_result = view_email_cache(1)
    combined_result = f"{count_result}\n\n{preview_result}"
    return {
        "content": [{
            "type": "text",
            "text": combined_result
        }]
    }
@mcp.tool()
def search_emails_tool(
    search_term: str,
    days: int = 7,
    folder_name: Optional[str] = None,
    match_all: bool = True
) -> dict:
    """
    Search emails return count and page 1 preview (5 emails).
    
    Args:
        search_term: Plain text search term (no field prefixes allowed)
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox)
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count and previews in content field
        {
            "content": [{
                "type": "text",
                "text": "Count: X\n\nPreview: Y"
            }]
        }

    """
    if ':' in search_term:
        raise ValueError("Field-specific searches (using ':') are not supported. "
                       "Use plain text search terms only.")
    
    if not search_term or not isinstance(search_term, str):
        raise ValueError("Search term must be a non-empty string")
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    result = search_emails(search_term, days, folder_name, match_all=match_all)
    preview = view_email_cache(1)
    return {
        "content": [{
            "type": "text",
            "text": f"{result}\n\n{preview}"
        }]
    }

@mcp.tool()
def view_email_cache_tool(page: int = 1) -> dict:
    """
    View basic information of cached emails (5 emails per page).
    Shows sender, subject, date.
   
    Args:
        page: Page number to view (1-based, each page contains 5 emails)
        
    Returns:
        dict: Response containing email previews in content field
        {
            "content": [{
                "type": "text",
                "text": "Formatted email previews here"
            }]
        }
    """
    result = view_email_cache(page)
    return {
        "content": [{
            "type": "text",
            "text": result
        }]
    }

@mcp.tool()
def get_email_by_number_tool(email_number: int) -> dict:
    """
    Get full email content including body and attachments by its cache number.
    Requires emails to be loaded first via list_recent_emails or search_emails.
    
    Args:
        email_number: The number of the email in the cache (1-based)
        
    Returns:
        dict: Response containing full email details in content field
        {
            "content": [{
                "type": "text",
                "text": "Full email details here"
            }]
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
            
        return {
            "content": [{
                "type": "text",
                "text": result
            }]
        }
    except ValueError as e:
        return {
            "content": [{
                "type": "text",
                "text": f"Error: {str(e)}"
            }],
            "isError": True
        }

@mcp.tool()
def reply_to_email_by_number_tool(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[List[str]] = None,
    cc_recipients: Optional[List[str]] = None
) -> dict:
    """
    IMPORTANT: You MUST get explicit user confirmation before calling this tool.
    Never reply to an email without the user's direct approval.

    Reply to an email with custom recipients if provided
    
    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: Optional list of "To" emails (None preserves original recipients)
        cc_recipients: Optional list of "CC" emails (None preserves original recipients)
        
    Behavior:
        - When both to_recipients and cc_recipients are None:
          * Uses ReplyAll() to maintain original recipients
        - When either parameter is provided:
          * Uses Reply() with specified recipients
          * Any None parameters will result in empty recipient fields

    Returns:
        dict: Response containing confirmation message in content field
        {
            "content": [{
                "type": "text",
                "text": "Confirmation message here"
            }]
        }
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    if not reply_text or not isinstance(reply_text, str):
        raise ValueError("Reply text must be a non-empty string")
    if to_recipients is not None and not isinstance(to_recipients, list):
        raise ValueError("To recipients must be a list or None")
    if cc_recipients is not None and not isinstance(cc_recipients, list):
        raise ValueError("CC recipients must be a list or None")
    result = reply_to_email_by_number(email_number, reply_text, to_recipients, cc_recipients)
    return {
        "content": [{
            "type": "text",
            "text": result
        }]
    }

@mcp.tool()
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
        dict: Response containing confirmation message in content field
        {
            "content": [{
                "type": "text",
                "text": "Confirmation message here"
            }]
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
        "content": [{
            "type": "text",
            "text": result
        }]
    }

# Run the server
if __name__ == "__main__":
    print("Starting Outlook MCP Server...")
    print("Connecting to Outlook...")
    
    try:
        # Test Outlook connection using context manager
        with OutlookSessionManager() as session:
            inbox = session.get_folder()
            print(f"Successfully connected to Outlook. Inbox has {inbox.Items.Count} items.")
            
            # Run the MCP server
            print("Starting MCP server. Press Ctrl+C to stop.")
            mcp.run()
    except Exception as e:
        print(f"Error starting server: {str(e)}")

