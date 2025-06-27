from typing import List, Optional
from mcp.server.fastmcp import FastMCP, Context
import outlook_operations as ops

# Initialize FastMCP server
mcp = FastMCP("outlook-assistant")

# MCP Tools - Imported from outlook_operations
@mcp.tool()
def list_folders() -> dict:
    """
    List all available mail folders in Outlook
    
    Returns:
        dict: Response containing folder list in content field
        {
            "content": [{
                "type": "text",
                "text": "Folder list string here"
            }]
        }
        
    Note:
        - Requires Outlook to be running
        - May raise RuntimeError if Outlook connection fails
    """
    return {
        "content": [{
            "type": "text",
            "text": ops.list_folders()
        }]
    }

@mcp.tool()
def list_recent_emails(days: int = 7, folder_name: Optional[str] = None) -> dict:
    """
    Get count of recent emails and load them into email cache.
    
    Args:
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox)
        
    Returns:
        dict: Response containing email count and previews in content field
        {
            "content": [{
                "type": "text",
                "text": "Count and preview string here"
            }]
        }
        
    Notes:
        - Performance limits:
          * MAX_DAYS (7 days) - Limits date range if requested days exceeds
          * MAX_EMAILS (1000) - Stops processing when reached
          * MAX_LOAD_TIME (58s) - Stops processing when exceeded
        - Emails are sorted by ReceivedTime (newest first)
        - May raise RuntimeError if Outlook connection fails
        - Uses batch processing for better performance with large folders
        - Automatically shows first page of results
    """
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    count_result = ops.list_recent_emails(days, folder_name)
    preview_result = ops.view_email_cache(page=1)
    combined_result = f"{count_result}\n\n{preview_result}"
    return {
        "content": [{
            "type": "text",
            "text": combined_result
        }]
    }

@mcp.tool()
def search_emails(search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True) -> dict:
    """
    Search emails and load matching ones into email cache.
    
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
                "text": "Count and preview string here"
            }]
        }
        
    Notes:
        - Performance limits:
          * MAX_DAYS (7 days) - Limits date range if requested days exceeds
          * MAX_EMAILS (1000) - Stops processing when reached
          * MAX_LOAD_TIME (58s) - Stops processing when exceeded
        - Search behavior:
          * Terms with spaces: quoted phrases as single terms (e.g., "project x")
          * Spaces outside quotes split terms
          * AND/OR logic based on match_all parameter
          * Case-insensitive matching
        - Performance optimizations:
          * Searches subject and body text
          * Uses date filtering first to improve speed
        - Restrictions:
          * Field-specific searches (like subject:) are not supported
        - May raise ValueError for invalid search terms
        - May raise RuntimeError if Outlook connection fails
        - Automatically shows first page of results
    """
    if ':' in search_term:
        raise ValueError("Field-specific searches (using ':') are not supported. "
                       "Use plain text search terms only.")
    
    if not search_term or not isinstance(search_term, str):
        raise ValueError("Search term must be a non-empty string")
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    result = ops.search_emails(search_term, days, folder_name, match_all)
    preview = ops.view_email_cache(1)
    return {
        "content": [{
            "type": "text",
            "text": f"{result}\n\n{preview}"
        }]
    }

@mcp.tool()
def view_email_cache(page: int = 1) -> dict:
    """
    View basic information of cached emails (5 emails per page).
    Shows sender, subject, date.
    
    IMPORTANT: Only call this after user explicitly requests to view emails.
    Only call get_email_by_number when user provides specific email number.
    
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
    result = ops.view_email_cache(page)
    return {
        "content": [{
            "type": "text",
            "text": result
        }]
    }

@mcp.tool()
def get_email_by_number(email_number: int) -> dict:
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
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    result = ops.get_email_by_number(email_number)
    return {
        "content": [{
            "type": "text",
            "text": result
        }]
    }

@mcp.tool()
def reply_to_email_by_number(
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
    result = ops.reply_to_email_by_number(email_number, reply_text, to_recipients, cc_recipients)
    return {
        "content": [{
            "type": "text",
            "text": result
        }]
    }

@mcp.tool()
def compose_email(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> dict:
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
    result = ops.compose_email([recipient_email], subject, body, [cc_email] if cc_email else None)
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
        with ops.OutlookSessionManager() as session:
            inbox = session.get_folder()
            print(f"Successfully connected to Outlook. Inbox has {inbox.Items.Count} items.")
            
            # Run the MCP server
            print("Starting MCP server. Press Ctrl+C to stop.")
            mcp.run()
    except Exception as e:
        print(f"Error starting server: {str(e)}")
