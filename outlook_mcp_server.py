from typing import List, Optional
from mcp.server.fastmcp import FastMCP, Context
import outlook_operations as ops

# Initialize FastMCP server
mcp = FastMCP("outlook-assistant")

# MCP Tools - Imported from outlook_operations
@mcp.tool()
def list_folders() -> str:
    """List all available mail folders in Outlook"""
    return ops.list_folders()

@mcp.tool()
def list_recent_emails(days: int = 7, folder_name: Optional[str] = None) -> str:
    """
    Get count of recent emails and load them into email cache.
    Returns number of emails found.
    
    Args:
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox)


    """
    return ops.list_recent_emails(days, folder_name)

@mcp.tool()
def search_emails(search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True) -> str:
    """
    Search emails and load matching ones into email cache.
    Returns count of matches.
    
    Args:
        search_term: Plain text search term (no field prefixes allowed)
        days: Number of days to look back (default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox)
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Note:
    - Search terms with spaces are handled by:
      * Treating quoted phrases as single terms (e.g., "project x")
      * Using spaces outside quotes to split terms
      * Applying AND/OR logic based on match_all parameter
    - Field-specific searches (like subject:) are not supported
    - For subject searches, use the subject parameter in search_emails directly
    """
    if ':' in search_term:
        raise ValueError("Field-specific searches (using ':') are not supported. "
                       "Use plain text search terms only.")
    
    return ops.search_emails(search_term, days, folder_name, match_all)

@mcp.tool()
def view_email_cache(page: int = 1) -> str:
    """
    View basic information of cached emails (5 emails per page).
    Shows sender, subject, date.
    
    IMPORTANT: Only call this after user explicitly requests to view emails.
    Only call get_email_by_number when user provides specific email number.
    
    Args:
        page: Page number to view (1-based, each page contains 5 emails)
    """
    return ops.view_email_cache(page)

@mcp.tool()
def get_email_by_number(email_number: int) -> str:
    """
    Get full email content including body and attachments by its cache number.
    Requires emails to be loaded first via list_recent_emails or search_emails.
    
    Args:
        email_number: The number of the email in the cache (1-based)
    """
    return ops.get_email_by_number(email_number)

@mcp.tool()
def reply_to_email_by_number(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[List[str]] = None,
    cc_recipients: Optional[List[str]] = None
) -> str:
    """
    IMPORTANT: You MUST get explicit user confirmation before calling this tool.
    Never reply to an email without the user's direct approval.

    Reply to an email with custom recipients if provided
    
    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: Optional list of "To" emails
        cc_recipients: Optional list of "CC" emails
    """
    return ops.reply_to_email_by_number(email_number, reply_text, to_recipients, cc_recipients)

@mcp.tool()
def compose_email(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> str:
    """
    IMPORTANT: You MUST get explicit user confirmation before calling this tool.
    Never send an email without the user's direct approval.

    Compose and send a new email
    
    Args:
        recipient_email: Email address of the recipient
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Optional CC email address
    """
    return ops.compose_email([recipient_email], subject, body, [cc_email] if cc_email else None)

# Run the server
if __name__ == "__main__":
    print("Starting Outlook MCP Server...")
    print("Connecting to Outlook...")
    
    try:
        # Test Outlook connection
        outlook, namespace = ops.connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)  # 6 is inbox
        print(f"Successfully connected to Outlook. Inbox has {inbox.Items.Count} items.")
        
        # Run the MCP server
        print("Starting MCP server. Press Ctrl+C to stop.")
        mcp.run()
    except Exception as e:
        print(f"Error starting server: {str(e)}")
