"""Email operations tools for Outlook MCP Server."""

from typing import Union, List, Optional
from ..backend.email_composition import reply_to_email_by_number, compose_email
from ..backend.outlook_session import OutlookSessionManager


def reply_to_email_by_number_tool(
    email_number: int, 
    reply_text: str, 
    to_recipients: Union[str, List[str], None] = None, 
    cc_recipients: Union[str, List[str], None] = None
) -> dict:
    """Reply to an email with custom recipients if provided

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
    
    try:
        result = reply_to_email_by_number(email_number, reply_text, to_recipients, cc_recipients)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error replying to email: {str(e)}"}


def compose_email_tool(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> dict:
    """Compose and send a new email

    Args:
        recipient_email: Email address(es) of the recipient(s) - can be single email or semicolon-separated list
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Optional CC email address(es) - can be single email or semicolon-separated list

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
    
    try:
        result = compose_email(recipient_email, subject, body, cc_email)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error composing email: {str(e)}"}


def move_email_tool(email_number: int, target_folder_name: str) -> dict:
    """Move an email to the specified folder.

    Args:
        email_number: The number of the email in the cache to move (1-based)
        target_folder_name: Name or path of the target folder (supports nested paths like "user@company.com/Inbox/SubFolder1/SubFolder2")

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved successfully to target_folder"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
        After moving, the cache will be cleared to reflect the new email positions.
        
        IMPORTANT: Target folder paths must include the email address as the root folder.
        Use format: "user@company.com/Inbox/SubFolder" not just "Inbox/SubFolder"
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    if not target_folder_name or not isinstance(target_folder_name, str):
        raise ValueError("Target folder name must be a non-empty string")

    try:
        # First get the email from cache to get the email ID
        from ..backend.email_data_extractor import get_email_by_number_unified
        email_data = get_email_by_number_unified(email_number)
        
        if email_data is None:
            return {
                "type": "text",
                "text": f"No email found at position {email_number}. Please load emails first using list_recent_emails or search_emails.",
            }

        # Get the email ID (EntryID) from the cached data
        email_id = email_data.get("id", "")
        if not email_id:
            return {
                "type": "text",
                "text": "Email data is missing the required ID for moving.",
            }

        # Now move the email using the email number
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.move_email_to_folder(email_number, target_folder_name)
            
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error moving email: {str(e)}"}


def delete_email_by_number_tool(email_number: int) -> dict:
    """Move an email to the Deleted Items folder.

    Args:
        email_number: The number of the email in the cache to delete (1-based)

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved to Deleted Items successfully"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
        This tool moves the email to the Deleted Items folder instead of permanently deleting it.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")

    try:
        # First get the email from cache to get the email ID
        from ..backend.email_data_extractor import get_email_by_number_unified
        email_data = get_email_by_number_unified(email_number)
        
        if email_data is None:
            return {
                "type": "text",
                "text": f"No email found at position {email_number}. Please load emails first using list_recent_emails or search_emails.",
            }

        # Get the email ID (EntryID) from the cached data
        email_id = email_data.get("id", "")
        if not email_id:
            return {
                "type": "text",
                "text": "Email data is missing the required ID for deletion.",
            }

        # Now delete the email using the email number
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.delete_email_by_number(email_number)
            
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error deleting email: {str(e)}"}