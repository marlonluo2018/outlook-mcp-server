from typing import List, Optional, Union
from fastmcp import FastMCP
from .backend.outlook_session import OutlookSessionManager

# Import shared module to ensure cache is loaded on startup
from .backend import shared
from .backend.email_search import (
    list_folders,
    search_email_by_subject,
    search_email_by_from,
    search_email_by_to,
    search_email_by_body,
    view_email_cache,
    list_recent_emails,
    get_emails_from_folder,
)
from .backend.email_data_extractor import get_email_by_number_unified, format_email_with_media
from .backend.shared import clear_email_cache, add_email_to_cache, save_email_cache
from .backend.email_composition import reply_to_email_by_number, compose_email
from .backend.batch_operations import batch_forward_emails

# Initialize FastMCP server
mcp = FastMCP("outlook-assistant")

# MCP Tools - Imported from outlook_operations


@mcp.tool
def move_folder_tool(source_folder_path: str, target_parent_path: str) -> dict:
    """Move a folder and all its emails to a new location.

    Args:
        source_folder_path: Path to the source folder (e.g., "Inbox/SubFolder1")
        target_parent_path: Path to the target parent folder (e.g., "Inbox/NewParent")

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Folder moved successfully from 'source_path' to 'target_path' (X emails moved)"
        }

    Note:
        This tool moves the entire folder structure and all emails contained within it.
        Cannot be used to move default folders like Inbox, Sent Items, etc.
    """
    if not source_folder_path or not isinstance(source_folder_path, str):
        raise ValueError("Source folder path must be a non-empty string")
    if not target_parent_path or not isinstance(target_parent_path, str):
        raise ValueError("Target parent path must be a non-empty string")

    try:
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.move_folder(source_folder_path, target_parent_path)
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error moving folder: {str(e)}"}


@mcp.tool
def get_folder_list_tool() -> dict:
    """Lists all Outlook mail folders as a string representation of a list.

    Returns:
        dict: MCP response with format {"type": "text", "text": "['Inbox', 'Sent', ...]"}

    """
    return {"type": "text", "text": list_folders()}


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
    # Input validation
    if not isinstance(days, int):
        raise ValueError("Days parameter must be an integer")
    
    # Ensure folder_name is a string or None
    if folder_name is not None and not isinstance(folder_name, str):
        raise ValueError("Folder name must be a string or None")
    
    # Log the request for debugging
    import logging
    logger = logging.getLogger(__name__)
    logger.info(f"list_recent_emails_tool called with days={days}, folder_name={folder_name}")
    
    # If folder_name is None, use the default
    if folder_name is None:
        folder_name = "Inbox"
    
    try:
        logger.info(f"Calling list_recent_emails with folder={folder_name}, days={days}")
        count_result = list_recent_emails(folder_name=folder_name, days=days)
        logger.info(f"list_recent_emails returned: {count_result}")
        
        # Add debugging info
        from .backend.shared import email_cache
        debug_info = f"\n\nDEBUG: Cache contains {len(email_cache)} emails"
        
        return {"type": "text", "text": count_result + debug_info}
    except Exception as e:
        logger.error(f"Error in list_recent_emails_tool: {e}")
        import traceback
        traceback.print_exc()
        return {"type": "text", "text": f"Error retrieving emails: {str(e)}"}


@mcp.tool
def search_email_by_subject_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
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
    return {"type": "text", "text": result}


@mcp.tool
def search_email_by_sender_name_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
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
    return {"type": "text", "text": result}


@mcp.tool
def search_email_by_recipient_name_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
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
    return {"type": "text", "text": result}


@mcp.tool
def search_email_by_body_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
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
    return {"type": "text", "text": result}


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
    return {"type": "text", "text": result}


@mcp.tool
def load_emails_by_folder_tool(folder_path: str, days: int = 7) -> dict:
    """Load emails from a specific folder into cache.

    Args:
        folder_path: Path to the folder (supports nested paths like "Inbox/SubFolder1")
        days: Number of days to look back (default: 7, max: 30)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X emails from last Y days. Use 'view_email_cache_tool' to view them."
        }

    Note:
        - Maximum 30 emails for Inbox folder
        - Maximum 1000 emails for other folders
        - Supports 2nd level folder paths (e.g., "Inbox/SubFolder1")
    """
    if not folder_path or not isinstance(folder_path, str):
        raise ValueError("Folder path must be a non-empty string")
    if not isinstance(days, int) or days < 1 or days > 30:
        raise ValueError("Days parameter must be an integer between 1 and 30")

    # Check if it's Inbox folder (case-insensitive)
    if folder_path.lower() == "inbox":
        # For Inbox, we'll use the existing get_emails_from_folder but limit to 30
        emails, note = get_emails_from_folder(folder_name=folder_path, days=days)

        # If we have more than 30 emails, limit to 30
        if len(emails) > 30:
            # Clear cache and reload with limit
            clear_email_cache()
            for email in emails[:30]:
                add_email_to_cache(email["id"], email)
            save_email_cache()
            note = " (limited to 30 emails)"

        result = f"Found {len(emails)} emails from last {days} days. Use 'view_email_cache_tool' to view them.{note}"
    else:
        # For non-Inbox folders, use normal loading
        result = list_recent_emails(folder_name=folder_path, days=days)

    return {"type": "text", "text": result}


@mcp.tool
def get_email_by_number_tool(email_number: int, mode: str = "basic", include_attachments: bool = True, embed_images: bool = True) -> dict:
    """
    Get email content by its cache number with configurable retrieval modes.
    
    This unified tool replaces the previous basic and enhanced email retrieval tools
    with a single interface that supports different retrieval modes:
    - "basic": Fast, lightweight retrieval for email listings and summaries
    - "enhanced": Full media support with attachments, inline images, and comprehensive metadata
    - "lazy": Intelligent mode that adapts based on cached data for optimal performance
    
    Requires emails to be loaded first via list_recent_emails or search_emails.

    Args:
        email_number: The number of the email in the cache (1-based)
        mode: Retrieval mode - "basic" (original), "enhanced" (with media), "lazy" (performance optimized)
        include_attachments: Whether to include attachment content (enhanced mode only, default: True)
        embed_images: Whether to embed inline images in HTML body (enhanced mode only, default: True)

    Returns:
        dict: Response containing email details based on requested mode
        {
            "type": "text",
            "text": "Email details here"
        }

    Raises:
        ValueError: If email number is invalid or no emails are loaded
        RuntimeError: If cache contains invalid data
    """
    try:
        result = get_email_by_number_unified(email_number, mode, include_attachments, embed_images)
        if result is None:
            raise ValueError(
                "No email found at that position. Please load emails first using list_recent_emails or search_emails."
            )
        
        # Format the result based on mode
        if mode == "enhanced" and result.get("attachments"):
            formatted_text = format_email_with_media(result)
        else:
            # For basic and lazy modes, use simple formatting
            formatted_text = f"Subject: {result.get('subject', 'N/A')}\n"
            formatted_text += f"From: {result.get('from', 'N/A')}\n"
            formatted_text += f"To: {result.get('to', 'N/A')}\n"
            formatted_text += f"Date: {result.get('received', 'N/A')}\n"
            formatted_text += f"Body: {result.get('body', 'N/A')}\n"
            if result.get("attachments"):
                formatted_text += f"Attachments: {len(result['attachments'])}\n"
        
        return {"type": "text", "text": formatted_text}
        
    except ValueError as e:
        return {"type": "text", "text": f"Error: {str(e)}"}


@mcp.tool
def reply_to_email_by_number_tool(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[Union[str, List[str]]] = None,
    cc_recipients: Optional[Union[str, List[str]]] = None,
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
    return {"type": "text", "text": result}


@mcp.tool
def compose_email_tool(
    recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None
) -> dict:
    """
    IMPORTANT: You MUST get explicit user confirmation before calling this tool.
    Never send an email without the user's direct approval.

    Compose and send a new email

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

    # Parse multiple recipients from semicolon-separated string
    to_recipients = [email.strip() for email in recipient_email.split(";") if email.strip()]

    # Parse multiple CC recipients from semicolon-separated string
    cc_recipients = []
    if cc_email:
        cc_recipients = [email.strip() for email in cc_email.split(";") if email.strip()]

    result = compose_email(to_recipients, subject, body, cc_recipients if cc_recipients else None)
    return {"type": "text", "text": result}


@mcp.tool
def batch_forward_email_tool(email_number: int, csv_path: str, custom_text: str = "") -> dict:
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
    return {"type": "text", "text": result}


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
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error creating folder: {str(e)}"}


@mcp.tool
def remove_folder_tool(folder_name: str) -> dict:
    """Remove an existing folder.

    Args:
        folder_name: Name or path of the folder to remove (supports nested paths like "Inbox/SubFolder1/SubFolder2")

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
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error removing folder: {str(e)}"}


@mcp.tool
def move_email_tool(email_number: int, target_folder_name: str) -> dict:
    """Move an email to the specified folder.

    Args:
        email_number: The number of the email in the cache to move (1-based)
        target_folder_name: Name or path of the target folder (supports nested paths like "Inbox/SubFolder1/SubFolder2")

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
        # Convert email number to actual email ID from cache
        if not shared.email_cache:
            raise ValueError("No emails in cache - load emails first")
        if email_number < 1 or email_number > len(shared.email_cache):
            raise ValueError("Invalid email number")

        email_id = list(shared.email_cache.keys())[email_number - 1]

        with OutlookSessionManager() as outlook_session:
            result = outlook_session.move_email(email_id, target_folder_name)
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error moving email: {str(e)}"}


@mcp.tool
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
        # Convert email number to actual email ID from cache
        if not shared.email_cache:
            raise ValueError("No emails in cache - load emails first")
        if email_number < 1 or email_number > len(shared.email_cache):
            raise ValueError("Invalid email number")

        email_id = list(shared.email_cache.keys())[email_number - 1]

        with OutlookSessionManager() as outlook_session:
            result = outlook_session.delete_email(email_id)
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error moving email to Deleted Items: {str(e)}"}


@mcp.tool
def get_email_policies_tool(email_number: int) -> dict:
    """Get Exchange retention policies assigned to an email.

    Args:
        email_number: The number of the email in the cache to check (1-based)

    Returns:
        dict: Response containing assigned Exchange policies
        {
            "type": "text",
            "text": "Exchange Retention Policies:\n- Policy1\n- Policy2\n\nor\nNo policies assigned"
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
                "text": f"No email found at position {email_number}. Please load emails first using list_recent_emails or search_emails.",
            }

        # Get the email ID (EntryID) from the cached data
        email_id = email_data.get("id", "")
        if not email_id:
            return {
                "type": "text",
                "text": "Email data is missing the required ID for policy retrieval.",
            }

        # Now get the policies using the email ID
        with OutlookSessionManager() as outlook_session:
            policy_data = outlook_session.get_email_policies(email_id)

            # Build response for Exchange retention policies
            policies = policy_data.get("policies", [])

            if policies:
                result = "Exchange Retention Policies:\n"
                result += "\n".join([f"- {policy}" for policy in policies])
            else:
                result = "No Exchange retention policies assigned to this email."

            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error getting email policies: {str(e)}"}


@mcp.tool
def get_policies_tool() -> dict:
    """Get available enterprise policies that can be assigned to emails.

    Returns:
        dict: Response containing structured policy data for LLM consumption
        {
            "type": "text",
            "text": "Structured policy data for programmatic consumption"
        }
    """
    try:
        with OutlookSessionManager() as outlook_session:
            available_policies = outlook_session.get_available_policies()

            # Build structured policy data using only what Outlook provides
            structured_policies = {}

            # Use a set to track processed policies and eliminate duplicates
            processed_policies = set()

            # Only show retention policies (exclude Outlook sensitivity policies)
            for policy in available_policies:
                # Skip Outlook sensitivity policies as requested
                if policy in ["Normal", "Personal", "Private", "Confidential"]:
                    continue

                # Handle IRM policies
                if policy.startswith("IRM Policy: "):
                    policy_name = policy.replace("IRM Policy: ", "")
                    if policy_name not in processed_policies:
                        structured_policies[policy_name] = {
                            "name": policy_name,
                            "description": "IRM policy for document protection and rights management",
                            "value": policy_name,  # Add value field for assignment
                            "category": "IRM Policy",
                        }
                        processed_policies.add(policy_name)

                # Handle retention policies (Exchange or custom)
                elif "year" in policy.lower():
                    if policy.endswith(" (Enterprise)"):
                        # Enterprise policy
                        clean_policy_name = policy.replace(" (Enterprise)", "").strip()
                        structured_policies[clean_policy_name] = {
                            "name": policy,  # Keep full name including Enterprise suffix
                            "description": f"Custom enterprise retention policy for {clean_policy_name}",
                            "value": policy,  # Use original policy name as value for assignment
                            "category": "Enterprise Retention Policy",
                        }
                        processed_policies.add(clean_policy_name)
                    else:
                        # Exchange retention policy
                        structured_policies[policy] = {
                            "name": policy,
                            "description": f"Exchange server retention policy for {policy}",
                            "value": policy,  # Use original policy name as value for assignment
                            "category": "Exchange Retention Policy",
                        }
                        processed_policies.add(policy)

                # Handle other custom policies
                else:
                    if policy not in processed_policies:
                        structured_policies[policy] = {
                            "name": policy,
                            "description": "Custom policy configured in your environment",
                            "value": policy,  # Add value field for assignment
                            "category": "Custom Policy",
                        }
                        processed_policies.add(policy)

            # Return structured data with human-readable format for backward compatibility
            formatted_text = "📋 Available Policies Retrieved from Outlook:\n\n"

            # Group policies by category
            categories = {}
            for policy_name, policy_info in structured_policies.items():
                category = policy_info["category"]
                if category not in categories:
                    categories[category] = []
                categories[category].append((policy_name, policy_info))

            # Display by category - use proper icons that match our categories
            category_icons = {
                "Exchange Retention Policy": "📧",
                "Enterprise Retention Policy": "🏢",
                "IRM Policy": "🔒",
                "Custom Policy": "🔧",
            }

            for category, policies in categories.items():
                icon = category_icons.get(category, "📄")
                formatted_text += f"{icon} {category.upper()}:\n"
                for policy_name, policy_info in policies:
                    formatted_text += f"• {policy_name}: {policy_info['description']}\n"
                formatted_text += "\n"

            # Add structured data in the text response for LLM consumption
            formatted_text += "🔧 STRUCTURED DATA FOR PROGRAMMATIC USE:\n"
            formatted_text += "```json\n"
            formatted_text += "{\n"
            formatted_text += '  "policies": {\n'

            for i, (policy_name, policy_info) in enumerate(structured_policies.items()):
                comma = "," if i < len(structured_policies) - 1 else ""
                formatted_text += f'    "{policy_name}": {{\n'
                formatted_text += f'      "name": "{policy_info["name"]}",\n'
                formatted_text += f'      "description": "{policy_info["description"]}",\n'
                formatted_text += f'      "value": "{policy_info["value"]}",\n'
                formatted_text += f'      "category": "{policy_info["category"]}"\n'
                formatted_text += f"    }}{comma}\n"

            formatted_text += "  }\n"
            formatted_text += "}\n"
            formatted_text += "```\n"

            return {"type": "text", "text": formatted_text}
    except Exception as e:
        return {"type": "text", "text": f"Error getting available policies: {str(e)}"}


@mcp.tool
def assign_policy_tool(email_number: int, policy_name: str) -> dict:
    """Assign an enterprise policy to an email.

    Args:
        email_number: The number of the email in the cache to assign policy to (1-based)
        policy_name: Name of the policy to assign

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Successfully assigned policy to email"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    if not policy_name or not isinstance(policy_name, str):
        raise ValueError("Policy name must be a non-empty string")

    try:
        # First get the email from cache to get the email ID
        email_data = get_email_by_number(email_number)
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
                "text": "Email data is missing the required ID for policy assignment.",
            }

        # Now assign the policy using the email ID
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.assign_policy(email_id, policy_name)
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error assigning policy: {str(e)}"}


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
