from typing import List, Optional

from .outlook_session import OutlookSessionManager
from .shared import email_cache

def reply_to_email_by_number(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[List[str]] = None,
    cc_recipients: Optional[List[str]] = None
) -> str:
    """Reply to an email with custom recipients if provided"""
    # Input validation
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    
    if not reply_text or not isinstance(reply_text, str):
        raise ValueError("Reply text must be a non-empty string")
    
    if to_recipients is not None:
        if not isinstance(to_recipients, list):
            raise ValueError("To recipients must be a list or None")
        if not all(isinstance(email, str) and email.strip() for email in to_recipients):
            raise ValueError("All To email addresses must be non-empty strings")
    
    if cc_recipients is not None:
        if not isinstance(cc_recipients, list):
            raise ValueError("CC recipients must be a list or None")
        if not all(isinstance(email, str) and email.strip() for email in cc_recipients):
            raise ValueError("All CC email addresses must be non-empty strings")
    
    if not email_cache:
        raise ValueError("No emails available - please list emails first.")
        
    cache_items = list(email_cache.values())
    if not 1 <= email_number <= len(cache_items):
        raise ValueError(f"Email #{email_number} not found in current listing.")
    
    email_id = cache_items[email_number - 1]["id"]
    
    with OutlookSessionManager() as session:
        try:
            email = session.namespace.GetItemFromID(email_id)
            if not email:
                raise RuntimeError("Could not retrieve the email from Outlook.")
                
            if to_recipients is None and cc_recipients is None:
                reply = email.ReplyAll()
            else:
                reply = email.Reply()
                if to_recipients:
                    reply.To = "; ".join(to_recipients)
                if cc_recipients:
                    reply.CC = "; ".join(cc_recipients)

            # Handle HTML formatting
            if hasattr(email, 'HTMLBody') and email.HTMLBody:
                formatted_reply = reply_text.replace('\n\n', '<br><br>').replace('\n', '<br>')
                reply.HTMLBody = f"<div>{formatted_reply}</div>" + email.HTMLBody
            else:
                reply.Body = reply_text + "\n\n" + email.Body
                
            reply.Send()
            return f"Successfully replied to email #{email_number}"
            
        except Exception as e:
            return f"Error replying to email: {str(e)}"

def compose_email(
    to_recipients: List[str],
    subject: str,
    body: str,
    cc_recipients: Optional[List[str]] = None,
    html: bool = False
) -> str:
    """Compose and send a new email using Outlook COM API
    
    Args:
        to_recipients: List of recipient email addresses
        subject: Email subject line
        body: Email body content
        cc_recipients: Optional list of CC email addresses
        html: If True, body is treated as HTML (default: False)
        
    Returns:
        str: Success/error message
    """
    # Input validation
    if not to_recipients or not isinstance(to_recipients, list):
        raise ValueError("To recipients must be a non-empty list")
    
    if not all(isinstance(email, str) and email.strip() for email in to_recipients):
        raise ValueError("All recipient email addresses must be non-empty strings")
    
    if not subject or not isinstance(subject, str):
        raise ValueError("Subject must be a non-empty string")
    
    if not body or not isinstance(body, str):
        raise ValueError("Body must be a non-empty string")
    
    if cc_recipients is not None:
        if not isinstance(cc_recipients, list):
            raise ValueError("CC recipients must be a list or None")
        if not all(isinstance(email, str) and email.strip() for email in cc_recipients):
            raise ValueError("All CC email addresses must be non-empty strings")
    
    with OutlookSessionManager() as session:
        try:
            mail = session.outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = "; ".join(to_recipients)
            mail.Subject = subject
            
            if cc_recipients:
                mail.CC = "; ".join(cc_recipients)
                
            if html:
                mail.HTMLBody = body
            else:
                mail.Body = body
                
            mail.Send()
            return "Email sent successfully"
            
        except Exception as e:
            return f"Error composing email: {str(e)}"