from typing import List, Optional

from backend.outlook_session import OutlookSessionManager
from backend.shared import email_cache

def reply_to_email_by_number(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[List[str]] = None,
    cc_recipients: Optional[List[str]] = None
) -> str:
    """Reply to an email with custom recipients if provided"""
    if not email_cache:
        return "No emails available - please list emails first."
        
    cache_items = list(email_cache.values())
    if not 1 <= email_number <= len(cache_items):
        return f"Email #{email_number} not found in current listing."
    
    email_id = cache_items[email_number - 1]["id"]
    
    with OutlookSessionManager() as session:
        try:
            email = session.namespace.GetItemFromID(email_id)
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