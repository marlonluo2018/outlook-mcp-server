"""Email composition and reply functions with improved encoding handling"""
from typing import List, Optional, Union
import logging

from .outlook_session import OutlookSessionManager
from .shared import email_cache
from .utils import safe_encode_text, normalize_email_address
from .validators import EmailReplyParams, EmailComposeParams

logger = logging.getLogger(__name__)


def reply_to_email_by_number(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[Union[str, List[str]]] = None,
    cc_recipients: Optional[Union[str, List[str]]] = None
) -> str:
    """
    Reply to an email with custom recipients if provided.
    
    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
        cc_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
        
    Returns:
        str: Success or error message
    """
    # Validate inputs using Pydantic
    try:
        params = EmailReplyParams(
            email_number=email_number,
            reply_text=reply_text,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients
        )
    except Exception as e:
        logger.error(f"Validation error in reply_to_email_by_number: {e}")
        raise ValueError(f"Invalid parameters: {e}")
    
    # Convert to list if needed (validator already did this)
    to_recipients = params.to_recipients
    cc_recipients = params.cc_recipients
    reply_text = params.reply_text
    
    if not email_cache:
        raise ValueError("No emails available - please list emails first.")
        
    cache_items = list(email_cache.values())
    if not 1 <= email_number <= len(cache_items):
        raise ValueError(f"Email #{email_number} not found in current listing.")
    
    cached_email = cache_items[email_number - 1]
    
    with OutlookSessionManager() as session:
        try:
            email = session.namespace.GetItemFromID(cached_email["id"])
            if not email:
                raise RuntimeError("Could not retrieve the email from Outlook.")
            
            # Create a new email message to have full control over formatting
            new_mail = session.outlook.CreateItem(0)  # 0 = olMailItem
            
            # Determine recipients based on parameters
            if to_recipients is None and cc_recipients is None:
                # ReplyAll behavior - get all original recipients
                sender_email = safe_encode_text(getattr(email, 'SenderEmailAddress', 'unknown@example.com'), 'to_address')
                new_mail.To = sender_email
                
                # Normalize sender email for comparison
                normalized_sender_email = normalize_email_address(sender_email)
                
                # Use cached recipient data to avoid Outlook name resolution issues
                unique_recipients = set()
                
                # Get TO recipients from cache using both display names and email addresses
                to_recipients_data = cached_email.get('to_recipients', [])
                for recipient_info in to_recipients_data:
                    if isinstance(recipient_info, dict):
                        recipient_email = recipient_info.get('email', '').strip()
                        recipient_display_name = recipient_info.get('display_name', '').strip()
                        normalized_recipient_email = normalize_email_address(recipient_email)
                        
                        if recipient_email and normalized_recipient_email != normalized_sender_email:
                            # Prefer display name with email, fallback to just email
                            if recipient_display_name:
                                recipient_string = f"{recipient_display_name} <{recipient_email}>"
                            else:
                                recipient_string = recipient_email
                            unique_recipients.add(recipient_string)
                
                # Get CC recipients from cache using both display names and email addresses
                cc_recipients_data = cached_email.get('cc_recipients', [])
                for recipient_info in cc_recipients_data:
                    if isinstance(recipient_info, dict):
                        recipient_email = recipient_info.get('email', '').strip()
                        recipient_display_name = recipient_info.get('display_name', '').strip()
                        normalized_recipient_email = normalize_email_address(recipient_email)
                        
                        if recipient_email and normalized_recipient_email != normalized_sender_email:
                            # Prefer display name with email, fallback to just email
                            if recipient_display_name:
                                recipient_string = f"{recipient_display_name} <{recipient_email}>"
                            else:
                                recipient_string = recipient_email
                            unique_recipients.add(recipient_string)
                
                # Set CC field with all unique recipients if any
                if unique_recipients:
                    new_mail.CC = '; '.join(sorted(unique_recipients))
            else:
                # Use custom recipients
                if to_recipients is not None:
                    new_mail.To = "; ".join(to_recipients)
                if cc_recipients is not None:
                    new_mail.CC = "; ".join(cc_recipients)

            # Set subject with RE: prefix
            subject = safe_encode_text(getattr(email, 'Subject', 'No Subject'), 'subject')
            new_mail.Subject = f"RE: {subject}"

            # Build the email body with proper formatting and encoding
            reply_text_safe = safe_encode_text(reply_text, 'reply_text')
            sender_name = safe_encode_text(getattr(email, 'SenderName', 'Unknown Sender'), 'sender_name')
            sent_on = safe_encode_text(str(getattr(email, 'SentOn', 'Unknown')), 'sent_on')
            to_field = safe_encode_text(getattr(email, 'To', 'Unknown'), 'to_field')
            
            # Build body content
            body_lines = [
                reply_text_safe,
                "",
                "_" * 50,
                f"From: {sender_name}",
                f"Sent: {sent_on}",
                f"To: {to_field}"
            ]
            
            # Add CC if present
            original_cc = safe_encode_text(getattr(email, 'CC', ''), 'original_cc')
            if original_cc and original_cc.strip():
                body_lines.append(f"Cc: {original_cc}")
            
            body_lines.extend([
                f"Subject: {subject}",
                ""
            ])
            
            # Add the original email content
            original_body = safe_encode_text(getattr(email, 'Body', ''), 'original_body')
            body_lines.append(original_body)
            
            # Join with proper line endings
            body_content = "\n".join(body_lines)
            
            # Set the body of the new email
            try:
                new_mail.Body = body_content
            except Exception as e:
                logger.warning(f"Failed to set email body, using simplified version: {e}")
                # Fallback to simple body
                new_mail.Body = f"{reply_text_safe}\n\n{'_' * 50}\n[Original email content unavailable]"
                
            new_mail.Send()
            logger.info(f"Successfully replied to email #{email_number}")
            return f"Successfully replied to email #{email_number}"
            
        except Exception as e:
            logger.error(f"Error replying to email #{email_number}: {e}")
            return f"Error replying to email: {str(e)}"


def compose_email(
    to_recipients: List[str],
    subject: str,
    body: str,
    cc_recipients: Optional[List[str]] = None,
    html: bool = False
) -> str:
    """
    Compose and send a new email using Outlook COM API.
    
    Args:
        to_recipients: List of recipient email addresses
        subject: Email subject line
        body: Email body content
        cc_recipients: Optional list of CC email addresses
        html: If True, body is treated as HTML (default: False)
        
    Returns:
        str: Success/error message
    """
    # Validate inputs using Pydantic
    try:
        params = EmailComposeParams(
            recipient_email=to_recipients[0] if to_recipients else "",
            subject=subject,
            body=body,
            cc_email=cc_recipients[0] if cc_recipients else None
        )
    except Exception as e:
        logger.error(f"Validation error in compose_email: {e}")
        raise ValueError(f"Invalid parameters: {e}")
    
    # Additional validation for list
    if not to_recipients or not isinstance(to_recipients, list):
        raise ValueError("To recipients must be a non-empty list")
    
    if not all(isinstance(email, str) and email.strip() for email in to_recipients):
        raise ValueError("All recipient email addresses must be non-empty strings")
    
    if cc_recipients is not None:
        if not isinstance(cc_recipients, list):
            raise ValueError("CC recipients must be a list or None")
        if not all(isinstance(email, str) and email.strip() for email in cc_recipients):
            raise ValueError("All CC email addresses must be non-empty strings")
    
    with OutlookSessionManager() as session:
        try:
            # Encode all components safely
            encoded_to = [safe_encode_text(recipient, 'to_recipient').strip() 
                         for recipient in to_recipients]
            subject_safe = safe_encode_text(subject, 'subject')
            body_safe = safe_encode_text(body, 'body')
            
            encoded_cc = []
            if cc_recipients:
                encoded_cc = [safe_encode_text(recipient, 'cc_recipient').strip() 
                             for recipient in cc_recipients]
            
            # Create and send the email
            mail = session.outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = "; ".join(encoded_to)
            mail.Subject = subject_safe
            
            if cc_recipients:
                mail.CC = "; ".join(encoded_cc)
                
            try:
                if html:
                    mail.HTMLBody = body_safe
                else:
                    mail.Body = body_safe
            except Exception as e:
                logger.warning(f"Failed to set email body format, using plain text: {e}")
                mail.Body = body_safe
                
            mail.Send()
            logger.info(f"Email sent successfully to {len(to_recipients)} recipients")
            return "Email sent successfully"
            
        except Exception as e:
            logger.error(f"Error composing email: {e}")
            return f"Error composing email: {str(e)}"