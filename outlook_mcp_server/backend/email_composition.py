from typing import List, Optional, Union

from .outlook_session import OutlookSessionManager
from .shared import email_cache

def reply_to_email_by_number(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[Union[str, List[str]]] = None,
    cc_recipients: Optional[Union[str, List[str]]] = None
) -> str:
    """Reply to an email with custom recipients if provided
    
    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
        cc_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
    """
    # Input validation
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    
    if not reply_text or not isinstance(reply_text, str):
        raise ValueError("Reply text must be a non-empty string")
    
    # Handle string to list conversion for recipient parameters
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
    
    # Validate recipient lists
    if to_recipients is not None:
        if not all(isinstance(email, str) and email.strip() for email in to_recipients):
            raise ValueError("All To email addresses must be non-empty strings")
    
    if cc_recipients is not None:
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
                
            # Create a new email message to have full control over formatting
            new_mail = session.outlook.CreateItem(0)  # 0 = olMailItem
            
            # Determine recipients based on parameters
            if to_recipients is None and cc_recipients is None:
                # ReplyAll behavior - get all original recipients
                # Use the original email's sender as the To recipient
                new_mail.To = getattr(email, 'SenderEmailAddress', 'unknown@example.com')
                # Get CC recipients from original email if any
                original_cc = getattr(email, 'CC', '')
                if original_cc:
                    new_mail.CC = original_cc
            else:
                # Use custom recipients
                if to_recipients:
                    new_mail.To = "; ".join(to_recipients)
                if cc_recipients:
                    new_mail.CC = "; ".join(cc_recipients)

            # Set subject with RE: prefix
            new_mail.Subject = f"RE: {getattr(email, 'Subject', 'No Subject')}"

            # Build the email body with proper formatting and encoding
            # First, get the reply text with proper encoding handling
            try:
                # Ensure reply text is properly encoded
                if isinstance(reply_text, bytes):
                    reply_text = reply_text.decode('utf-8', errors='replace')
                elif not isinstance(reply_text, str):
                    reply_text = str(reply_text)
                
                # Build body content with explicit encoding control
                body_lines = [
                    reply_text,
                    "",
                    "_" * 50,
                    f"From: {getattr(email, 'SenderName', 'Unknown Sender')}",
                    f"Sent: {str(getattr(email, 'SentOn', 'Unknown'))}",
                    f"To: {getattr(email, 'To', 'Unknown')}"
                ]
                
                # Add CC if present
                original_cc = getattr(email, 'CC', '')
                if original_cc and original_cc.strip():
                    body_lines.append(f"Cc: {original_cc}")
                
                body_lines.extend([
                    f"Subject: {getattr(email, 'Subject', 'No Subject')}",
                    ""
                ])
                
                # Add the original email content with robust encoding handling
                original_body = getattr(email, 'Body', '')
                try:
                    if isinstance(original_body, bytes):
                        # Try multiple encodings for robustness
                        for encoding in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1', 'cp1252']:
                            try:
                                original_body = original_body.decode(encoding)
                                break
                            except (UnicodeDecodeError, LookupError):
                                continue
                        else:
                            # If all encodings fail, use error replacement
                            original_body = original_body.decode('utf-8', errors='replace')
                    elif not isinstance(original_body, str):
                        original_body = str(original_body)
                    
                    # Additional safety: ensure the content is safe for email
                    safe_body = original_body.encode('ascii', errors='replace').decode('ascii')
                    body_lines.append(safe_body)
                    
                except Exception as content_error:
                    print(f"WARNING: Could not process original email content: {content_error}")
                    body_lines.append("[Original email content unavailable due to formatting issues]")
                
                # Join with proper line endings
                body_content = "\n".join(body_lines)
                
                # Debug print to see what body content is being built
                print("DEBUG: Built body content:")
                print(body_content)
                print("DEBUG: End of body content")
                
            except (UnicodeDecodeError, UnicodeEncodeError) as e:
                print(f"WARNING: Encoding error in reply content: {e}")
                # Fallback: use ASCII-safe encoding for reply only
                if isinstance(reply_text, bytes):
                    reply_text = reply_text.decode('utf-8', errors='replace')
                else:
                    reply_text = str(reply_text).encode('ascii', errors='replace').decode('ascii')
                
                # Skip original content if encoding issues persist
                body_content = f"{reply_text}\n\n{'_' * 50}\n[Original email content omitted due to encoding compatibility]"
            except Exception as e:
                print(f"WARNING: Error building email body: {e}")
                # More user-friendly error handling
                try:
                    safe_reply = str(reply_text).encode('ascii', errors='replace').decode('ascii')
                    body_content = f"{safe_reply}\n\n{'_' * 50}\n[Original email content unavailable]"
                except:
                    # Ultimate fallback
                    body_content = "Thank you for your message. Unable to include original content due to technical limitations."
            
            # Set the body of the new email with proper encoding
            try:
                new_mail.Body = body_content
            except (UnicodeEncodeError, AttributeError) as e:
                print(f"WARNING: Failed to set email body with encoding: {e}")
                # Enhanced fallback: create ASCII-safe version
                try:
                    # Try to extract just the reply text
                    if isinstance(reply_text, bytes):
                        reply_text = reply_text.decode('utf-8', errors='replace')
                    elif not isinstance(reply_text, str):
                        reply_text = str(reply_text)
                    
                    ascii_safe_reply = reply_text.encode('ascii', errors='replace').decode('ascii')
                    simple_body = f"{ascii_safe_reply}\n\n{'_' * 50}\n[Original email content unavailable]"
                    new_mail.Body = simple_body
                except Exception as fallback_error:
                    print(f"ERROR: Ultimate fallback failed: {fallback_error}")
                    # Last resort: minimal content
                    new_mail.Body = "Thank you for your message."
            except Exception as e:
                print(f"ERROR: Failed to set email body: {e}")
                # Emergency fallback
                try:
                    new_mail.Body = "Reply sent. Original content unavailable due to technical limitations."
                except:
                    # If all else fails, send without body
                    print("CRITICAL: Could not set email body at all")
                    new_mail.Body = "Reply sent."
                
            new_mail.Send()
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
            # Ensure proper encoding for all email components
            try:
                # Encode recipients
                encoded_to = []
                for recipient in to_recipients:
                    if isinstance(recipient, bytes):
                        recipient = recipient.decode('utf-8', errors='replace')
                    elif not isinstance(recipient, str):
                        recipient = str(recipient)
                    encoded_to.append(recipient.strip())
                
                # Encode subject
                if isinstance(subject, bytes):
                    subject = subject.decode('utf-8', errors='replace')
                elif not isinstance(subject, str):
                    subject = str(subject)
                
                # Encode body content
                if isinstance(body, bytes):
                    body = body.decode('utf-8', errors='replace')
                elif not isinstance(body, str):
                    body = str(body)
                
                # Encode CC recipients if present
                encoded_cc = []
                if cc_recipients:
                    for recipient in cc_recipients:
                        if isinstance(recipient, bytes):
                            recipient = recipient.decode('utf-8', errors='replace')
                        elif not isinstance(recipient, str):
                            recipient = str(recipient)
                        encoded_cc.append(recipient.strip())
                
            except UnicodeDecodeError as e:
                print(f"WARNING: Encoding error in compose email content: {e}")
                # Fallback: use ASCII-safe encoding
                subject = subject.encode('ascii', errors='replace').decode('ascii')
                body = body.encode('ascii', errors='replace').decode('ascii')
                encoded_to = [r.encode('ascii', errors='replace').decode('ascii') for r in to_recipients]
                if cc_recipients:
                    encoded_cc = [r.encode('ascii', errors='replace').decode('ascii') for r in cc_recipients]
            
            # Create and send the email with encoded content
            mail = session.outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = "; ".join(encoded_to)
            mail.Subject = subject
            
            if cc_recipients:
                mail.CC = "; ".join(encoded_cc)
                
            try:
                if html:
                    mail.HTMLBody = body
                else:
                    mail.Body = body
            except (UnicodeEncodeError, AttributeError) as body_error:
                print(f"WARNING: Failed to set email body with encoding: {body_error}")
                # Enhanced fallback to ASCII-safe body
                try:
                    safe_body = body.encode('ascii', errors='replace').decode('ascii')
                    if html:
                        mail.HTMLBody = f"<pre>{safe_body}</pre>"
                    else:
                        mail.Body = safe_body
                except Exception as fallback_error:
                    print(f"ERROR: Fallback body setting failed: {fallback_error}")
                    # Emergency minimal content
                    if html:
                        mail.HTMLBody = "<p>Email sent. Content unavailable due to technical limitations.</p>"
                    else:
                        mail.Body = "Email sent. Content unavailable due to technical limitations."
            except Exception as body_error:
                print(f"ERROR: Failed to set email body: {body_error}")
                # Emergency fallback
                if html:
                    mail.HTMLBody = "<p>Email sent.</p>"
                else:
                    mail.Body = "Email sent."
                
            mail.Send()
            return "Email sent successfully"
            
        except Exception as e:
            return f"Error composing email: {str(e)}"