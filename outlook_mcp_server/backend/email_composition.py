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
    cc_recipients: Optional[Union[str, List[str]]] = None,
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
            cc_recipients=cc_recipients,
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

            # Extract sender email early for use in CC filtering
            sender_email = safe_encode_text(
                getattr(email, "SenderEmailAddress", "unknown@example.com"), "to_address"
            )
            normalized_sender_email = normalize_email_address(sender_email)

            # Additional sender extraction for robustness
            sender_name = getattr(email, "SenderName", "")
            sender_address = getattr(email, "SenderEmailAddress", "")
            
            # Log comprehensive sender information for debugging
            logger.debug(f"=== SENDER EXTRACTION DEBUG ===")
            logger.debug(f"SenderEmailAddress: {sender_email}")
            logger.debug(f"SenderName: {sender_name}")
            logger.debug(f"Combined sender info: {sender_name} <{sender_address}>")
            logger.debug(f"Normalized sender email: {normalized_sender_email}")
            logger.debug(f"=== END SENDER EXTRACTION DEBUG ===")
            
            # Also check if sender appears in original email fields
            original_to = safe_encode_text(getattr(email, "To", ""), "original_to")
            original_cc = safe_encode_text(getattr(email, "CC", ""), "original_cc")
            logger.debug(f"Original TO field: {original_to}")
            logger.debug(f"Original CC field: {original_cc}")
            
            # Create a comprehensive list of sender variations to filter against
            sender_variations = set()
            sender_variations.add(normalized_sender_email)
            
            # Add display name variations
            if sender_name and sender_address:
                # "Name <email@domain.com>" format
                display_format = f"{sender_name} <{sender_address}>".strip()
                sender_variations.add(normalize_email_address(display_format))
                
                # Also check individual components
                sender_variations.add(normalize_email_address(sender_name))
            
            # Check if sender appears in original To field
            if original_to:
                to_emails = [addr.strip() for addr in original_to.split(";") if addr.strip()]
                for to_email in to_emails:
                    normalized_to = normalize_email_address(to_email)
                    sender_variations.add(normalized_to)
                    if normalized_to == normalized_sender_email:
                        logger.debug(f"Found sender in original TO field: {to_email}")
            
            # Check if sender appears in original CC field
            if original_cc:
                cc_emails = [addr.strip() for addr in original_cc.split(";") if addr.strip()]
                for cc_email in cc_emails:
                    normalized_cc = normalize_email_address(cc_email)
                    sender_variations.add(normalized_cc)
                    if normalized_cc == normalized_sender_email:
                        logger.debug(f"Found sender in original CC field: {cc_email}")
            
            logger.debug(f"Sender variations to filter against: {sorted(sender_variations)}")
            
            # Create a comprehensive filtering function
            def is_sender_email(email_address):
                """Check if an email address matches any sender variation"""
                normalized = normalize_email_address(email_address)
                return normalized in sender_variations

            # Determine recipients based on parameters
            if to_recipients is None and cc_recipients is None:
                # ReplyAll behavior - get all original recipients
                new_mail.To = sender_email

                # Use cached recipient data to avoid Outlook name resolution issues
                cc_recipients_set = set()

                # Get CC recipients from cache using both display names and email addresses
                cc_recipients_data = cached_email.get("cc_recipients", [])
                logger.debug(f"Processing {len(cc_recipients_data)} CC recipients from cache")
                
                for i, recipient_info in enumerate(cc_recipients_data):
                    if isinstance(recipient_info, dict):
                        recipient_email = recipient_info.get("email", "").strip()
                        recipient_display_name = recipient_info.get("display_name", "").strip()
                        normalized_recipient_email = normalize_email_address(recipient_email)
                        
                        logger.debug(f"CC recipient {i+1}: {recipient_info}")
                        logger.debug(f"  Extracted email: '{recipient_email}'")
                        logger.debug(f"  Extracted display name: '{recipient_display_name}'")
                        logger.debug(f"  Normalized email: '{normalized_recipient_email}'")
                        logger.debug(f"  Sender normalized: '{normalized_sender_email}'")
                        logger.debug(f"  Is sender: {is_sender_email(recipient_email)}")

                        if recipient_email:
                            if not is_sender_email(recipient_email):
                                # Prefer display name with email, fallback to just email
                                if recipient_display_name:
                                    recipient_string = f"{recipient_display_name} <{recipient_email}>"
                                else:
                                    recipient_string = recipient_email
                                cc_recipients_set.add(recipient_string)
                                logger.debug(f"  -> ADDED to CC: {recipient_string}")
                            else:
                                logger.debug(f"  -> FILTERED OUT (matches sender): {recipient_email}")
                        else:
                            logger.debug(f"  -> SKIPPED (empty email)")
                    else:
                        logger.debug(f"CC recipient {i+1}: Non-dict format: {recipient_info}")
                        
                logger.debug(f"Total CC recipients after filtering: {len(cc_recipients_set)}")
                if cc_recipients_set:
                    logger.debug(f"CC recipients list: {sorted(cc_recipients_set)}")

                # Set CC field with filtered CC recipients if any
                if cc_recipients_set:
                    logger.debug(f"Setting CC to (ReplyAll): {sorted(cc_recipients_set)}")
                    new_mail.CC = "; ".join(sorted(cc_recipients_set))
                else:
                    # Explicitly clear CC field if no valid recipients remain
                    logger.debug("No CC recipients after filtering - clearing CC field")
                    new_mail.CC = ""
            else:
                # Use custom recipients, but ensure original sender is not in CC
                if to_recipients is not None:
                    new_mail.To = "; ".join(to_recipients)
                if cc_recipients is not None:
                    # Filter out the original sender from CC recipients
                    filtered_cc = []
                    for recipient in cc_recipients:
                        # Use comprehensive sender filtering
                        if not is_sender_email(recipient):
                            filtered_cc.append(recipient)
                            logger.debug(f"CC recipient kept: {recipient}")
                        else:
                            logger.info(f"Filtered out original sender from CC: {recipient}")

                    # Explicitly set CC field
                    if filtered_cc:
                        logger.debug(f"Setting CC to: {filtered_cc}")
                        new_mail.CC = "; ".join(filtered_cc)
                    else:
                        # Explicitly clear CC field if no valid recipients remain
                        logger.debug("No CC recipients after filtering - clearing CC field")
                        new_mail.CC = ""

            # Set subject with RE: prefix
            subject = safe_encode_text(getattr(email, "Subject", "No Subject"), "subject")
            new_mail.Subject = f"RE: {subject}"

            # Build the email body with proper formatting and encoding
            reply_text_safe = safe_encode_text(reply_text, "reply_text")
            sender_name = safe_encode_text(
                getattr(email, "SenderName", "Unknown Sender"), "sender_name"
            )
            sent_on = safe_encode_text(str(getattr(email, "SentOn", "Unknown")), "sent_on")
            to_field = safe_encode_text(getattr(email, "To", "Unknown"), "to_field")

            # Build body content
            body_lines = [
                reply_text_safe,
                "",
                "_" * 50,
                f"From: {sender_name}",
                f"Sent: {sent_on}",
                f"To: {to_field}",
            ]

            # Add CC if present
            original_cc = safe_encode_text(getattr(email, "CC", ""), "original_cc")
            if original_cc and original_cc.strip():
                body_lines.append(f"Cc: {original_cc}")

            body_lines.extend([f"Subject: {subject}", ""])

            # Add the original email content
            original_body = safe_encode_text(getattr(email, "Body", ""), "original_body")
            body_lines.append(original_body)

            # Join with proper line endings
            body_content = "\n".join(body_lines)

            # Set the body of the new email
            try:
                new_mail.Body = body_content
            except Exception as e:
                logger.warning(f"Failed to set email body, using simplified version: {e}")
                # Fallback to simple body
                new_mail.Body = (
                    f"{reply_text_safe}\n\n{'_' * 50}\n[Original email content unavailable]"
                )

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
    html: bool = False,
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
            cc_email=cc_recipients[0] if cc_recipients else None,
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
            encoded_to = [
                safe_encode_text(recipient, "to_recipient").strip() for recipient in to_recipients
            ]
            subject_safe = safe_encode_text(subject, "subject")
            body_safe = safe_encode_text(body, "body")

            encoded_cc = []
            if cc_recipients:
                encoded_cc = [
                    safe_encode_text(recipient, "cc_recipient").strip()
                    for recipient in cc_recipients
                ]

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
