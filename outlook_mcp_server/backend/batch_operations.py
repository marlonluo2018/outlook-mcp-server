import csv
import logging

from .outlook_session.session_manager import OutlookSessionManager
from .shared import email_cache
from .utils import safe_encode_text, validate_email_address

logger = logging.getLogger(__name__)


def batch_forward_emails(email_number: int, csv_path: str, custom_text: str = "") -> str:
    """Forward email to recipients in batches of 500 (Outlook BCC limit)"""
    # Input validation
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")

    if not csv_path or not isinstance(csv_path, str):
        raise ValueError("CSV path must be a non-empty string")

    if not isinstance(custom_text, str):
        raise ValueError("Custom text must be a string")

    if not email_cache:
        raise ValueError("No emails available - please list emails first.")

    cache_items = list(email_cache.values())
    if not 1 <= email_number <= len(cache_items):
        raise ValueError(f"Email #{email_number} not found in current listing.")

    try:
        # Clean and validate CSV path
        clean_path = csv_path.strip("\"'")

        # Read recipients from CSV
        with open(clean_path, "r", newline="", encoding="utf-8-sig") as csvfile:
            reader = csv.DictReader(csvfile)
            if "email" not in reader.fieldnames:
                raise ValueError("CSV must contain an 'email' column")

            recipients = []
            invalid_emails = []

            for row in reader:
                email = row.get("email", "").strip()
                if email:
                    if validate_email_address(email):
                        recipients.append(email)
                    else:
                        invalid_emails.append(email)
                        logger.warning(f"Invalid email address found: {email}")

        if invalid_emails:
            raise ValueError(
                f"Invalid email addresses found: {', '.join(invalid_emails[:5])}{'...' if len(invalid_emails) > 5 else ''}"
            )

        if not recipients:
            raise ValueError("No valid email addresses found in CSV")

        # Process in batches of 500 (Outlook BCC limit)
        batch_size = 500
        batches = [recipients[i : i + batch_size] for i in range(0, len(recipients), batch_size)]
        total_recipients = len(recipients)
        results = []

        with OutlookSessionManager() as session:
            # Get email data from cache - use entry_id instead of id
            email_data = cache_items[email_number - 1]
            email_id = email_data.get("entry_id") or email_data.get("id")
            if not email_id:
                raise ValueError(f"Email #{email_number} does not have a valid ID field")
            template = session.namespace.GetItemFromID(email_id)

            for i, batch in enumerate(batches, 1):
                try:
                    # Create a regular mail item instead of using Forward()
                    mail = session.outlook.CreateItem(0)  # 0 = olMailItem

                    # Copy relevant properties from template with encoding handling
                    try:
                        # Handle subject encoding using safe utility
                        subject = safe_encode_text(template.Subject, "batch_subject")
                        mail.Subject = f"FW: {subject}"
                    except Exception as e:
                        logger.error(f"Encoding error in batch subject: {e}")
                        mail.Subject = "FW: [Subject encoding error]"

                    mail.BCC = "; ".join(batch)

                    # Copy body content from template with proper encoding and email headers
                    try:
                        # Extract email metadata for headers
                        sender_name = safe_encode_text(
                            getattr(template, "SenderName", "Unknown Sender"), "sender_name"
                        )
                        sent_on = safe_encode_text(
                            str(getattr(template, "SentOn", "Unknown")), "sent_on"
                        )
                        to_field = safe_encode_text(getattr(template, "To", "Unknown"), "to_field")
                        subject = safe_encode_text(
                            getattr(template, "Subject", "No Subject"), "subject"
                        )

                        if hasattr(template, "HTMLBody") and template.HTMLBody:
                            mail.BodyFormat = 2  # 2 = olFormatHTML
                            html_body = safe_encode_text(template.HTMLBody, "batch_html_body")

                            # Build HTML email headers
                            header_html = f"""
<div>
{'' if not custom_text else f'<div>{safe_encode_text(custom_text, "batch_custom_text")}</div><br>'}
<div style="margin-bottom: 10px;">__________________________________________________</div>
<div><strong>From:</strong> {sender_name}</div>
<div><strong>Sent:</strong> {sent_on}</div>
<div><strong>To:</strong> {to_field}</div>
<div><strong>Subject:</strong> {subject}</div>
<div style="margin-top: 10px; margin-bottom: 10px;">__________________________________________________</div>
</div>
<br><br>"""

                            mail.HTMLBody = header_html + html_body
                        else:
                            mail.BodyFormat = 1  # 1 = olFormatPlain
                            plain_body = safe_encode_text(
                                getattr(template, "Body", ""), "batch_plain_body"
                            )

                            # Build plain text email headers
                            header_lines = []
                            if custom_text:
                                header_lines.append(
                                    safe_encode_text(custom_text, "batch_custom_text")
                                )
                            header_lines.extend(
                                [
                                    "",
                                    "_" * 50,
                                    f"From: {sender_name}",
                                    f"Sent: {sent_on}",
                                    f"To: {to_field}",
                                    f"Subject: {subject}",
                                    "_" * 50,
                                    "",
                                ]
                            )

                            mail.Body = "\n".join(header_lines) + plain_body
                    except Exception as e:
                        logger.error(f"Error processing batch body: {e}")
                        mail.Body = "[Content processing error - please view original email]"

                    mail.Send()
                    logger.info(f"Batch {i} sent to {len(batch)} recipients")
                    results.append(f"Batch {i} sent to {len(batch)} recipients")
                except Exception as e:
                    logger.error(f"Error sending batch {i}: {e}")
                    results.append(f"Error sending batch {i}: {str(e)}")

        return "\n".join(
            [
                f"Batch sending completed for {total_recipients} recipients in {len(batches)} batches:",
                *results,
            ]
        )

    except Exception as e:
        logger.error(f"Error in batch sending process: {e}")
        return f"Error in batch sending process: {str(e)}"
