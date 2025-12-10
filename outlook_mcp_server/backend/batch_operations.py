import csv

from .outlook_session import OutlookSessionManager
from .shared import email_cache

def send_batch_emails(
    email_number: int,
    csv_path: str,
    custom_text: str = ""
) -> str:
    """Send email to recipients in batches of 500 (Outlook BCC limit)"""
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
        clean_path = csv_path.strip('"\'')
        
        # Validate email format
        import re
        email_pattern = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')
        
        # Read recipients from CSV
        with open(clean_path, 'r', newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)
            if 'email' not in reader.fieldnames:
                raise ValueError("CSV must contain an 'email' column")
                
            recipients = []
            invalid_emails = []
            
            for row in reader:
                email = row.get('email', '').strip()
                if email:
                    if email_pattern.match(email):
                        recipients.append(email)
                    else:
                        invalid_emails.append(email)
            
        if invalid_emails:
            raise ValueError(f"Invalid email addresses found: {', '.join(invalid_emails[:5])}{'...' if len(invalid_emails) > 5 else ''}")
            
        if not recipients:
            raise ValueError("No valid email addresses found in CSV")

        # Process in batches of 500 (Outlook BCC limit)
        batch_size = 500
        batches = [recipients[i:i + batch_size] 
                  for i in range(0, len(recipients), batch_size)]
        total_recipients = len(recipients)
        results = []

        with OutlookSessionManager() as session:
            email_id = cache_items[email_number - 1]["id"]
            template = session.namespace.GetItemFromID(email_id)
            
            for i, batch in enumerate(batches, 1):
                try:
                    # Create a regular mail item instead of using Forward()
                    mail = session.outlook.CreateItem(0)  # 0 = olMailItem
                    
                    # Copy relevant properties from template with encoding handling
                    try:
                        # Handle subject encoding
                        subject = template.Subject
                        if isinstance(subject, bytes):
                            subject = subject.decode('utf-8', errors='replace')
                        elif not isinstance(subject, str):
                            subject = str(subject)
                        
                        mail.Subject = f"FW: {subject}"
                    except Exception as e:
                        print(f"WARNING: Encoding error in batch subject: {e}")
                        mail.Subject = "FW: [Subject encoding error]"
                    
                    mail.BCC = "; ".join(batch)
                    
                    # Copy body content from template with proper encoding
                    try:
                        if hasattr(template, 'HTMLBody') and template.HTMLBody:
                            mail.BodyFormat = 2  # 2 = olFormatHTML
                            
                            # Handle HTML body encoding
                            html_body = template.HTMLBody
                            if isinstance(html_body, bytes):
                                html_body = html_body.decode('utf-8', errors='replace')
                            elif not isinstance(html_body, str):
                                html_body = str(html_body)
                            
                            # Handle custom text encoding
                            if custom_text:
                                if isinstance(custom_text, bytes):
                                    custom_text = custom_text.decode('utf-8', errors='replace')
                                elif not isinstance(custom_text, str):
                                    custom_text = str(custom_text)
                                
                                safe_custom = custom_text.encode('ascii', errors='replace').decode('ascii')
                                mail.HTMLBody = f"<div>{safe_custom}</div><br><br>" + html_body
                            else:
                                mail.HTMLBody = html_body
                        else:
                            mail.BodyFormat = 1  # 1 = olFormatPlain
                            
                            # Handle plain text body encoding
                            plain_body = template.Body
                            if isinstance(plain_body, bytes):
                                plain_body = plain_body.decode('utf-8', errors='replace')
                            elif not isinstance(plain_body, str):
                                plain_body = str(plain_body)
                            
                            if custom_text:
                                # Handle custom text encoding
                                if isinstance(custom_text, bytes):
                                    custom_text = custom_text.decode('utf-8', errors='replace')
                                elif not isinstance(custom_text, str):
                                    custom_text = str(custom_text)
                                
                                safe_custom = custom_text.encode('ascii', errors='replace').decode('ascii')
                                mail.Body = safe_custom + "\n\n-----Original Email-----\n\n" + plain_body
                            else:
                                mail.Body = plain_body
                    except UnicodeDecodeError as e:
                        print(f"WARNING: Encoding error in batch body content: {e}")
                        # Fallback to simple ASCII-safe content
                        if hasattr(template, 'HTMLBody') and template.HTMLBody:
                            mail.BodyFormat = 1  # Fallback to plain text
                            mail.Body = "[HTML content with encoding issues - please view original email]"
                        else:
                            mail.Body = "[Content encoding error - please view original email]"
                    except Exception as e:
                        print(f"WARNING: Error processing batch body: {e}")
                        mail.Body = "[Content processing error - please view original email]"
                    
                    mail.Send()
                    results.append(f"Batch {i} sent to {len(batch)} recipients")
                except Exception as e:
                    results.append(f"Error sending batch {i}: {str(e)}")

        return "\n".join([
            f"Batch sending completed for {total_recipients} recipients in {len(batches)} batches:",
            *results
        ])
        
    except Exception as e:
        return f"Error in batch sending process: {str(e)}"