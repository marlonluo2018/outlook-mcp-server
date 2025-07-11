import csv

from backend.outlook_session import OutlookSessionManager
from backend.shared import email_cache

def send_batch_emails(
    email_number: int,
    csv_path: str,
    custom_text: str = ""
) -> str:
    """Send email to recipients in batches of 500 (Outlook BCC limit)"""
    if not email_cache:
        return "No emails available - please list emails first."

    cache_items = list(email_cache.values())
    if not 1 <= email_number <= len(cache_items):
        return f"Email #{email_number} not found in current listing."

    try:
        # Clean and validate CSV path
        clean_path = csv_path.strip('"\'')
        
        # Read recipients from CSV
        with open(clean_path, 'r', newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)
            if 'email' not in reader.fieldnames:
                return "CSV must contain an 'email' column"
                
            recipients = [row['email'] for row in reader if row.get('email')]
            
        if not recipients:
            return "No valid email addresses found in CSV"

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
                    mail = template.Forward()
                    mail.BCC = "; ".join(batch)
                    
                    if custom_text:
                        if hasattr(mail, 'HTMLBody') and mail.HTMLBody:
                            mail.HTMLBody = f"<div>{custom_text}</div>" + mail.HTMLBody
                        else:
                            mail.Body = custom_text + "\n\n" + mail.Body
                    
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