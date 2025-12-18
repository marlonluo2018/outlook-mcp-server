import json
from outlook_mcp_server.backend.shared import email_cache, email_cache_order

print(f"Cache contains {len(email_cache)} emails")
print(f"Cache order contains {len(email_cache_order)} emails")

# Look for Kirk Abbott emails in cache
kirk_emails = []
for email_id in email_cache_order:
    email_data = email_cache.get(email_id)
    if email_data and 'Kirk Abbott' in email_data.get('sender', ''):
        kirk_emails.append(email_data)

print(f"\nFound {len(kirk_emails)} Kirk Abbott emails in cache")
for i, email in enumerate(kirk_emails[:3]):
    print(f"Email {i+1}:")
    print(f"  Subject: {email.get('subject', 'N/A')}")
    print(f"  Sender: {email.get('sender', 'N/A')}")
    print(f"  Has attachments: {email.get('has_attachments', 'N/A')}")
    print()