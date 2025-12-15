#!/usr/bin/env python3
"""
Test script to simulate actual email retrieval and composition workflow
"""

import sys
import os

# Add the backend to path for imports
sys.path.append('outlook_mcp_server/backend')

def test_actual_workflow():
    print("=== Testing Actual Email Retrieval → Composition Workflow ===\n")
    
    # Simulate what the email_retrieval.py should produce
    print("1. Simulating enhanced email cache data:")
    
    # This simulates the new format from email_retrieval.py
    simulated_cache = {
        'id': 'test_email_123',
        'subject': 'Test Email',
        'sender': 'John Smith',
        'sender_email': 'john.smith@company.com',
        'to_recipients': [
            {
                'display_name': 'Jane Doe', 
                'email': 'jane.doe@company.com'
            },
            {
                'display_name': 'Bob Wilson',
                'email': 'bob.wilson@company.com'
            }
        ],
        'cc_recipients': [
            {
                'display_name': 'Alice Johnson',
                'email': 'alice.johnson@company.com'
            }
        ]
    }
    
    print("   Cached email data:")
    for key, value in simulated_cache.items():
        if key in ['to_recipients', 'cc_recipients']:
            print(f"   {key}: {value}")
        else:
            print(f"   {key}: {value}")
    
    print("\n2. Testing email composition logic:")
    
    # Simulate the enhanced composition logic from email_composition.py
    sender_email = 'john.smith@company.com'  # The sender replying
    
    # Normalize email function (simplified version)
    def normalize_email_address(email: str) -> str:
        if not email:
            return ""
        normalized = email.strip().rstrip(';').strip()
        if '<' in normalized and '>' in normalized:
            start = normalized.find('<')
            end =')
            if start normalized.find('> < end:
                normalized = normalized[start+1:end]
        normalized = normalized.lower()
        return normalized
    
    unique_recipients = set()
    normalized_sender_email = normalize_email_address(sender_email)
    
    # Get TO recipients using enhanced logic
    to_recipients_data = simulated_cache.get('to_recipients', [])
    print(f"   Processing TO recipients: {to_recipients_data}")
    
    for recipient_info in to_recipients_data:
        if isinstance(recipient_info, dict):
            recipient_email = recipient_info.get('email', '').strip()
            recipient_display_name = recipient_info.get('display_name', '').strip()
            normalized_recipient_email = normalize_email_address(recipient_email)
            
            print(f"     - Recipient: '{recipient_email}' (display: '{recipient_display_name}')")
            print(f"     - Normalized recipient: '{normalized_recipient_email}'")
            print(f"     - Normalized sender: '{normalized_sender_email}'")
            
            if recipient_email and normalized_recipient_email != normalized_sender_email:
                if recipient_display_name:
                    recipient_string = f"{recipient_display_name} <{recipient_email}>"
                else:
                    recipient_string = recipient_email
                unique_recipients.add(recipient_string)
                print(f"     ✅ Added: {recipient_string}")
            else:
                print(f"     ❌ Excluded (sender or no email)")
    
    # Get CC recipients using enhanced logic
    cc_recipients_data = simulated_cache.get('cc_recipients', [])
    print(f"\n   Processing CC recipients: {cc_recipients_data}")
    
    for recipient_info in cc_recipients_data:
        if isinstance(recipient_info, dict):
            recipient_email = recipient_info.get('email', '').strip()
            recipient_display_name = recipient_info.get('display_name', '').strip()
            normalized_recipient_email = normalize_email_address(recipient_email)
            
            print(f"     - Recipient: '{recipient_email}' (display: '{recipient_display_name}')")
            print(f"     - Normalized recipient: '{normalized_recipient_email}'")
            print(f"     - Normalized sender: '{normalized_sender_email}'")
            
            if recipient_email and normalized_recipient_email != normalized_sender_email:
                if recipient_display_name:
                    recipient_string = f"{recipient_display_name} <{recipient_email}>"
                else:
                    recipient_string = recipient_email
                unique_recipients.add(recipient_string)
                print(f"     ✅ Added: {recipient_string}")
            else:
                print(f"     ❌ Excluded (sender or no email)")
    
    print(f"\n3. Final result:")
    if unique_recipients:
        cc_result = '; '.join(sorted(unique_recipients))
        print(f"   CC field: {cc_result}")
    else:
        print("   CC field: (empty)")
    
    print(f"\n4. Expected behavior:")
    print("   - Sender 'john.smith@company.com' should be excluded")
    print("   - Other recipients should appear with display names")
    print("   - Should see: 'Alice Johnson <alice.johnson@company.com>; Bob Wilson <bob.wilson@company.com>; Jane Doe <jane.doe@company.com>'")

if __name__ == "__main__":
    test_actual_workflow()