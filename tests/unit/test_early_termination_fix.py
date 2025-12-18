#!/usr/bin/env python3
"""
Test script to verify the early termination bug fix.
This will test the email search with different timeframes.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'outlook_mcp_server'))

from datetime import datetime, timezone
from outlook_mcp_server.backend.email_search import list_recent_emails

def test_search_timeframes():
    """Test email search with different timeframes."""
    print("Testing email search with different timeframes...")
    print("=" * 60)
    
    timeframes = [
        (1, "1 day"),
        (3, "3 days"), 
        (7, "7 days"),
        (14, "14 days")
    ]
    
    for days, description in timeframes:
        print(f"\nTesting {description} search:")
        print("-" * 30)
        
        try:
            emails, message = list_recent_emails("Inbox", days)
            
            print(f"Results: {len(emails)} emails")
            print(f"Message: {message}")
            
            if emails:
                # Show first few emails
                print(f"\nFirst 3 emails:")
                for i, email in enumerate(emails[:3]):
                    subject = email.get('subject', 'No Subject')
                    received_time = email.get('received_time', 'Unknown')
                    print(f"  {i+1}. {subject[:50]}... ({received_time})")
                
                # Show last email
                if len(emails) > 3:
                    last_email = emails[-1]
                    subject = last_email.get('subject', 'No Subject')
                    received_time = last_email.get('received_time', 'Unknown')
                    print(f"  ...")
                    print(f"  {len(emails)}. {subject[:50]}... ({received_time})")
            
            print(f"\nRange: {len(emails)} emails found")
            
        except Exception as e:
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()

def test_specific_date_range():
    """Test emails from a specific recent date range."""
    print("\n" + "=" * 60)
    print("Testing specific date range verification...")
    
    # Get emails from last 7 days
    emails, message = list_recent_emails("Inbox", 7)
    
    if emails:
        print(f"\nFound {len(emails)} emails in last 7 days")
        
        # Check the date range
        from datetime import datetime, timezone, timedelta
        seven_days_ago = datetime.now(timezone.utc) - timedelta(days=7)
        
        valid_emails = 0
        for email in emails:
            received_time_str = email.get('received_time', '')
            if received_time_str and received_time_str != 'Unknown':
                try:
                    # Parse the received time
                    received_time = datetime.fromisoformat(received_time_str.replace('Z', '+00:00'))
                    if received_time >= seven_days_ago:
                        valid_emails += 1
                except:
                    pass
        
        print(f"Emails within 7-day window: {valid_emails}")
        print(f"Expected range: 200-2000 emails based on our analysis")
        
        if len(emails) < 100:
            print("⚠️  WARNING: Still getting fewer emails than expected!")
        else:
            print("✅ SUCCESS: Getting reasonable number of emails!")

if __name__ == "__main__":
    print("Testing early termination bug fix...")
    test_search_timeframes()
    test_specific_date_range()
    print("\n" + "=" * 60)
    print("Test completed!")