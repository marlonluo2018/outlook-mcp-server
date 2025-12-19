#!/usr/bin/env python3
"""Debug script to check email dates in the inbox."""

import sys
import os
import logging
import win32com.client
import pythoncom
from datetime import datetime, timedelta, timezone

# Enable debug logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def check_email_dates():
    """Check the dates of emails in the inbox to understand the date distribution."""
    pythoncom.CoInitialize()
    
    try:
        # Initialize Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        
        total_items = inbox.Items.Count
        print(f"Total items in inbox: {total_items}")
        
        # Check first 50 emails to see their dates
        print("\nChecking first 50 emails (newest first):")
        
        now = datetime.now(timezone.utc)
        one_day_ago = now - timedelta(days=1)
        seven_days_ago = now - timedelta(days=7)
        
        recent_count = 0
        week_count = 0
        
        for i in range(1, min(51, total_items + 1)):  # Check first 50 emails
            try:
                item_index = total_items - i + 1  # Process newest first
                item = inbox.Items.Item(item_index)
                
                if not hasattr(item, 'ReceivedTime') or not item.ReceivedTime:
                    continue
                
                received_time = item.ReceivedTime
                if received_time.tzinfo is None:
                    received_time = received_time.replace(tzinfo=timezone.utc)
                
                # Check if recent
                if received_time >= one_day_ago:
                    recent_count += 1
                elif received_time >= seven_days_ago:
                    week_count += 1
                
                # Print first 10 for debugging
                if i <= 10:
                    print(f"  Email {i}: {received_time.strftime('%Y-%m-%d %H:%M')} - {getattr(item, 'Subject', 'No subject')[:50]}")
                
            except Exception as e:
                print(f"  Error processing email {i}: {e}")
                continue
        
        print(f"\nSummary:")
        print(f"  Emails from last 24 hours: {recent_count}")
        print(f"  Emails from last 7 days: {week_count + recent_count}")
        print(f"  Emails older than 7 days: {50 - recent_count - week_count}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    check_email_dates()