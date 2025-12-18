#!/usr/bin/env python3
"""Direct test of the search functionality."""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from outlook_mcp_server.backend.email_search.subject_search import search_email_by_subject

def test_direct_search():
    """Test the search functionality directly."""
    print("Testing direct search for 'Your Approval is required'...")
    
    try:
        results, message = search_email_by_subject('Your Approval is required', days=1, match_all=False)
        
        print(f"Message: {message}")
        print(f"Found {len(results)} emails")
        
        if results:
            print("\nFirst 3 results:")
            for i, email in enumerate(results[:3]):
                print(f"{i+1}. {email['subject']} - {email['received_time']}")
        else:
            print("No results found")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_direct_search()