from outlook_mcp_server.backend.email_search import search_email_by_subject

print('Testing server-side search for "Your Approval is required"...')

try:
    # Test the search function directly
    results, message = search_email_by_subject(
        "Your Approval is required", 
        days=1,  # Only search last 1 day as requested
        match_all=False,  # Use LIKE search
    )
    
    print(f'Search completed: {message}')
    print(f'Results found: {len(results)}')
    
    if results:
        print('\nFirst 5 results:')
        for i, email in enumerate(results[:5]):
            print(f'{i+1}. Subject: {email.get("subject", "No Subject")}')
            print(f'   Received: {email.get("received_time", "Unknown")}')
            print()
    else:
        print('No results found.')
        
except Exception as e:
    print(f'Error: {e}')
    import traceback
    traceback.print_exc()