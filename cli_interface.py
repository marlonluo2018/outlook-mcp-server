import sys
import os
from typing import Optional

# Import handling for both direct execution and module usage
try:
    # First try absolute imports (direct execution)
    from backend.outlook_session import OutlookSessionManager
    from backend.email_retrieval import (
        list_recent_emails,
        search_emails,
        get_email_by_number,
        list_folders,
        view_email_cache
    )
    from backend.email_composition import (
        compose_email,
        reply_to_email_by_number
    )
    from backend.batch_operations import send_batch_emails
    from backend.shared import email_cache
except ImportError:
    try:
        # Then try relative imports (module usage)
        from .backend.outlook_session import OutlookSessionManager
        from .backend.email_retrieval import (
            search_emails,
            get_email_by_number,
            list_folders
        )
        from .backend.email_composition import (
            compose_email,
            reply_to_email_by_number
        )
        from .backend.batch_operations import send_batch_emails
        from .backend.shared import email_cache
    except ImportError:
        # Finally try direct imports from same directory
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from backend.outlook_session import OutlookSessionManager
        from backend.email_retrieval import (
            search_emails,
            get_email_by_number,
            list_folders
        )
        from backend.email_composition import (
            compose_email,
            reply_to_email_by_number
        )
        from backend.batch_operations import send_batch_emails
        from backend.shared import email_cache

def show_menu():
    print("\nOutlook MCP Server - Interactive Mode")
    print("1. List folders")
    print("2. List recent emails")
    print("3. Search emails")
    print("4. View email cache")
    print("5. Get email details")
    print("6. Reply to email")
    print("7. Compose new email")
    print("8. Send batch emails")
    print("0. Exit")

def interactive_mode():
    session = OutlookSessionManager()
    session._connect()
    
    while True:
        show_menu()
        choice = input("\nEnter command number: ").strip()
        
        try:
            if choice == "1":
                # List folders first
                folders = list_folders()
                print("\nAvailable folders:")
                for folder in folders:
                    print(folder)
                    
            elif choice == "2":
                # List recent emails
                days = input("Enter number of days (1-30): ").strip()
                folder = input("Enter folder name (leave blank for Inbox): ").strip() or None
                try:
                    days_int = int(days)
                    result = list_recent_emails(folder, days_int)
                    print(result)
                except ValueError as e:
                    print(f"Error: {str(e)}")
                    
            elif choice == "3":
                # Search emails
                term = input("Enter search term: ").strip()
                days_input = input("Enter number of days (1-30, leave blank for all): ").strip()
                folder = input("Enter folder name (leave blank for Inbox): ").strip() or None
                match_all = input("Match all terms? (y/n, default=y): ").strip().lower() != 'n'
                try:
                    days = int(days_input) if days_input else None
                    result = search_emails(term, days, folder, match_all)
                    print(result)
                except ValueError as e:
                    print(f"Error: {str(e)}")
                
            elif choice == "4":
                # View email cache with pagination
                try:
                    page = int(input("Enter starting page number (default: 1): ").strip() or 1)
                    if page < 1:
                        print("Page number must be positive, using page 1")
                        page = 1
                except ValueError:
                    print("Invalid page number, using page 1")
                    page = 1
                    
                while True:
                    result = view_email_cache(page)
                    print(f"\n{result}")
                    
                    print("\nNavigation:")
                    print("n - Next page")
                    print("p - Previous page")
                    print("q - Quit to menu")
                    
                    nav = input("\nEnter command: ").strip().lower()
                    if nav == 'n':
                        page += 1
                    elif nav == 'p':
                        page = max(1, page - 1)
                    elif nav == 'q':
                        break

            elif choice == "5":
                # Get full email by number
                if not email_cache:
                    print("\nNo emails in cache - load emails first")
                    continue
                    
                try:
                    num = int(input("Enter email number: ").strip())
                    if num < 1 or num > len(email_cache):
                        print("\nInvalid email number")
                        continue
                        
                    email_id = list(email_cache.keys())[num-1]
                    full_email = get_email_by_number(num)
                    if full_email:
                        print("\nFull email details:")
                        print(f"Subject: {full_email.get('subject')}")
                        print(f"From: {full_email.get('sender')}")
                        print(f"Date: {full_email.get('received')}")
                        print(f"\nBody:\n{full_email.get('body')}")
                        if full_email.get('attachments'):
                            print("\nAttachments:")
                            for attach in full_email['attachments']:
                                print(f"- {attach['name']} ({attach['size']} bytes)")
                except ValueError:
                    print("\nInvalid input - must be a number")

            elif choice == "6":
                # Reply to email
                if not email_cache:
                    print("\nNo emails in cache - load emails first")
                    continue
                    
                try:
                    num = int(input("Enter email number to reply to: ").strip())
                    if num < 1 or num > len(email_cache):
                        print("\nInvalid email number")
                        continue
                        
                    body = input("Enter reply text: ").strip()
                    print(reply_to_email_by_number(num, body))
                except ValueError:
                    print("\nInvalid input - must be a number")

            elif choice == "7":
                # Compose new email
                to = input("Enter To recipients (comma separated): ").strip()
                subject = input("Enter subject: ").strip()
                body = input("Enter email body: ").strip()
                cc = input("Enter CC recipients (comma separated, blank for none): ").strip()
                try:
                    to_list = [x.strip() for x in to.split(",")] if to else []
                    cc_list = [x.strip() for x in cc.split(",")] if cc else []
                    print(compose_email(to_list, subject, body, cc_list))
                except Exception as e:
                    print(f"Error composing email: {str(e)}")
                    
            elif choice == "8":
                # Batch send emails
                if not email_cache:
                    print("\nNo emails in cache - load emails first")
                    continue
                    
                try:
                    num = int(input("Enter email number from cache: ").strip())
                    if num < 1 or num > len(email_cache):
                        print("\nInvalid email number")
                        continue
                        
                    csv_path = input("Enter path to CSV file: ").strip()
                    custom_text = input("Enter custom text to prepend (optional): ").strip()
                    print(send_batch_emails(num, csv_path, custom_text))
                except ValueError:
                    print("\nInvalid input - must be a number")
                
            elif choice == "0":
                break
                
        except Exception as e:
            print(f"Error: {str(e)}", file=sys.stderr)
            continue

def main():
    interactive_mode()

if __name__ == '__main__':
    main()