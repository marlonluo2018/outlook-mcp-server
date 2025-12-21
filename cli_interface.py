import sys
import os
from typing import Optional

# Import handling for both direct execution and module usage
try:
    # First try imports from outlook_mcp_server package (direct execution)
    from outlook_mcp_server.backend.outlook_session.session_manager import OutlookSessionManager
    from outlook_mcp_server.backend.email_search import (
        list_recent_emails,
        search_email_by_subject,
        search_email_by_from,
        search_email_by_to,
        search_email_by_body,
        list_folders
    )
    from outlook_mcp_server.tools.viewing_tools import view_email_cache_tool as view_email_cache
    from outlook_mcp_server.backend.email_data_extractor import get_email_by_number_unified
    from outlook_mcp_server.backend.email_composition import (
        compose_email,
        reply_to_email_by_number
    )
    from outlook_mcp_server.backend.batch_operations import batch_forward_emails
    from outlook_mcp_server.backend.shared import email_cache
except ImportError:
    try:
        # Then try relative imports (module usage)
        from .backend.outlook_session.session_manager import OutlookSessionManager
        from .backend.email_search import (
            list_recent_emails,
            search_email_by_subject,
            search_email_by_from,
            search_email_by_to,
            search_email_by_body,
            list_folders
        )
        from .tools.viewing_tools import view_email_cache_tool as view_email_cache
        from .backend.email_data_extractor import get_email_by_number_unified
        from .backend.email_composition import (
            compose_email,
            reply_to_email_by_number
        )
        from .backend.batch_operations import batch_forward_emails
        from .backend.shared import email_cache
    except ImportError:
        # Finally try direct imports from same directory
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from outlook_mcp_server.backend.outlook_session.session_manager import OutlookSessionManager
        from outlook_mcp_server.backend.email_search import (
            list_recent_emails,
            search_email_by_subject,
            search_email_by_from,
            search_email_by_to,
            search_email_by_body,
            list_folders
        )
        from outlook_mcp_server.tools.viewing_tools import view_email_cache_tool as view_email_cache
        from outlook_mcp_server.backend.email_data_extractor import get_email_by_number_unified
        from outlook_mcp_server.backend.email_composition import (
            compose_email,
            reply_to_email_by_number
        )
        from outlook_mcp_server.backend.batch_operations import batch_forward_emails
        from outlook_mcp_server.backend.shared import email_cache

def show_menu():
    print("\nOutlook MCP Server - Interactive Mode")
    print("1. List folders")
    print("2. List recent emails")
    print("3. Search email subjects")
    print("4. Search emails by sender name")
    print("5. Search emails by recipient name")
    print("6. Search emails by body content")
    print("7. View email cache")
    print("8. Get email details")
    print("9. Reply to email")
    print("10. Compose new email")
    print("11. Batch forward emails")
    print("12. Create folder")
    print("13. Remove folder")
    print("14. Move email")
    print("15. Delete email")
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
                # Search email subjects
                print("This function only searches the email subject field.")
                print("It does not search in the email body, sender name, recipients, or other fields.")
                term = input("Enter search term: ").strip()
                days_input = input("Enter number of days (1-30, leave blank for all): ").strip()
                folder = input("Enter folder name (leave blank for Inbox): ").strip() or None
                match_all = input("Match all terms? (y/n, default=y): ").strip().lower() != 'n'
                try:
                    days = int(days_input) if days_input else None
                    emails, note = search_email_by_subject(term, days, folder, match_all)
                    result = f"Found {len(emails)} matching emails{note}"
                    print(result)
                except ValueError as e:
                    print(f"Error: {str(e)}")
            
            elif choice == "4":
                # Search emails by sender name
                print("This function only searches the sender name field.")
                print("It does not search in the email body, subject, recipients, or other fields.")
                print("IMPORTANT: Due to Microsoft Exchange's Distinguished Name format for internal email addresses,")
                print("this function only searches by display name, not email address.")
                print("Use display names (e.g., 'John Doe') instead of email addresses (e.g., 'john.doe@example.com').")
                term = input("Enter sender name to search for: ").strip()
                days_input = input("Enter number of days (1-30, leave blank for all): ").strip()
                folder = input("Enter folder name (leave blank for Inbox): ").strip() or None
                try:
                    days = int(days_input) if days_input else None
                    emails, note = search_email_by_from(term, days, folder, True)
                    result = f"Found {len(emails)} matching emails{note}"
                    print(result)
                except ValueError as e:
                    print(f"Error: {str(e)}")
            
            elif choice == "5":
                # Search emails by recipient name
                print("This function only searches the recipient (To) field.")
                print("It does not search in the email body, subject, sender, or other fields.")
                print("IMPORTANT: Due to Microsoft Exchange's Distinguished Name format for internal email addresses,")
                print("this function only searches by display name, not email address.")
                print("Use display names (e.g., 'John Doe') instead of email addresses (e.g., 'john.doe@example.com').")
                term = input("Enter recipient name to search for: ").strip()
                days_input = input("Enter number of days (1-30, leave blank for all): ").strip()
                folder = input("Enter folder name (leave blank for Inbox): ").strip() or None
                try:
                    days = int(days_input) if days_input else None
                    emails, note = search_email_by_to(term, days, folder, True)
                    result = f"Found {len(emails)} matching emails{note}"
                    print(result)
                except ValueError as e:
                    print(f"Error: {str(e)}")
            
            elif choice == "6":
                # Search emails by body content
                print("This function searches the email body content.")
                print("It does not search in the subject, sender name, recipients, or other fields.")
                print("Note: Searching email body is slower than searching other fields as it requires")
                print("loading the full content of each email.")
                print("\nSearch options:")
                print("- For exact phrase matching, enclose your search term in quotes (e.g., \"red hat partner day\")")
                print("- For word-based matching, enter terms without quotes (e.g., red hat partner day)")
                print("- Word-based matching uses AND logic by default (all words must be present)")
                term = input("Enter search term for email body: ").strip()
                days_input = input("Enter number of days (1-30, leave blank for all): ").strip()
                folder = input("Enter folder name (leave blank for Inbox): ").strip() or None
                match_all = input("Match all terms? (y/n, default=y): ").strip().lower() != 'n'
                try:
                    days = int(days_input) if days_input else None
                    emails, note = search_email_by_body(term, days, folder, match_all)
                    result = f"Found {len(emails)} matching emails{note}"
                    print(result)
                except ValueError as e:
                    print(f"Error: {str(e)}")
                
            elif choice == "7":
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
                    
                    # Handle JSON response format
                    if result.get("type") == "json":
                        data = result.get("data", {})
                        
                        # Check for errors
                        if "error" in data:
                            print(f"\nError: {data['error']}")
                            if "message" in data:
                                print(f"Message: {data['message']}")
                            break
                        
                        # Display email information
                        print(f"\nCached Emails (Page {data['page']} of {data['total_pages']}):")
                        print(f"Total emails in cache: {data['total_emails']}\n")
                        
                        for email in data['emails']:
                            print(f"{email['number']}. {email['subject']}")
                            print(f"   From: {email['from']}")
                            print(f"   To: {email['to']}")
                            print(f"   CC: {email['cc']}")
                            print(f"   Received: {email['received']}")
                            print(f"   Status: {email['status']}")
                            embedded_images_display = str(email['embedded_images_count']) if email['embedded_images_count'] > 0 else "None"
                            print(f"   Embedded Images: {embedded_images_display}")
                            print(f"   Attachments: {email['attachments']}")
                            print()
                    else:
                        # Fallback for any other format
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

            elif choice == "8":
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
                    full_email = get_email_by_number_unified(num)
                    if full_email:
                        print("\nFull email details:")
                        print(f"Subject: {full_email.get('subject')}")
                        print(f"From: {full_email.get('sender')}")
                        if full_email.get('to'):
                            print(f"To: {full_email.get('to')}")
                        if full_email.get('cc'):
                            print(f"CC: {full_email.get('cc')}")
                        print(f"Date: {full_email.get('received_time')}")
                        print(f"\nBody:\n{full_email.get('body')}")
                        if full_email.get('attachments'):
                            print("\nAttachments:")
                            for attach in full_email['attachments']:
                                print(f"- {attach['name']} ({attach['size']} bytes)")
                except ValueError:
                    print("\nInvalid input - must be a number")

            elif choice == "9":
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

            elif choice == "10":
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
                    
            elif choice == "11":
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
                    print("\n" + batch_forward_emails(num, csv_path, custom_text))
                except ValueError:
                    print("\nInvalid input - must be a number")
                    
            elif choice == "12":
                # Create folder
                try:
                    folder_name = input("Enter new folder name: ").strip()
                    parent_folder = input("Enter parent folder name (leave blank for Inbox): ").strip() or None
                    with OutlookSessionManager() as outlook_session:
                        result = outlook_session.create_folder(folder_name, parent_folder)
                        print(f"\n{result}")
                except Exception as e:
                    print(f"\nError creating folder: {str(e)}")
                    
            elif choice == "13":
                # Remove folder
                try:
                    folder_name = input("Enter folder name or path to remove: ").strip()
                    with OutlookSessionManager() as outlook_session:
                        result = outlook_session.remove_folder(folder_name)
                        print(f"\n{result}")
                except Exception as e:
                    print(f"\nError removing folder: {str(e)}")
                    
            elif choice == "14":
                # Move email
                if not email_cache:
                    print("\nNo emails in cache - load emails first")
                    continue
                    
                try:
                    num = int(input("Enter email number to move: ").strip())
                    if num < 1 or num > len(email_cache):
                        print("\nInvalid email number")
                        continue
                        
                    target_folder = input("Enter target folder name: ").strip()
                    email_id = list(email_cache.keys())[num-1]
                    
                    with OutlookSessionManager() as outlook_session:
                        result = outlook_session.move_email(email_id, target_folder)
                        print(f"\n{result}")
                        # Refresh cache after moving
                        email_cache.clear()
                        print("Cache cleared - reload emails to see updated status")
                except ValueError:
                    print("\nInvalid input - must be a number")
                except Exception as e:
                    print(f"\nError moving email: {str(e)}")
                    
            elif choice == "15":
                # Delete email
                if not email_cache:
                    print("\nNo emails in cache - load emails first")
                    continue
                    
                try:
                    num = int(input("Enter email number to delete: ").strip())
                    if num < 1 or num > len(email_cache):
                        print("\nInvalid email number")
                        continue
                        
                    email_id = list(email_cache.keys())[num-1]
                    
                    with OutlookSessionManager() as outlook_session:
                        result = outlook_session.delete_email(email_id)
                        print(f"\n{result}")
                        # Refresh cache after deleting
                        email_cache.pop(email_id, None)
                        print("Email removed from cache")
                except ValueError:
                    print("\nInvalid input - must be a number")
                except Exception as e:
                    print(f"\nError deleting email: {str(e)}")
                    
            elif choice == "0":
                break
                
        except Exception as e:
            print(f"Error: {str(e)}", file=sys.stderr)
            continue

def main():
    interactive_mode()

if __name__ == '__main__':
    main()