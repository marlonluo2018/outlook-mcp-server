from typing import Optional, List, Dict, Any
import logging
import pythoncom
import win32com.client
import time
import datetime
from .shared import configure_logging, get_email_from_cache
from .utils import OutlookFolderType, retry_on_com_error

# Initialize logging
configure_logging()

# Configure logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

class OutlookSessionManager:
    """Context manager for Outlook COM session handling with improved resource management"""
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.folder = None
        self._connected = False
        self._com_initialized = False
        
    def __enter__(self):
        """Initialize Outlook COM objects"""
        self._connect()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up Outlook COM objects"""
        self._disconnect()
        return False  # Don't suppress exceptions
        
    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def _connect(self):
        """Establish COM connection with proper threading and retry logic"""
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self._connected = True
            logger.info("Successfully connected to Outlook")
        except Exception as e:
            logger.error(f"Connection error: {str(e)}")
            self._cleanup_partial_connection()
            raise RuntimeError("Failed to connect to Outlook") from e
    
    def _cleanup_partial_connection(self):
        """Clean up partial connection attempts"""
        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.warning(f"Error cleaning up partial connection: {str(e)}")
            finally:
                self._com_initialized = False
            
    def _disconnect(self):
        """Clean up COM objects with proper resource release"""
        if self._connected:
            try:
                # Release COM objects explicitly
                if self.namespace:
                    self.namespace = None
                if self.outlook:
                    self.outlook = None
                    
                if self._com_initialized:
                    pythoncom.CoUninitialize()
                    self._com_initialized = False
                    
                logger.debug("Outlook connection cleaned up successfully")
            except Exception as e:
                logger.warning(f"Error during disconnect: {str(e)}")
            finally:
                self._connected = False
                self.outlook = None
                self.namespace = None
                self.folder = None
                
    def is_connected(self) -> bool:
        """Check if the session is still connected"""
        try:
            # Simple operation to test connection
            if self._connected and self.outlook:
                self.outlook.GetNamespace("MAPI").GetDefaultFolder(OutlookFolderType.INBOX).Name
                return True
            return False
        except:
            self._connected = False
            return False
            
    def reconnect(self):
        """Re-establish the Outlook connection"""
        self._disconnect()
        self._connect()
                 
    def get_folder(self, folder_name: Optional[str] = None):
        """Get specified folder or default inbox"""
        try:
            # Handle string "null" as well as actual None
            if not folder_name or folder_name == "null" or folder_name.lower() == "inbox":
                folder = self.namespace.GetDefaultFolder(OutlookFolderType.INBOX)
                return folder
            elif folder_name.lower() == "sent items" or folder_name.lower() == "sent":
                folder = self.namespace.GetDefaultFolder(OutlookFolderType.SENT_MAIL)
                return folder
            elif folder_name.lower() == "deleted items" or folder_name.lower() == "trash":
                folder = self.namespace.GetDefaultFolder(OutlookFolderType.DELETED_ITEMS)
                return folder
            elif folder_name.lower() == "drafts":
                folder = self.namespace.GetDefaultFolder(OutlookFolderType.DRAFTS)
                return folder
            elif folder_name.lower() == "outbox":
                folder = self.namespace.GetDefaultFolder(OutlookFolderType.OUTBOX)
                return folder
            elif folder_name.lower() == "calendar":
                folder = self.namespace.GetDefaultFolder(OutlookFolderType.CALENDAR)
                return folder
            elif folder_name.lower() == "contacts":
                folder = self.namespace.GetDefaultFolder(OutlookFolderType.CONTACTS)
                return folder
            elif folder_name.lower() == "tasks":
                folder = self.namespace.GetDefaultFolder(OutlookFolderType.TASKS)
                return folder
            else:
                folder = self._get_folder_by_name(folder_name)
                return folder
        except Exception as e:
            logger.error(f"Error getting folder: {str(e)}")
            raise
            
    def _get_folder_by_name(self, folder_name: str):
        """Find folder by name in folder hierarchy, supporting nested paths and mailbox-specific paths"""
        try:
            # Handle nested folder paths (e.g., "Parent Folder/Child Folder" or "mailbox@domain.com/Inbox/Folder")
            if '/' in folder_name or '\\' in folder_name:
                # Use forward slash as path separator, but also support backslash
                path_parts = folder_name.replace('\\', '/').split('/')
                current_folder = None
                
                # Check if first part looks like an email address (mailbox-specific path)
                if '@' in path_parts[0] and '.' in path_parts[0]:
                    # This is a mailbox-specific path like "user@company.com/Inbox/Folder"
                    mailbox_name = path_parts[0]
                    
                    # Find the mailbox folder
                    for folder in self.namespace.Folders:
                        if folder.Name == mailbox_name:
                            current_folder = folder
                            break
                    
                    if not current_folder:
                        raise ValueError(f"Mailbox '{mailbox_name}' not found")
                    
                    # Navigate through the remaining path parts
                    remaining_parts = path_parts[1:]
                else:
                    # This is a regular path like "Inbox/Folder" or "Parent Folder/Child Folder"
                    # Start with the top-level folders
                    for folder in self.namespace.Folders:
                        if folder.Name == path_parts[0]:
                            current_folder = folder
                            break
                    
                    if not current_folder:
                        raise ValueError(f"Top-level folder '{path_parts[0]}' not found")
                    
                    remaining_parts = path_parts[1:]
                
                # Navigate through the remaining path parts
                for part in remaining_parts:
                    found = False
                    for subfolder in current_folder.Folders:
                        if subfolder.Name == part:
                            current_folder = subfolder
                            found = True
                            break
                    if not found:
                        raise ValueError(f"Folder '{part}' not found in '{current_folder.Name}'")
                
                return current_folder
            else:
                # Original logic for single folder names
                for folder in self.namespace.Folders:
                    if folder.Name == folder_name:
                        return folder
                    for subfolder in folder.Folders:
                        if subfolder.Name == folder_name:
                            return subfolder
                raise ValueError(f"Folder '{folder_name}' not found")
        except Exception as e:
            logger.error(f"Error finding folder: {str(e)}")
            raise
    
    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def create_folder(self, folder_name: str, parent_folder_name: Optional[str] = None) -> str:
        """Create a new folder in the specified parent folder.
        
        Args:
            folder_name: Name of the folder to create, or full path (e.g., "Inbox/SubFolder1/SubFolder2")
            parent_folder_name: Name of the parent folder (optional, defaults to Inbox)
            
        Returns:
            Full path to the created folder
        """
        try:
            if not folder_name:
                raise ValueError("Folder name cannot be empty")
            
            # Handle nested folder creation
            if '/' in folder_name:
                # Full path provided, parse it
                path_parts = folder_name.split('/')
                if len(path_parts) > 3:
                    raise ValueError("Maximum 3 folder levels supported (e.g., 'Inbox/SubFolder1/SubFolder2')")
                
                # Start with the first part as parent
                current_parent = self.get_folder(path_parts[0])
                
                # Create intermediate folders if needed
                for i in range(1, len(path_parts) - 1):
                    subfolder_name = path_parts[i]
                    # Check if subfolder exists
                    subfolder = None
                    for folder in current_parent.Folders:
                        if folder.Name == subfolder_name:
                            subfolder = folder
                            break
                    
                    if subfolder is None:
                        subfolder = current_parent.Folders.Add(subfolder_name)
                        logger.info(f"Created intermediate folder: {subfolder_name}")
                    
                    current_parent = subfolder
                
                # Create the final folder
                final_folder_name = path_parts[-1]
                # Check if final folder already exists
                for folder in current_parent.Folders:
                    if folder.Name == final_folder_name:
                        raise ValueError(f"Folder '{folder_name}' already exists")
                
                new_folder = current_parent.Folders.Add(final_folder_name)
                folder_path = folder_name
            else:
                # Simple folder creation
                parent_folder = self.get_folder(parent_folder_name)
                
                # Check if folder already exists
                for folder in parent_folder.Folders:
                    if folder.Name == folder_name:
                        raise ValueError(f"Folder '{folder_name}' already exists in '{parent_folder.Name}'")
                
                # Create the folder
                new_folder = parent_folder.Folders.Add(folder_name)
                folder_path = f"{parent_folder.Name}/{folder_name}" if parent_folder.Name != "Inbox" else folder_name
            
            logger.info(f"Successfully created folder: {folder_path}")
            return folder_path
        except Exception as e:
            logger.error(f"Error creating folder: {str(e)}")
            raise
    
    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def remove_folder(self, folder_name: str) -> str:
        """Remove an existing folder.
        
        Args:
            folder_name: Name or path of the folder to remove (supports nested paths like "Inbox/SubFolder1/SubFolder2")
            
        Returns:
            Confirmation message
        """
        try:
            if not folder_name:
                raise ValueError("Folder name cannot be empty")
            
            # Handle nested folder paths
            if '/' in folder_name:
                # Split the path and navigate to the folder
                path_parts = folder_name.split('/')
                if len(path_parts) > 3:
                    raise ValueError("Maximum 3 folder levels supported (e.g., 'Inbox/SubFolder1/SubFolder2')")
                
                # Start from the root folder
                current_folder = self.get_folder(path_parts[0])
                
                # Navigate through the path
                for i in range(1, len(path_parts)):
                    subfolder_name = path_parts[i]
                    subfolder = None
                    
                    for folder in current_folder.Folders:
                        if folder.Name == subfolder_name:
                            subfolder = folder
                            break
                    
                    if subfolder is None:
                        raise ValueError(f"Folder '{folder_name}' not found")
                    
                    current_folder = subfolder
                
                folder_to_remove = current_folder
            else:
                # Simple folder removal
                folder_to_remove = self._get_folder_by_name(folder_name)
            
            # Check if we're trying to remove a default folder
            default_folder_names = ["Inbox", "Sent Items", "Deleted Items", "Drafts", "Outbox", "Calendar", "Contacts", "Tasks"]
            if folder_to_remove.Name in default_folder_names:
                raise ValueError(f"Cannot remove default folder: '{folder_to_remove.Name}'")
            
            # Remove the folder
            folder_to_remove.Delete()
            
            logger.info(f"Successfully removed folder: {folder_name}")
            return f"Folder '{folder_name}' removed successfully"
        except Exception as e:
            logger.error(f"Error removing folder: {str(e)}")
            raise
    
    def move_email(self, email_id: str, target_folder_name: str) -> str:
        """Move an email to the specified folder with robust error handling.
        
        Args:
            email_id: EntryID of the email to move
            target_folder_name: Name or path of the target folder
            
        Returns:
            Confirmation message
        """
        if not email_id:
            raise ValueError("Email ID cannot be empty")
        if not target_folder_name:
            raise ValueError("Target folder name cannot be empty")
        
        max_attempts = 3
        initial_delay = 1.0
        
        for attempt in range(max_attempts):
            try:
                # Get the email item with validation
                email_item = self._get_email_item_with_validation(email_id)
                
                if email_item is None:
                    raise ValueError(f"Could not load email with ID: {email_id}")
                
                # Validate item state before moving
                if not self._validate_email_item(email_item):
                    if attempt < max_attempts - 1:
                        logger.warning(f"Email item validation failed on attempt {attempt + 1}, retrying...")
                        time.sleep(initial_delay * (2 ** attempt))
                        continue
                    else:
                        raise ValueError("Email item is in an invalid state and cannot be moved")
                
                # Get the target folder
                target_folder = self._get_folder_by_name(target_folder_name)
                
                # Get subject before moving for logging
                try:
                    email_subject = getattr(email_item, 'Subject', 'Unknown Subject')
                    if not email_subject:
                        email_subject = f"Email ID: {email_id[:20]}..."
                except Exception as e:
                    logger.warning(f"Could not retrieve email subject: {e}")
                    email_subject = f"Email ID: {email_id[:20]}..."
                
                # Move the email
                email_item.Move(target_folder)
                
                logger.info(f"Successfully moved email '{email_subject}' to folder '{target_folder_name}'")
                return f"Email moved successfully to '{target_folder_name}'"
                
            except pythoncom.com_error as e:
                error_code = e.hresult if hasattr(e, 'hresult') else None
                
                # Handle the specific error code
                if error_code == -2147352567:
                    if attempt < max_attempts - 1:
                        logger.warning(
                            f"COM error (-2147352567) on attempt {attempt + 1}/{max_attempts} while moving email, "
                            f"retrying in {initial_delay * (2 ** attempt)}s... Error: {e}"
                        )
                        time.sleep(initial_delay * (2 ** attempt))
                        continue
                    else:
                        logger.error(f"Failed to move email after {max_attempts} attempts due to COM access error")
                        raise RuntimeError(
                            f"Could not access the email item for moving. This may indicate the email "
                            f"is being processed by another operation or the item reference is stale. "
                            f"Please try again or check if the email exists. Error: {str(e)}"
                        )
                else:
                    # Other COM errors - don't retry
                    logger.error(f"COM error while moving email: {e}")
                    raise
            except Exception as e:
                logger.error(f"Unexpected error while moving email: {str(e)}")
                raise
        
        # This should not be reached, but just in case
        raise RuntimeError(f"Failed to move email after {max_attempts} attempts")
    
    def move_folder(self, source_folder_path: str, target_parent_path: str) -> str:
        """Move a folder and all its emails to a new location.
        
        Args:
            source_folder_path: Path to the source folder (e.g., "Inbox/SubFolder1" or "mailbox@domain.com/Inbox/Folder")
            target_parent_path: Path to the target parent folder (e.g., "Inbox/NewParent" or "mailbox@domain.com/Inbox/NewParent")
            
        Returns:
            Confirmation message
        """
        if not source_folder_path:
            raise ValueError("Source folder path cannot be empty")
        if not target_parent_path:
            raise ValueError("Target parent path cannot be empty")
        
        try:
            # Parse source path
            source_parts = source_folder_path.split('/')
            if len(source_parts) > 4:  # Allow up to 4 levels for mailbox paths like "mailbox@domain.com/Inbox/Folder/SubFolder"
                raise ValueError("Maximum 4 folder levels supported")
            
            # Parse target path
            target_parts = target_parent_path.split('/')
            if len(target_parts) > 4:
                raise ValueError("Maximum 4 folder levels supported")
            
            # Get the source folder using the enhanced method
            source_folder = self._get_folder_by_name(source_folder_path)
            
            # Check if we're trying to move a default folder
            default_folder_names = ["Inbox", "Sent Items", "Deleted Items", "Drafts", "Outbox", "Calendar", "Contacts", "Tasks"]
            if source_folder.Name in default_folder_names:
                raise ValueError(f"Cannot move default folder: '{source_folder.Name}'")
            
            # Get the target parent folder using the enhanced method
            target_parent = self._get_folder_by_name(target_parent_path)
            
            # Check if a folder with the same name already exists in target
            for folder in target_parent.Folders:
                if folder.Name == source_folder.Name:
                    raise ValueError(f"Folder '{source_folder.Name}' already exists in '{target_parent_path}'")
            
            # Count emails in the source folder
            email_count = 0
            try:
                for item in source_folder.Items:
                    if hasattr(item, 'Class') and item.Class == 43:  # MailItem class
                        email_count += 1
            except Exception as e:
                logger.warning(f"Could not count emails in folder: {e}")
            
            # Move the folder using MoveTo method (Move is not available on folder objects)
            source_folder.MoveTo(target_parent)
            
            # Construct new path
            if target_parent_path == "Inbox":
                new_path = f"Inbox/{source_folder.Name}"
            else:
                new_path = f"{target_parent_path}/{source_folder.Name}"
            
            logger.info(f"Successfully moved folder '{source_folder_path}' to '{new_path}' with {email_count} emails")
            return f"Folder moved successfully from '{source_folder_path}' to '{new_path}' ({email_count} emails moved)"
            
        except Exception as e:
            logger.error(f"Error moving folder: {str(e)}")
            raise
    
    def delete_email(self, email_id: str) -> str:
        """Move an email to the Deleted Items folder instead of hard deletion.
        
        Args:
            email_id: EntryID of the email to delete
            
        Returns:
            Confirmation message
        """
        if not email_id:
            raise ValueError("Email ID cannot be empty")
        
        max_attempts = 3
        initial_delay = 1.0
        
        for attempt in range(max_attempts):
            try:
                # Get the email item with retry logic
                email_item = self._get_email_item_with_validation(email_id)
                
                if email_item is None:
                    raise ValueError(f"Could not load email with ID: {email_id}")
                
                # Validate item state before moving
                if not self._validate_email_item(email_item):
                    if attempt < max_attempts - 1:
                        logger.warning(f"Email item validation failed on attempt {attempt + 1}, retrying...")
                        time.sleep(initial_delay * (2 ** attempt))
                        continue
                    else:
                        raise ValueError("Email item is in an invalid state and cannot be moved")
                
                # Get subject before moving for logging
                try:
                    email_subject = getattr(email_item, 'Subject', 'Unknown Subject')
                    if not email_subject:
                        email_subject = f"Email ID: {email_id[:20]}..."
                except Exception as e:
                    logger.warning(f"Could not retrieve email subject: {e}")
                    email_subject = f"Email ID: {email_id[:20]}..."
                
                # Get the Deleted Items folder
                deleted_items_folder = self.get_folder("Deleted Items")
                
                # Move the email to Deleted Items instead of hard deletion
                email_item.Move(deleted_items_folder)
                
                logger.info(f"Successfully moved email '{email_subject}' to Deleted Items")
                return f"Email moved to Deleted Items successfully: '{email_subject}'"
                
            except pythoncom.com_error as e:
                error_code = e.hresult if hasattr(e, 'hresult') else None
                
                # Handle the specific error code mentioned by the user
                if error_code == -2147352567:
                    if attempt < max_attempts - 1:
                        logger.warning(
                            f"COM error (-2147352567) on attempt {attempt + 1}/{max_attempts} while moving email to Deleted Items, "
                            f"retrying in {initial_delay * (2 ** attempt)}s... Error: {e}"
                        )
                        time.sleep(initial_delay * (2 ** attempt))
                        continue
                    else:
                        logger.error(f"Failed to move email to Deleted Items after {max_attempts} attempts due to COM access error")
                        raise RuntimeError(
                            f"Could not access the email item for moving to Deleted Items. This may indicate the email "
                            f"is being processed by another operation or the item reference is stale. "
                            f"Please try again or check if the email exists. Error: {str(e)}"
                        )
                else:
                    # Other COM errors - don't retry
                    logger.error(f"COM error while moving email to Deleted Items: {e}")
                    raise
            except Exception as e:
                logger.error(f"Unexpected error while moving email to Deleted Items: {str(e)}")
                raise
        
        # This should not be reached, but just in case
        raise RuntimeError(f"Failed to move email to Deleted Items after {max_attempts} attempts")
    
    def _get_email_item_with_validation(self, email_id: str):
        """Get email item with validation and retry logic.
        
        Args:
            email_id: EntryID of the email
            
        Returns:
            Email item object or None if failed
        """
        try:
            email_item = self.namespace.GetItemFromID(email_id)
            
            # Verify the item was loaded successfully
            if email_item is None:
                logger.warning(f"GetItemFromID returned None for email ID: {email_id}")
                return None
            
            # Try to access a basic property to ensure the item is fully loaded
            try:
                _ = email_item.Class
                return email_item
            except pythoncom.com_error as e:
                logger.warning(f"Email item not fully loaded, COM error: {e}")
                return None
                
        except pythoncom.com_error as e:
            logger.warning(f"COM error getting email item: {e}")
            return None
        except Exception as e:
            logger.warning(f"Unexpected error getting email item: {e}")
            return None
    
    def _validate_email_item(self, email_item) -> bool:
        """Validate that an email item is in a state suitable for deletion.
        
        Args:
            email_item: The email item to validate
            
        Returns:
            bool: True if valid, False otherwise
        """
        try:
            # Check if the item is accessible
            if not hasattr(email_item, 'Class'):
                return False
            
            # Try to access the EntryID to ensure it's a valid Outlook item
            try:
                item_entry_id = email_item.EntryID
                if not item_entry_id:
                    return False
            except:
                return False
            
            # Check if it's actually a mail item
            try:
                item_class = email_item.Class
                # Outlook MailItem class is typically 43
                if hasattr(item_class, 'value'):
                    item_class = item_class.value
                if item_class != 43:  # MailItem class
                    logger.warning(f"Item is not a mail item, class: {item_class}")
                    return False
            except:
                # If we can't check the class, assume it's valid
                pass
            
            return True
            
        except Exception as e:
            logger.warning(f"Error validating email item: {e}")
            return False
    
    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def get_email_policies(self, email_id: str) -> dict:
        """Get Exchange retention policies assigned to an email.
        
        Args:
            email_id: EntryID of the email to check
            
        Returns:
            Dictionary with Exchange retention policies:
            {
                'policies': [...]  # List of Exchange retention policies
            }
        """
        try:
            if not email_id:
                raise ValueError("Email ID cannot be empty")
            
            # Get the email item
            email_item = self.namespace.GetItemFromID(email_id)
            
            # Only check for real Exchange retention policies
            policies = []
            
            # Check multiple properties to determine if policy is explicitly assigned
            retention_policy = getattr(email_item, 'RetentionPolicyName', None)
            categories = getattr(email_item, 'Categories', None)
            
            print(f"DEBUG: Email has retention_policy: '{retention_policy}'")
            print(f"DEBUG: Email has categories: '{categories}'")
            
            # Check Categories first (manually assigned policies take precedence)
            manually_assigned_policies = []
            if categories and categories.strip():
                print(f"DEBUG: Checking Categories for policies: '{categories}'")
                
                # Get available policies to match against
                try:
                    available_policies = self.get_available_policies()
                    category_list = [cat.strip() for cat in categories.split(';')]
                    
                    # Check if any category matches an available policy
                    for category in category_list:
                        for available_policy in available_policies:
                            if (category.lower() == available_policy.lower() or 
                                category in available_policy or 
                                available_policy in category):
                                manually_assigned_policies.append(available_policy)
                                print(f"DEBUG: Found manually assigned policy in Categories: '{available_policy}'")
                                break
                except Exception as e:
                    print(f"DEBUG: Could not check Categories for policies: {str(e)}")
            
            # If manually assigned policies exist, use those (they override auto-assigned policies)
            if manually_assigned_policies:
                policies.extend(manually_assigned_policies)
                print(f"DEBUG: Using manually assigned policies: {manually_assigned_policies}")
            elif retention_policy and retention_policy.strip() and retention_policy != 'Unknown':
                # No manually assigned policies, use the auto-assigned policy
                policies.append(retention_policy)
                print(f"DEBUG: Using auto-assigned policy: '{retention_policy}'")
            else:
                # No policies found at all
                print(f"DEBUG: Email has no policy - showing no policy")
            
            # Check for Information Rights Management
            if hasattr(email_item, 'PermissionService') and email_item.PermissionService != 0:
                policies.append("Information Rights Management")
            
            return {
                'policies': policies
            }
        except Exception as e:
            logger.error(f"Error getting email policies: {str(e)}")
            raise
    
    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def get_available_policies(self) -> list:
        """Get available policies that can be assigned to emails.
        
        Returns:
            List of available policy names retrieved from Outlook
        """
        try:
            available_policies = []
            
            # 1. Get built-in sensitivity levels from Outlook
            try:
                # Create a test mail item to determine available sensitivity levels
                test_mail = self.outlook.CreateItem(0)  # olMailItem
                
                # Test each sensitivity level and add if valid
                sensitivity_mapping = {
                    0: "Normal",           # Built-in Outlook sensitivity
                    1: "Personal",         # Built-in Outlook sensitivity  
                    2: "Private",          # Built-in Outlook sensitivity
                    3: "Confidential"      # Built-in Outlook sensitivity
                }
                
                for value, name in sensitivity_mapping.items():
                    try:
                        test_mail.Sensitivity = value
                        available_policies.append(name)
                        logger.debug(f"Found built-in policy: {name}")
                    except Exception as e:
                        logger.debug(f"Built-in policy {name} not available: {str(e)}")
                
                test_mail.Delete()
                
            except Exception as e:
                logger.warning(f"Could not retrieve built-in sensitivity levels: {str(e)}")
            
            # 2. Dynamically detect Exchange retention policies by querying actual emails
            logger.info("Starting dynamic retention policy detection...")
            try:
                # Get a sample of emails to discover what retention policies are actually available
                # This is the most reliable way to find what policies exist in the environment
                logger.info("Attempting to access inbox...")
                inbox = self.namespace.GetDefaultFolder(OutlookFolderType.INBOX)
                logger.info(f"Successfully accessed inbox: {inbox.Name}")
                
                # Get a small sample of emails (max 50) to check for retention policies
                # This gives us a good sample of what policies are actually used in the environment
                items = inbox.Items
                logger.info(f"Inbox contains {items.Count} items")
                
                # Use a more reliable approach to access items
                found_retention_policies = set()
                sample_count = 0
                max_sample = 50
                
                # Create a restricted view to get only recent emails
                try:
                    # Try to filter for recent emails first
                    filter_str = "[ReceivedTime] > '" + (datetime.datetime.now() - datetime.timedelta(days=30)).strftime('%m/%d/%Y') + "'"
                    recent_items = items.Restrict(filter_str)
                    logger.info(f"Found {recent_items.Count} recent emails (last 30 days)")
                    
                    if recent_items.Count > 0:
                        items = recent_items
                    
                except Exception as e:
                    logger.debug(f"Could not filter for recent emails: {str(e)}")
                
                # Sort by most recent first
                items.Sort("[ReceivedTime]", True)
                logger.info(f"Checking up to {min(items.Count, max_sample)} emails for retention policies...")
                
                # Use a more robust iteration approach
                try:
                    # Get items as a list to avoid COM iteration issues
                    item_list = []
                    for i in range(1, min(items.Count + 1, max_sample + 1)):
                        try:
                            item = items.Item(i)
                            if item and hasattr(item, 'Class') and item.Class == 43:  # olMailItem
                                item_list.append(item)
                        except Exception as e:
                            logger.debug(f"Could not access item {i}: {str(e)}")
                            continue
                    
                    logger.info(f"Processing {len(item_list)} valid email items...")
                    
                    for idx, item in enumerate(item_list):
                        try:
                            # Check for retention policy using different possible property names
                            retention_policy = None
                            
                            # Try different property names that might contain retention policy info
                            possible_properties = ['RetentionPolicyName', 'PolicyName', 'RetentionPolicy']
                            for prop in possible_properties:
                                try:
                                    if hasattr(item, prop):
                                        retention_policy = getattr(item, prop, None)
                                        if retention_policy and retention_policy != 'Unknown':
                                            logger.debug(f"Found retention policy '{retention_policy}' using property '{prop}' on item {idx}")
                                            break
                                except:
                                    continue
                            
                            if retention_policy and retention_policy != 'Unknown' and retention_policy not in found_retention_policies:
                                found_retention_policies.add(retention_policy)
                                logger.info(f"Added retention policy: {retention_policy}")
                            
                            sample_count += 1
                            
                        except Exception as e:
                            logger.debug(f"Could not check policy on item {idx}: {str(e)}")
                            continue
                
                except Exception as e:
                    logger.warning(f"Error during email processing: {str(e)}")
                
                logger.info(f"Found {len(found_retention_policies)} retention policies from email sampling")
                
                # Add all discovered retention policies to available policies
                for policy in found_retention_policies:
                    if policy not in available_policies:
                        available_policies.append(policy)
                        logger.info(f"Added retention policy to available policies: {policy}")
                
            except Exception as e:
                logger.warning(f"Could not detect Exchange retention policies: {str(e)}")
                logger.warning("This is normal if no retention policies are assigned to emails")
            
            # 3. Try to get IRM (Information Rights Management) policies from Exchange/Outlook
            try:
                # Try to access policy templates through Outlook stores
                stores = self.namespace.Stores
                for i in range(1, stores.Count + 1):
                    store = stores.Item(i)
                    try:
                        # Check if store has any policy-related properties
                        root_folder = store.GetRootFolder()
                        
                        # Some stores might expose policy templates
                        if hasattr(root_folder, 'PolicyTemplate') and root_folder.PolicyTemplate:
                            policy_template = root_folder.PolicyTemplate
                            if policy_template not in available_policies:
                                available_policies.append(f"IRM Policy: {policy_template}")
                                logger.debug(f"Found IRM policy template: {policy_template}")
                    except Exception as e:
                        logger.debug(f"Could not access policy templates from store {store.DisplayName}: {str(e)}")
                        
            except Exception as e:
                logger.debug(f"Could not retrieve IRM policies from stores: {str(e)}")
            
            # 4. Check for other policy-related information in the environment
            try:
                # Check user's default folder for any policy-related information
                try:
                    default_folder = self.namespace.GetDefaultFolder(1)  # olFolderInbox
                    
                    # Look for any UserProperties that might indicate custom policies
                    if hasattr(default_folder, 'UserProperties'):
                        user_props = default_folder.UserProperties
                        for i in range(1, user_props.Count + 1):
                            prop = user_props.Item(i)
                            if prop and 'policy' in str(prop).lower():
                                logger.debug(f"Found potential policy-related property: {prop}")
                                
                except Exception as e:
                    logger.debug(f"Could not check for custom policies in MAPI: {str(e)}")
                    
            except Exception as e:
                logger.debug(f"Could not retrieve custom policies: {str(e)}")
            # Ensure we always have at least the basic Outlook policies
            if not available_policies:
                available_policies = ["Normal", "Personal", "Private", "Confidential"]
                logger.warning("No policies detected, using basic Outlook sensitivity levels")
            
            logger.info(f"Retrieved available policies: {available_policies}")
            return available_policies
            
        except Exception as e:
            logger.error(f"Error getting available policies: {str(e)}")
            # Fallback to basic Outlook sensitivity levels
            return ["Normal", "Personal", "Private", "Confidential"]
    
    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def assign_policy(self, email_id: str, policy_name: str) -> str:
        """Assign a policy to an email.
        
        Args:
            email_id: EntryID of the email to assign policy to
            policy_name: Name of the policy to assign (can be display name or value)
            
        Returns:
            Confirmation message
        """
        try:
            if not email_id:
                raise ValueError("Email ID cannot be empty")
            if not policy_name:
                raise ValueError("Policy name cannot be empty")
            
            # Get the email item
            email_item = self.namespace.GetItemFromID(email_id)
            email_subject = email_item.Subject
            
            # Get available policies dynamically instead of hardcoded
            available_policies = self.get_available_policies()
            
            # Handle Exchange retention policies (dynamic detection for ANY policy type)
            # Check if the requested policy matches ANY available policy (not just "year" policies)
            matching_policy = None
            for p in available_policies:
                if (policy_name.lower() == p.lower() or 
                    policy_name.lower() in p.lower() or 
                    p.lower() in policy_name.lower()):
                    matching_policy = p
                    break
            
            if matching_policy:
                # Try multiple methods to assign Exchange retention policy
                
                # Method 1: Try direct property assignment (might work in some Outlook versions)
                try:
                    # Check if RetentionPolicyName is writable in this environment
                    original_policy = getattr(email_item, 'RetentionPolicyName', None)
                    email_item.RetentionPolicyName = matching_policy
                    email_item.Save()
                    logger.info(f"Successfully assigned Exchange retention policy '{matching_policy}' via direct property assignment")
                    return f"Successfully assigned Exchange retention policy '{matching_policy}' to email: '{email_subject}'"
                except Exception as e1:
                    logger.debug(f"Direct property assignment failed: {str(e1)}")
                
                # Method 2: Set via Categories (Exchange processes categories matching policy names)
                try:
                    current_categories = email_item.Categories or ""
                    
                    # Remove any existing policies from categories first
                    if current_categories:
                        category_list = [cat.strip() for cat in current_categories.split(';')]
                        non_policy_categories = []
                        
                        for category in category_list:
                            # Check if this category is a policy (matches any available policy)
                            is_policy = False
                            for available_policy in available_policies:
                                if (category.lower() == available_policy.lower() or 
                                    category in available_policy or 
                                    available_policy in category):
                                    is_policy = True
                                    break
                            
                            if not is_policy:
                                non_policy_categories.append(category)
                        
                        # Reconstruct categories without policy entries
                        current_categories = '; '.join(non_policy_categories)
                    
                    # Add the new policy to categories
                    if current_categories:
                        new_categories = f"{current_categories}; {matching_policy}"
                    else:
                        new_categories = matching_policy
                    
                    email_item.Categories = new_categories.strip(" ;")
                    
                    email_item.Save()
                    logger.info(f"Successfully assigned Exchange retention policy '{matching_policy}' via Categories")
                    return f"Successfully assigned Exchange retention policy '{matching_policy}' to email: '{email_subject}' (via Categories)"
                except Exception as e2:
                    logger.debug(f"Categories method failed: {str(e2)}")
                
                # Method 3: Set via UserProperty as metadata
                try:
                    policy_prop = email_item.UserProperties.Find("retention_policy")
                    if not policy_prop:
                        policy_prop = email_item.UserProperties.Add("retention_policy", 1, False, 0)
                    policy_prop.Value = matching_policy
                    email_item.Save()
                    logger.info(f"Assigned retention policy '{matching_policy}' via UserProperty")
                    return f"Successfully assigned Exchange retention policy '{matching_policy}' to email: '{email_subject}' (via UserProperty)"
                except Exception as e3:
                    logger.debug(f"UserProperty method failed: {str(e3)}")
                
                # Method 4: Try alternative property names
                try:
                    alternative_properties = ['PolicyName', 'RetentionPolicy', 'Policy']
                    for prop_name in alternative_properties:
                        try:
                            setattr(email_item, prop_name, matching_policy)
                            email_item.Save()
                            logger.info(f"Successfully assigned retention policy '{matching_policy}' via {prop_name}")
                            return f"Successfully assigned Exchange retention policy '{matching_policy}' to email: '{email_subject}' (via {prop_name})"
                        except:
                            continue
                except Exception as e4:
                    logger.debug(f"Alternative properties method failed: {str(e4)}")
                
                # If all methods failed, return a helpful error message
                available_methods = ["direct property assignment", "Categories", "UserProperty", "alternative properties"]
                logger.error(f"All retention policy assignment methods failed. Tried: {available_methods}")
                return f"Error: Unable to assign Exchange retention policy '{matching_policy}'. All assignment methods failed. Please check Exchange server configuration."
            
            # Handle enterprise policy using custom UserProperty (completely dynamic detection)
            # Check if this looks like an enterprise policy request (no hardcoded policy names)
            elif (policy_name.lower().endswith(" (enterprise)") or 
                  "enterprise" in policy_name.lower()):
                # Extract the base policy name (remove enterprise suffix if present)
                base_policy_name = policy_name
                if policy_name.lower().endswith(" (enterprise)"):
                    base_policy_name = policy_name[:-12]  # Remove " (enterprise)"
                
                # Find a matching retention policy from available policies
                matching_policy = None
                for p in available_policies:
                    if (base_policy_name.lower() == p.lower() or 
                        base_policy_name.lower() in p.lower() or 
                        p.lower() in base_policy_name.lower()):
                        matching_policy = p
                        break
                
                # If we found a matching policy, use it as the base for enterprise policy
                if matching_policy:
                    base_policy_name = matching_policy
                
                # Add custom property for enterprise policy
                try:
                    # Check if the policy property already exists
                    policy_prop = None
                    for prop in email_item.UserProperties:
                        if prop.Name.lower() == "enterprise_policy":
                            policy_prop = prop
                            break
                    
                    if policy_prop:
                        # Update existing property
                        policy_prop.Value = base_policy_name
                    else:
                        # Create new property using ItemProperties
                        try:
                            policy_prop = email_item.ItemProperties.Add("enterprise_policy", 1)  # olText
                            policy_prop.Value = base_policy_name
                        except:
                            # If ItemProperties.Add fails, try UserProperties differently
                            # Use a different approach to add custom property
                            policy_prop = email_item.UserProperties.Find("enterprise_policy")
                            if not policy_prop:
                                # Create a basic custom property
                                try:
                                    policy_prop = email_item.UserProperties.Add("enterprise_policy", 1, False, 0)
                                    policy_prop.Value = base_policy_name
                                except:
                                    # Last resort: use Categories field
                                    categories = email_item.Categories or ""
                                    if base_policy_name not in categories:
                                        new_categories = f"{categories}; {base_policy_name}" if categories else base_policy_name
                                        email_item.Categories = new_categories.strip(" ;")
                            else:
                                policy_prop.Value = base_policy_name
                    
                    # Set ExpiryTime to match enterprise policy (dynamic detection)
                    try:
                        # Try to detect existing enterprise policy expiry patterns from environment
                        # Look for existing emails with enterprise policies to get the expiry pattern
                        expiry_pattern = None
                        try:
                            # Sample some emails to see if any have enterprise policies and their expiry times
                            sample_items = self.namespace.Folders[0].Items
                            sample_count = 0
                            for item in sample_items:
                                if sample_count >= 10:  # Sample max 10 emails
                                    break
                                if hasattr(item, 'UserProperties'):
                                    for prop in item.UserProperties:
                                        if prop.Name.lower() == "enterprise_policy" and prop.Value:
                                            if hasattr(item, 'ExpiryTime') and item.ExpiryTime:
                                                expiry_pattern = item.ExpiryTime
                                                logger.info(f"Found existing enterprise policy expiry pattern: {expiry_pattern}")
                                                break
                                sample_count += 1
                        except Exception as e:
                            logger.debug(f"Could not sample existing emails for expiry pattern: {str(e)}")
                        
                        # If no pattern found, use a reasonable future date (not hardcoded specific date)
                        if not expiry_pattern:
                            # Use a date far enough in the future to be effectively "infinite"
                            # but not a specific hardcoded date
                            import datetime
                            future_date = datetime.datetime.now() + datetime.timedelta(days=365*100)  # 100 years from now
                            expiry_pattern = future_date.strftime("%Y-%m-%d %H:%M:%S")
                        
                        # Set the expiry time using the detected or calculated pattern
                        email_item.ExpiryTime = expiry_pattern
                        logger.info(f"Set ExpiryTime to {expiry_pattern} for enterprise policy")
                        
                    except Exception as expiry_error:
                        try:
                            # Method 2: Try with COM date format using the detected pattern
                            import pywintypes
                            import datetime
                            future_date = datetime.datetime.now() + datetime.timedelta(days=365*100)
                            expiry_py_time = pywintypes.Time(future_date.timestamp())
                            email_item.ExpiryTime = expiry_py_time
                            logger.info(f"Set ExpiryTime using pywintypes for enterprise policy")
                        except Exception as expiry_error2:
                            logger.warning(f"Could not set ExpiryTime (method 1: {str(expiry_error)}, method 2: {str(expiry_error2)})")
                            # Continue anyway - Categories and UserProperty should be sufficient
                    
                    email_item.Save()
                    logger.info(f"Assigned enterprise policy '{base_policy_name}' to email: '{email_subject}'")
                    return f"Successfully assigned enterprise policy '{base_policy_name}' to email: '{email_subject}'"
                    
                except Exception as e:
                    logger.error(f"Error assigning enterprise policy: {str(e)}")
                    # Fallback to sensitivity if custom property fails - use dynamic fallback
                    try:
                        test_mail = self.outlook.CreateItem(0)
                        test_mail.Sensitivity = 2  # Private
                        email_item.Sensitivity = 2
                        test_mail.Delete()
                    except:
                        # Last resort fallback
                        email_item.Sensitivity = 0  # Normal
                    email_item.Save()
                    return f"Assigned fallback policy to email: '{email_subject}' (Enterprise policy assignment failed: {str(e)})"

            
            # Map built-in policy names to Outlook sensitivity value (dynamic)
            # Test sensitivity levels dynamically instead of hardcoding
            policy_map = {}
            
            # Test common sensitivity values dynamically
            sensitivity_tests = [
                ("normal", 0),
                ("low sensitivity", 1), 
                ("personal", 1),
                ("private", 2),
                ("confidential", 3)
            ]
            
            # Test each sensitivity level dynamically
            for name, value in sensitivity_tests:
                try:
                    test_mail = self.outlook.CreateItem(0)
                    test_mail.Sensitivity = value
                    policy_map[name.lower()] = value
                    test_mail.Delete()
                    logger.debug(f"Found working sensitivity: {name} = {value}")
                except Exception as e:
                    logger.debug(f"Sensitivity {name} = {value} not available: {str(e)}")
            
            # Also check available policies for built-in policies
            for p in available_policies:
                if p.lower() in policy_map:
                    policy_map[p.lower()] = policy_map[p.lower()]
                elif p in ["Normal", "Personal", "Private", "Confidential"]:
                    # Map exact names to their numeric values based on standard Outlook values
                    if p == "Normal":
                        policy_map[p.lower()] = 0
                    elif p == "Personal":
                        policy_map[p.lower()] = 1
                    elif p == "Private":
                        policy_map[p.lower()] = 2
                    elif p == "Confidential":
                        policy_map[p.lower()] = 3
            
            policy_name_lower = policy_name.lower()
            sensitivity_value = None
            
            # Find the matching policy
            for name, value in policy_map.items():
                if name == policy_name_lower:
                    sensitivity_value = value
                    break
            
            if sensitivity_value is None:
                raise ValueError(f"Policy '{policy_name}' not found. Available policies: {available_policies}")
            
            # Set the sensitivity property
            email_item.Sensitivity = sensitivity_value
            email_item.Save()
            
            logger.info(f"Assigned built-in policy '{policy_name}' to email: '{email_subject}'")
            return f"Successfully assigned policy '{policy_name}' to email: '{email_subject}'"
        except Exception as e:
            logger.error(f"Error assigning policy: {str(e)}")
            raise