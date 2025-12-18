"""
Folder operations for Outlook session management.

This module provides folder-related operations such as creation, deletion, moving, and retrieval.
"""

import logging
from typing import Optional, Any, Dict, List, Tuple
from datetime import datetime

from ..utils import retry_on_com_error, OutlookFolderType
from .exceptions import FolderNotFoundError, OperationFailedError, InvalidParameterError

logger = logging.getLogger(__name__)


class FolderOperations:
    """Handles all folder-related operations for Outlook."""

    def __init__(self, session_manager):
        """Initialize with a session manager instance."""
        self.session_manager = session_manager

    def get_folder(self, folder_name: Optional[str] = None):
        """Get specified folder or default inbox."""
        # Handle string "null" as well as actual None
        if not folder_name or folder_name == "null" or folder_name.lower() == "inbox":
            folder = self.session_manager.outlook_namespace.GetDefaultFolder(OutlookFolderType.INBOX)
            return folder
        elif folder_name.lower() == "sent items" or folder_name.lower() == "sent":
            folder = self.session_manager.outlook_namespace.GetDefaultFolder(OutlookFolderType.SENT_MAIL)
            return folder
        elif folder_name.lower() == "deleted items" or folder_name.lower() == "trash":
            folder = self.session_manager.outlook_namespace.GetDefaultFolder(OutlookFolderType.DELETED_ITEMS)
            return folder
        elif folder_name.lower() == "drafts":
            folder = self.session_manager.outlook_namespace.GetDefaultFolder(OutlookFolderType.DRAFTS)
            return folder
        elif folder_name.lower() == "outbox":
            folder = self.session_manager.outlook_namespace.GetDefaultFolder(OutlookFolderType.OUTBOX)
            return folder
        elif folder_name.lower() == "calendar":
            folder = self.session_manager.outlook_namespace.GetDefaultFolder(OutlookFolderType.CALENDAR)
            return folder
        elif folder_name.lower() == "contacts":
            folder = self.session_manager.outlook_namespace.GetDefaultFolder(OutlookFolderType.CONTACTS)
            return folder
        elif folder_name.lower() == "tasks":
            folder = self.session_manager.outlook_namespace.GetDefaultFolder(OutlookFolderType.TASKS)
            return folder
        else:
            folder = self._get_folder_by_name(folder_name)
            return folder

    def _get_folder_by_name(self, folder_name: str):
        """Find folder by name in folder hierarchy, supporting nested paths and mailbox-specific paths."""
        try:
            # Handle nested folder paths (e.g., "Parent Folder/Child Folder" or "mailbox@domain.com/Inbox/Folder")
            if "/" in folder_name or "\\" in folder_name:
                # Use forward slash as path separator, but also support backslash
                path_parts = folder_name.replace("\\", "/").split("/")
                current_folder = None

                # Check if first part looks like an email address (mailbox-specific path)
                if "@" in path_parts[0] and "." in path_parts[0]:
                    # This is a mailbox-specific path like "user@company.com/Inbox/Folder"
                    mailbox_name = path_parts[0]

                    # Find the mailbox folder
                    for folder in self.session_manager.outlook_namespace.Folders:
                        if folder.Name == mailbox_name:
                            current_folder = folder
                            break

                    if not current_folder:
                        raise FolderNotFoundError(f"Mailbox '{mailbox_name}' not found")

                    # Navigate through the remaining path parts
                    remaining_parts = path_parts[1:]
                else:
                    # This is a regular path like "Inbox/Folder" or "Parent Folder/Child Folder"
                    # Start with the top-level folders
                    for folder in self.session_manager.outlook_namespace.Folders:
                        if folder.Name == path_parts[0]:
                            current_folder = folder
                            break

                    if not current_folder:
                        raise FolderNotFoundError(f"Top-level folder '{path_parts[0]}' not found")

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
                        raise FolderNotFoundError(f"Folder '{part}' not found in '{current_folder.Name}'")

                return current_folder
            else:
                # Original logic for single folder names
                for folder in self.session_manager.outlook_namespace.Folders:
                    if folder.Name == folder_name:
                        return folder
                    for subfolder in folder.Folders:
                        if subfolder.Name == folder_name:
                            return subfolder
                raise FolderNotFoundError(f"Folder '{folder_name}' not found")
        except Exception as e:
            logger.error(f"Error finding folder: {str(e)}")
            if isinstance(e, FolderNotFoundError):
                raise
            raise OperationFailedError(f"Error finding folder '{folder_name}': {str(e)}")

    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def create_folder(self, folder_name: str, parent_folder_name: Optional[str] = None) -> str:
        """Create a new folder in the specified parent folder.
        
        Args:
            folder_name: Name of the folder to create
            parent_folder_name: Name of the parent folder (optional, defaults to Inbox)
            
        Returns:
            Success message
        """
        if not folder_name or not isinstance(folder_name, str):
            raise InvalidParameterError("Folder name must be a non-empty string")
            
        try:
            parent_folder = self.get_folder(parent_folder_name)
            
            # Check if folder already exists
            for folder in parent_folder.Folders:
                if folder.Name == folder_name:
                    return f"Folder '{folder_name}' already exists in '{parent_folder.Name}'"
            
            # Create the new folder
            new_folder = parent_folder.Folders.Add(folder_name)
            logger.info(f"Created folder '{folder_name}' in '{parent_folder.Name}'")
            return f"Folder '{folder_name}' created successfully in '{parent_folder.Name}'"
            
        except Exception as e:
            error_msg = f"Error creating folder '{folder_name}': {str(e)}"
            logger.error(error_msg)
            raise OperationFailedError(error_msg)

    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def remove_folder(self, folder_name: str) -> str:
        """Remove an existing folder.
        
        Args:
            folder_name: Name or path of the folder to remove
            
        Returns:
            Success message
        """
        if not folder_name or not isinstance(folder_name, str):
            raise InvalidParameterError("Folder name must be a non-empty string")
            
        try:
            folder = self._get_folder_by_name(folder_name)
            
            # Check if it's a default folder
            folder_path = folder.FolderPath if hasattr(folder, 'FolderPath') else folder.Name
            if self._is_default_folder(folder_path):
                raise OperationFailedError(f"Cannot remove default folder '{folder_name}'")
            
            # Get parent folder for the success message
            parent_folder = folder.Parent
            folder_name_only = folder.Name
            
            # Delete the folder
            folder.Delete()
            
            logger.info(f"Removed folder '{folder_name_only}' from '{parent_folder.Name}'")
            return f"Folder '{folder_name_only}' removed successfully from '{parent_folder.Name}'"
            
        except Exception as e:
            error_msg = f"Error removing folder '{folder_name}': {str(e)}"
            logger.error(error_msg)
            if isinstance(e, (FolderNotFoundError, OperationFailedError)):
                raise
            raise OperationFailedError(error_msg)

    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def move_folder(self, source_folder_path: str, target_parent_path: str) -> str:
        """Move a folder to a different parent folder.
        
        Args:
            source_folder_path: Path to the source folder
            target_parent_path: Path to the target parent folder
            
        Returns:
            Success message
        """
        if not source_folder_path or not isinstance(source_folder_path, str):
            raise InvalidParameterError("Source folder path must be a non-empty string")
        if not target_parent_path or not isinstance(target_parent_path, str):
            raise InvalidParameterError("Target parent path must be a non-empty string")
            
        try:
            # Get source folder
            source_folder = self._get_folder_by_name(source_folder_path)
            
            # Get target parent folder
            target_parent = self._get_folder_by_name(target_parent_path)
            
            # Check if it's a default folder (cannot move default folders)
            source_folder_path_attr = source_folder.FolderPath if hasattr(source_folder, 'FolderPath') else source_folder.Name
            if self._is_default_folder(source_folder_path_attr):
                raise OperationFailedError(f"Cannot move default folder '{source_folder_path}'")
            
            # Move the folder
            source_folder.MoveTo(target_parent)
            
            logger.info(f"Moved folder '{source_folder.Name}' to '{target_parent.Name}'")
            return f"Folder '{source_folder.Name}' moved successfully to '{target_parent.Name}'"
            
        except Exception as e:
            error_msg = f"Error moving folder '{source_folder_path}' to '{target_parent_path}': {str(e)}"
            logger.error(error_msg)
            if isinstance(e, (FolderNotFoundError, OperationFailedError)):
                raise
            raise OperationFailedError(error_msg)

    def get_folder_list(self):
        """Get list of all folders."""
        try:
            folders = []
            for folder in self.session_manager.outlook_namespace.Folders:
                folders.append(folder)
                # Also add subfolders
                self._add_subfolders(folder, folders)
            return folders
        except Exception as e:
            logger.error(f"Error getting folder list: {str(e)}")
            raise OperationFailedError(f"Error getting folder list: {str(e)}")

    def _add_subfolders(self, folder, folders_list):
        """Recursively add subfolders to the list."""
        try:
            for subfolder in folder.Folders:
                folders_list.append(subfolder)
                self._add_subfolders(subfolder, folders_list)
        except Exception as e:
            logger.warning(f"Error accessing subfolders of {folder.Name}: {str(e)}")

    def _is_default_folder(self, folder_path: str) -> bool:
        """Check if a folder is a default Outlook folder."""
        default_folders = [
            "Inbox", "Sent Items", "Deleted Items", "Drafts", 
            "Outbox", "Junk Email", "Archive", "Conversation History"
        ]
        # Extract folder name from path
        folder_name = folder_path.split("\\")[-1] if "\\" in folder_path else folder_path
        return folder_name in default_folders

    def get_folder_emails(self, folder_name: str = "Inbox", max_emails: int = 100) -> Tuple[List[Dict[str, Any]], str]:
        """
        Get emails from a folder with pagination support.
        
        Args:
            folder_name: Name of the folder to get emails from
            max_emails: Maximum number of emails to return
        
        Returns:
            Tuple of (list of email dictionaries, status message)
        """
        try:
            # Validate input parameters
            if not folder_name or not folder_name.strip():
                return [], f"Error: Invalid folder name: {folder_name}"
            if max_emails < 1:
                return [], f"Error: Invalid max_emails value: {max_emails}"
        except Exception as e:
            logger.error(f"Validation error in get_folder_emails: {e}")
            return [], f"Error: Invalid parameters: {e}"

        try:
            folder = self.get_folder(folder_name)
            if not folder:
                return [], f"Error: Folder '{folder_name}' not found"
            
            logger.info(f"Getting emails from folder '{folder_name}' with limit {max_emails}")
            
            # Get items from the folder
            items = list(folder.Items)
            
            if not items:
                return [], f"No emails found in '{folder_name}'"
            
            # Sort by received time (newest first)
            try:
                items.sort(key=lambda x: x.ReceivedTime if hasattr(x, 'ReceivedTime') and x.ReceivedTime else datetime.min, reverse=True)
            except Exception as e:
                logger.warning(f"Error sorting emails: {e}")
            
            # Limit the number of emails
            limited_items = items[:max_emails]
            
            # Convert to cache format
            email_list = []
            for item in limited_items:
                try:
                    from ..email_search.search_common import extract_email_info
                    email_data = extract_email_info(item)
                    if email_data:
                        from ..shared import add_email_to_cache
                        add_email_to_cache(email_data["entry_id"], email_data)
                        email_list.append(email_data)
                except Exception as e:
                    logger.warning(f"Failed to process email: {e}")
                    continue
            
            if not email_list:
                return [], f"No valid emails found in '{folder_name}'"
            
            message = f"Found {len(email_list)} emails in '{folder_name}'"
            return email_list, message
            
        except Exception as e:
            error_msg = f"Error getting emails from folder: {e}"
            logger.error(error_msg)
            return [], f"Error: {error_msg}"


def list_folders():
    """Get list of all folder names.
    
    Returns:
        List of folder names
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            folders = folder_ops.get_folder_list()
            return [folder.Name for folder in folders]
    except Exception as e:
         logger.error(f"Error getting folder list: {str(e)}")
         return []


def create_folder(folder_name: str, parent_folder_name: Optional[str] = None) -> str:
    """Create a new folder in the specified parent folder.
    
    Args:
        folder_name: Name of the folder to create
        parent_folder_name: Name of the parent folder (optional, defaults to Inbox)
        
    Returns:
        Success message
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            return folder_ops.create_folder(folder_name, parent_folder_name)
    except Exception as e:
        logger.error(f"Error creating folder: {str(e)}")
        return f"Error: {str(e)}"


def remove_folder(folder_name: str) -> str:
    """Remove an existing folder.
    
    Args:
        folder_name: Name or path of the folder to remove
        
    Returns:
        Success message
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            return folder_ops.remove_folder(folder_name)
    except Exception as e:
        logger.error(f"Error removing folder: {str(e)}")
        return f"Error: {str(e)}"


def move_folder(source_folder_path: str, target_parent_path: str) -> str:
    """Move a folder to a different parent folder.
    
    Args:
        source_folder_path: Path to the source folder
        target_parent_path: Path to the target parent folder
        
    Returns:
        Success message
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            return folder_ops.move_folder(source_folder_path, target_parent_path)
    except Exception as e:
        logger.error(f"Error moving folder: {str(e)}")
        return f"Error: {str(e)}"


def get_folder_emails(folder_name: str = "Inbox", max_emails: int = 100) -> Tuple[List[Dict[str, Any]], str]:
    """Get emails from a folder with pagination support.
    
    Args:
        folder_name: Name of the folder to get emails from
        max_emails: Maximum number of emails to return
        
    Returns:
        Tuple of (list of email dictionaries, status message)
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            return folder_ops.get_folder_emails(folder_name, max_emails)
    except Exception as e:
        logger.error(f"Error getting folder emails: {str(e)}")
        return [], f"Error: {str(e)}"