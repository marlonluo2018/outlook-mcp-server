"""
Folder operations for Outlook session management.

This module provides folder-related operations such as creation, deletion, moving, and retrieval.
"""

import logging
from typing import Optional, Any, Dict, List, Tuple
from datetime import datetime
import time
from functools import lru_cache

from ..utils import retry_on_com_error, OutlookFolderType
from .exceptions import FolderNotFoundError, OperationFailedError, InvalidParameterError

logger = logging.getLogger(__name__)


class FolderOperations:
    """Handles all folder-related operations for Outlook."""

    def __init__(self, session_manager):
        """Initialize with a session manager instance."""
        self.session_manager = session_manager
        self._folder_cache = {}
        self._cache_timestamp = 0
        self._cache_ttl = 300  # 5 minutes TTL for folder cache

    def _is_cache_valid(self):
        """Check if folder cache is still valid."""
        return time.time() - self._cache_timestamp < self._cache_ttl

    def _get_cached_folder(self, folder_name: str):
        """Get folder from cache if available and valid."""
        if self._is_cache_valid() and folder_name in self._folder_cache:
            logger.debug(f"Folder cache hit for: {folder_name}")
            return self._folder_cache[folder_name]
        return None

    def _cache_folder(self, folder_name: str, folder):
        """Cache folder for future use."""
        self._folder_cache[folder_name] = folder
        self._cache_timestamp = time.time()
        logger.debug(f"Folder cached: {folder_name}")

    def clear_folder_cache(self):
        """Clear the folder cache."""
        self._folder_cache.clear()
        self._cache_timestamp = 0
        logger.info("Folder cache cleared")

    def get_folder(self, folder_name: Optional[str] = None):
        """Get specified folder or default inbox with caching."""
        # Normalize folder name for caching
        cache_key = folder_name.lower() if folder_name else "inbox"
        
        # Try cache first
        cached_folder = self._get_cached_folder(cache_key)
        if cached_folder:
            return cached_folder
        
        # Get folder and cache it
        folder = self._get_folder_internal(folder_name)
        if folder:
            self._cache_folder(cache_key, folder)
        return folder

    def _get_folder_internal(self, folder_name: Optional[str] = None):
        """Internal method to get folder without caching."""
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

                # Navigate through the remaining path parts - optimized search
                for part in remaining_parts:
                    found = False
                    # Try direct access first (much faster for well-known folders)
                    try:
                        current_folder = current_folder.Folders[part]
                        found = True
                    except Exception:
                        # Fall back to iteration if direct access fails
                        for subfolder in current_folder.Folders:
                            if subfolder.Name == part:
                                current_folder = subfolder
                                found = True
                                break
                    if not found:
                        raise FolderNotFoundError(f"Folder '{part}' not found in '{current_folder.Name}'")

                return current_folder
            else:
                # Original logic for single folder names - optimized
                # Try direct access first for common folders
                try:
                    return self.session_manager.outlook_namespace.Folders[folder_name]
                except Exception:
                    pass
                
                # Fall back to iteration
                for folder in self.session_manager.outlook_namespace.Folders:
                    if folder.Name == folder_name:
                        return folder
                    # Try direct access for subfolders
                    try:
                        return folder.Folders[folder_name]
                    except Exception:
                        # Fall back to iteration for subfolders
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

    def get_folder_emails(self, folder_name: str = "Inbox", max_emails: int = 100, fast_mode: bool = True, days_filter: int = None) -> Tuple[List[Dict[str, Any]], str]:
        """
        Get emails from a folder with pagination support - optimized for performance.
        
        Args:
            folder_name: Name of the folder to get emails from
            max_emails: Maximum number of emails to return
            fast_mode: If True, use minimal extraction for better performance
            days_filter: Number of days to filter by (None for number-based loading)
        
        Returns:
            Tuple of (list of email dictionaries, status message)
        """
        start_time = time.time()
        
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
            # Get folder with caching
            folder = self.get_folder(folder_name)
            if not folder:
                return [], f"Error: Folder '{folder_name}' not found"
            
            logger.info(f"Getting emails from folder '{folder_name}' with limit {max_emails}, fast_mode={fast_mode}")
            
            # Use server-side filtering with Restrict method for better performance
            from datetime import datetime, timedelta
            
            # Determine optimal filtering strategy based on parameters
            items = []
            filter_time = time.time()
            
            if days_filter is None:
                # Number-based loading: get items without date filtering, but ensure we get newest first
                logger.info(f"Number-based loading: getting up to {max_emails} items without date filter")
                
                # Use a much smaller date range initially for better performance
                # Start with 7 days, then expand gradually if needed
                days_to_try = [7, 14, 30, 60, 90]
                items = []
                
                for days in days_to_try:
                    date_limit = datetime.now() - timedelta(days=days)
                    date_filter = f"@SQL=urn:schemas:httpmail:datereceived >= '{date_limit.strftime('%Y-%m-%d')}'"
                    
                    try:
                        filtered_items = folder.Items.Restrict(date_filter)
                        if filtered_items.Count > 0:
                            # Use a more efficient approach - get only what we need
                            # Instead of converting entire collection to list, iterate efficiently
                            temp_items = []
                            count = 0
                            # Use GetLast/GetPrevious for newest-first order (better performance)
                            item = filtered_items.GetLast()
                            while item and count < max_emails * 2:  # Get 2x to account for filtering
                                temp_items.append(item)
                                count += 1
                                item = filtered_items.GetPrevious()
                            
                            items = temp_items
                            logger.info(f"{days}-day filter returned {len(items)} items in {time.time() - filter_time:.2f}s")
                            
                            # If we got enough items, break out of the loop
                            if len(items) >= max_emails:
                                break
                        else:
                            continue  # Try next larger time window
                    except Exception as e:
                        logger.warning(f"Restrict method failed for {days} days: {e}")
                        continue
                
                # If no items found with date filtering, try reverse indexing as fallback
                if not items:
                    logger.info("No items found with date filtering, trying reverse indexing fallback")
                    try:
                        # Get items in reverse order (newest first) using a more efficient approach
                        total_count = folder.Items.Count
                        if total_count > 0:
                            # Start from the end (newest items) and work backwards efficiently
                            # Use GetLast/GetPrevious for better performance
                            items = []
                            item = folder.Items.GetLast()
                            count = 0
                            while item and count < max_emails * 2:  # Get 2x to be safe
                                items.append(item)
                                count += 1
                                item = folder.Items.GetPrevious()
                            logger.info(f"Retrieved {len(items)} items using GetLast/GetPrevious in {time.time() - filter_time:.2f}s")
                        else:
                            items = []
                    except Exception as final_e:
                        logger.error(f"All fallback methods failed: {final_e}")
                        items = []
            
            else:
                # Time-based loading: use date filtering (existing logic)
                # For small requests (≤50), use very recent filter for speed
                if max_emails <= 50:
                    date_limit = datetime.now() - timedelta(days=7)  # Only 7 days for small requests
                    date_filter = f"@SQL=urn:schemas:httpmail:datereceived >= '{date_limit.strftime('%Y-%m-%d')}'"
                    
                    try:
                        filtered_items = folder.Items.Restrict(date_filter)
                        if filtered_items.Count > 0:
                            # Use efficient iteration instead of list conversion
                            items = []
                            count = 0
                            item = filtered_items.GetFirst()
                            while item and count < max_emails * 2:
                                items.append(item)
                                count += 1
                                item = filtered_items.GetNext()
                            logger.info(f"7-day filter returned {len(items)} items in {time.time() - filter_time:.2f}s")
                            # If we got enough items, use them
                            if len(items) >= max_emails:
                                pass  # We have enough
                            else:
                                # Not enough recent items, fall back to larger filter
                                logger.info(f"7-day filter only returned {len(items)} items, expanding to 30 days")
                                date_limit = datetime.now() - timedelta(days=30)
                                date_filter = f"@SQL=urn:schemas:httpmail:datereceived >= '{date_limit.strftime('%Y-%m-%d')}'"
                                filtered_items = folder.Items.Restrict(date_filter)
                                # Use efficient iteration
                                items = []
                                count = 0
                                item = filtered_items.GetFirst()
                                while item and count < max_emails * 2:
                                    items.append(item)
                                    count += 1
                                    item = filtered_items.GetNext()
                                logger.info(f"30-day filter returned {len(items)} items in {time.time() - filter_time:.2f}s")
                        else:
                            items = []
                    except Exception as e:
                        logger.warning(f"Restrict method failed: {e}, using sorted list approach")
                        # Use sorted list approach to get newest emails first
                        items = []
                        try:
                            all_items = list(folder.Items)
                            # Sort by received time (newest first) before limiting
                            all_items.sort(key=lambda x: x.ReceivedTime if hasattr(x, 'ReceivedTime') and x.ReceivedTime else datetime.min, reverse=True)
                            items = all_items[:max_emails * 2]  # Get 2x to account for filtering
                        except Exception as inner_e:
                            logger.error(f"Sorted list approach failed: {inner_e}")
                            # Final fallback - try GetFirst/GetNext
                            try:
                                item = folder.Items.GetFirst()
                                count = 0
                                while item and count < max_emails * 2:
                                    items.append(item)
                                    item = folder.Items.GetNext()
                                    count += 1
                            except Exception as final_e:
                                logger.error(f"All fallback methods failed: {final_e}")
                                items = []
                else:
                    # For larger requests, use 30-day filter first
                    date_limit = datetime.now() - timedelta(days=30)
                    date_filter = f"@SQL=urn:schemas:httpmail:datereceived >= '{date_limit.strftime('%Y-%m-%d')}'"
                    
                    try:
                        filtered_items = folder.Items.Restrict(date_filter)
                        if filtered_items.Count > 0:
                            # Use efficient iteration instead of list conversion
                            items = []
                            count = 0
                            item = filtered_items.GetFirst()
                            while item and count < max_emails * 2:
                                items.append(item)
                                count += 1
                                item = filtered_items.GetNext()
                            logger.info(f"30-day filter returned {len(items)} items in {time.time() - filter_time:.2f}s")
                        else:
                            items = []
                    except Exception as e:
                        logger.warning(f"Restrict method failed: {e}, falling back to sorted list approach")
                        # Use sorted list approach to get newest emails first
                        items = []
                        try:
                            all_items = list(folder.Items)
                            # Sort by received time (newest first) before limiting
                            all_items.sort(key=lambda x: x.ReceivedTime if hasattr(x, 'ReceivedTime') and x.ReceivedTime else datetime.min, reverse=True)
                            items = all_items[:max_emails * 2]  # Get 2x to account for filtering
                        except Exception as inner_e:
                            logger.error(f"Sorted list approach failed: {inner_e}")
                            # Final fallback - try GetFirst/GetNext
                            try:
                                item = folder.Items.GetFirst()
                                count = 0
                                while item and count < max_emails * 2:
                                    items.append(item)
                                    item = folder.Items.GetNext()
                                    count += 1
                            except Exception as final_e:
                                logger.error(f"All fallback methods failed: {final_e}")
                                items = []
            
            if not items:
                return [], f"No emails found in '{folder_name}'"
            
            # Quick sort by received time (newest first) - only sort what we need
            sort_time = time.time()
            try:
                # Only sort the items we actually need, not the entire collection
                items.sort(key=lambda x: x.ReceivedTime if hasattr(x, 'ReceivedTime') and x.ReceivedTime else datetime.min, reverse=True)
                logger.info(f"Sorting completed in {time.time() - sort_time:.2f}s")
            except Exception as e:
                logger.warning(f"Error sorting emails: {e}")
            
            # Limit the number of emails
            limited_items = items[:max_emails]
            
            # Batch process emails for better performance
            extraction_time = time.time()
            email_list = []
            
            # Import extraction functions once, outside the loop
            if fast_mode:
                from ..email_search.search_common import extract_email_info_minimal
                extractor = extract_email_info_minimal
            else:
                from ..email_search.search_common import extract_email_info
                extractor = extract_email_info
            
            # Process in batches with progress indication
            batch_size = 50 if fast_mode else 25  # Smaller batches for full extraction
            total_items = len(limited_items)
            
            for i in range(0, total_items, batch_size):
                batch = limited_items[i:i + batch_size]
                batch_start = time.time()
                
                for item in batch:
                    try:
                        email_data = extractor(item)
                        if email_data and (fast_mode and email_data.get("entry_id") or not fast_mode):
                            email_list.append(email_data)
                    except Exception as e:
                        logger.warning(f"Failed to process email: {e}")
                        continue
                
                # Log progress for large batches
                if total_items > 100 and (i + batch_size) % 100 == 0:
                    progress = (i + batch_size) / total_items * 100
                    logger.info(f"Progress: {progress:.1f}% ({i + batch_size}/{total_items} items processed)")
            
            logger.info(f"Email extraction completed in {time.time() - extraction_time:.2f}s")
            
            if not email_list:
                return [], f"No valid emails found in '{folder_name}'"
            
            # Use unified cache loading workflow for consistent cache management
            cache_time = time.time()
            from ..email_search.search_common import unified_cache_load_workflow
            success = unified_cache_load_workflow(email_list, f"get_folder_emails({folder_name})")
            
            if success:
                logger.info(f"Unified cache workflow completed in {time.time() - cache_time:.2f}s")
            else:
                logger.warning("Unified cache workflow failed")
            
            total_time = time.time() - start_time
            message = f"Found {len(email_list)} emails in '{folder_name}' (completed in {total_time:.2f}s)"
            
            # Log performance metrics
            logger.info(f"Performance: Folder='{folder_name}', Emails={len(email_list)}, TotalTime={total_time:.2f}s, "
                       f"FilterTime={filter_time - start_time:.2f}s, SortTime={sort_time - filter_time:.2f}s, "
                       f"ExtractTime={extraction_time - sort_time:.2f}s, CacheTime={cache_time - extraction_time:.2f}s")
            
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


def get_folder_emails(folder_name: str = "Inbox", max_emails: int = 100, days_filter: int = None) -> Tuple[List[Dict[str, Any]], str]:
    """Get emails from a folder with pagination support.
    
    Args:
        folder_name: Name of the folder to get emails from
        max_emails: Maximum number of emails to return
        days_filter: Number of days to filter by (None for number-based loading)
        
    Returns:
        Tuple of (list of email dictionaries, status message)
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            return folder_ops.get_folder_emails(folder_name, max_emails, True, days_filter)
    except Exception as e:
        logger.error(f"Error getting folder emails: {str(e)}")
        return [], f"Error: {str(e)}"