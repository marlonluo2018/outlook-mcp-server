from typing import Optional
import logging
import pythoncom
import win32com.client
from backend.shared import configure_logging

# Initialize logging
configure_logging()

# Configure logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

class OutlookSessionManager:
    """Context manager for Outlook COM session handling"""
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.folder = None
        self._connected = False
        
    def __enter__(self):
        """Initialize Outlook COM objects"""
        self._connect()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up Outlook COM objects"""
        self._disconnect()
        
    def _connect(self):
        """Establish COM connection with proper threading"""
        try:
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self._connected = True
        except Exception as e:
            logger.error(f"Connection error: {str(e)}")
            raise RuntimeError("Failed to connect to Outlook") from e
            
    def _disconnect(self):
        """Clean up COM objects"""
        if self._connected:
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.warning(f"Error during disconnect: {str(e)}")
            finally:
                self._connected = False
                
    def is_connected(self) -> bool:
        """Check if the session is still connected"""
        try:
            # Simple operation to test connection
            if self._connected and self.outlook:
                self.outlook.GetNamespace("MAPI").GetDefaultFolder(6).Name
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
                folder = self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
                return folder
            elif folder_name.lower() == "sent items" or folder_name.lower() == "sent":
                folder = self.namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
                return folder
            elif folder_name.lower() == "deleted items" or folder_name.lower() == "trash":
                folder = self.namespace.GetDefaultFolder(3)  # 3 = olFolderDeletedItems
                return folder
            elif folder_name.lower() == "drafts":
                folder = self.namespace.GetDefaultFolder(16)  # 16 = olFolderDrafts
                return folder
            elif folder_name.lower() == "outbox":
                folder = self.namespace.GetDefaultFolder(4)  # 4 = olFolderOutbox
                return folder
            elif folder_name.lower() == "calendar":
                folder = self.namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
                return folder
            elif folder_name.lower() == "contacts":
                folder = self.namespace.GetDefaultFolder(10)  # 10 = olFolderContacts
                return folder
            elif folder_name.lower() == "tasks":
                folder = self.namespace.GetDefaultFolder(13)  # 13 = olFolderTasks
                return folder
            else:
                folder = self._get_folder_by_name(folder_name)
                return folder
        except Exception as e:
            logger.error(f"Error getting folder: {str(e)}")
            raise
            
    def _get_folder_by_name(self, folder_name: str):
        """Find folder by name in folder hierarchy, supporting nested paths"""
        try:
            # Handle nested folder paths (e.g., "Parent Folder/Child Folder")
            if '/' in folder_name or '\\' in folder_name:
                # Use forward slash as path separator, but also support backslash
                path_parts = folder_name.replace('\\', '/').split('/')
                current_folder = None
                
                # Start with the top-level folders
                for folder in self.namespace.Folders:
                    if folder.Name == path_parts[0]:
                        current_folder = folder
                        break
                
                if not current_folder:
                    raise ValueError(f"Top-level folder '{path_parts[0]}' not found")
                
                # Navigate through the path
                for part in path_parts[1:]:
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