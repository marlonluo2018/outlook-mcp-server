from backend.shared import configure_logging

import pythoncom
import win32com.client
from typing import Optional
import logging

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
                
    def get_folder(self, folder_name: Optional[str] = None):
        """Get specified folder or default inbox"""
        try:
            if not folder_name:
                return self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            return self._get_folder_by_name(folder_name)
        except Exception as e:
            logger.error(f"Error getting folder: {str(e)}")
            raise
            
    def _get_folder_by_name(self, folder_name: str):
        """Find folder by name in folder hierarchy"""
        try:
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