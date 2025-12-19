"""
Outlook session management functionality.

This module provides the core session management capabilities for Outlook COM operations.
"""

import logging
import pythoncom
import win32com.client
from typing import Optional

from ..shared import configure_logging
from ..utils import retry_on_com_error
from .exceptions import ConnectionError
from .folder_operations import FolderOperations

# Initialize logging
configure_logging()
logger = logging.getLogger(__name__)


class OutlookSessionManager:
    """Context manager for Outlook COM session handling with improved resource management."""

    def __init__(self):
        self.outlook = None
        self.namespace = None
        self._connected = False
        self._com_initialized = False
        self._folder_operations = None

    def __enter__(self):
        """Initialize Outlook COM objects."""
        self._connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up Outlook COM objects."""
        self._disconnect()
        return False  # Don't suppress exceptions

    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def _connect(self):
        """Establish COM connection with proper threading and retry logic."""
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self._folder_operations = FolderOperations(self)
            self._connected = True
            logger.info("Successfully connected to Outlook")
        except Exception as e:
            logger.error(f"Connection error: {str(e)}")
            self._cleanup_partial_connection()
            raise ConnectionError(f"Failed to connect to Outlook: {str(e)}") from e

    def _cleanup_partial_connection(self):
        """Clean up partial connection attempts."""
        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.warning(f"Error cleaning up partial connection: {str(e)}")
            finally:
                self._com_initialized = False

    def _disconnect(self):
        """Clean up COM objects with proper resource release."""
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

    def is_connected(self) -> bool:
        """Check if the session is still connected."""
        try:
            # Simple operation to test connection
            if self._connected and self.outlook:
                from ..utils import OutlookFolderType
                self.outlook.GetNamespace("MAPI").GetDefaultFolder(OutlookFolderType.INBOX).Name
                return True
            return False
        except:
            self._connected = False
            return False

    def reconnect(self):
        """Re-establish the Outlook connection."""
        self._disconnect()
        self._connect()

    @property
    def outlook_app(self):
        """Get the Outlook application object."""
        return self.outlook

    @property
    def outlook_namespace(self):
        """Get the Outlook namespace object."""
        return self.namespace

    def get_folder(self, folder_name: Optional[str] = None):
        """Get a folder by name using folder operations."""
        if not self._folder_operations:
            raise ConnectionError("Folder operations not initialized. Ensure Outlook is connected.")
        return self._folder_operations.get_folder(folder_name)

    def get_folder_emails(self, folder_name: str = "Inbox", max_emails: int = 100, fast_mode: bool = True):
        """Get emails from a folder using folder operations."""
        if not self._folder_operations:
            raise ConnectionError("Folder operations not initialized. Ensure Outlook is connected.")
        return self._folder_operations.get_folder_emails(folder_name, max_emails, fast_mode)