"""
Email operations for Outlook session management.

This module provides email-related operations such as moving emails between folders,
managing email policies, and retrieving email details.
"""

import logging
from typing import Any, Dict, List, Optional, Tuple

from ..outlook_session.session_manager import OutlookSessionManager
from ..shared import email_cache, email_cache_order
from ..validators import EmailNumberParam
from .exceptions import OperationFailedError, InvalidParameterError

# Set up logging
logger = logging.getLogger(__name__)


class EmailOperations:
    """Handles all email-related operations for Outlook."""
    
    def __init__(self, session_manager):
        """Initialize with a session manager instance."""
        self.session_manager = session_manager

    def get_email_by_number(self, email_number: int) -> Dict[str, Any]:
        """
        Get email details by cache number.
        
        Args:
            email_number: The number of the email in the cache (1-based)
        
        Returns:
            Email dictionary with full details
        """
        try:
            # Validate input parameters
            if email_number < 1:
                raise ValueError(f"Invalid email number: {email_number}")
        except Exception as e:
            logger.error(f"Validation error in get_email_by_number: {e}")
            raise ValueError(f"Invalid parameters: {e}")

        try:
            if email_number not in email_cache:
                raise ValueError(f"Email #{email_number} not found in cache")
            
            # Get the entry_id from the cache
            entry_id = list(email_cache.keys())[email_number - 1] if email_number <= len(email_cache) else None
            if not entry_id:
                raise ValueError(f"Email #{email_number} not found in cache")
            
            email_data = email_cache[entry_id]
            logger.info(f"Retrieved email #{email_number}: {email_data.get('subject', 'No Subject')}")
            return email_data
            
        except Exception as e:
            error_msg = f"Error getting email by number: {e}"
            logger.error(error_msg)
            raise ValueError(error_msg)

    def move_email_to_folder(self, email_number: int, target_folder_name: str) -> str:
        """
        Move an email to a different folder.
        
        Args:
            email_number: The number of the email in the cache (1-based)
            target_folder_name: Name of the target folder
        
        Returns:
            Success message or error message
        """
        try:
            # Validate input parameters
            if email_number < 1:
                return f"Error: Invalid email number: {email_number}"
            if not target_folder_name or not target_folder_name.strip():
                return f"Error: Invalid folder name: {target_folder_name}"
        except Exception as e:
            logger.error(f"Validation error in move_email_to_folder: {e}")
            return f"Error: Invalid parameters: {e}"

        try:
            # Get email from cache using the correct logic
            if not email_cache_order or email_number > len(email_cache_order):
                return f"Error: Email #{email_number} not found in cache"
            
            # Get the entry_id from the cache order
            entry_id = email_cache_order[email_number - 1]
            if not entry_id:
                return f"Error: Email #{email_number} has no entry ID"
            
            with OutlookSessionManager() as session:
                # Get the email item
                item = session.namespace.GetItemFromID(entry_id)
                if not item:
                    return f"Error: Could not find email with entry ID {entry_id}"
                
                # Get target folder
                target_folder = session.get_folder(target_folder_name)
                if not target_folder:
                    return f"Error: Target folder '{target_folder_name}' not found"
                
                # Move the email
                item.Move(target_folder)
                
                # Remove from cache since it's been moved
                if entry_id in email_cache:
                    del email_cache[entry_id]
                    if entry_id in email_cache_order:
                        email_cache_order.remove(entry_id)
                
                logger.info(f"Moved email #{email_number} to '{target_folder_name}'")
                return f"Email moved successfully to '{target_folder_name}'"
                
        except Exception as e:
            error_msg = f"Error moving email: {e}"
            logger.error(error_msg)
            return f"Error: {error_msg}"

    def delete_email_by_number(self, email_number: int) -> str:
        """
        Delete an email by moving it to the Deleted Items folder.
        
        Args:
            email_number: The number of the email in the cache (1-based)
        
        Returns:
            Success message or error message
        """
        try:
            # Validate input parameters
            if email_number < 1:
                return f"Error: Invalid email number: {email_number}"
        except Exception as e:
            logger.error(f"Validation error in delete_email_by_number: {e}")
            return f"Error: Invalid parameters: {e}"

        try:
            return self.move_email_to_folder(email_number, "Deleted Items")
            
        except Exception as e:
            error_msg = f"Error deleting email: {e}"
            logger.error(error_msg)
            return f"Error: {error_msg}"


def get_email_by_number(email_number: int) -> Dict[str, Any]:
    """Get email details by cache number."""
    with OutlookSessionManager() as session_manager:
        email_ops = EmailOperations(session_manager)
        return email_ops.get_email_by_number(email_number)


def move_email_to_folder(email_number: int, target_folder_name: str) -> str:
    """Move an email to a different folder."""
    with OutlookSessionManager() as session_manager:
        email_ops = EmailOperations(session_manager)
        return email_ops.move_email_to_folder(email_number, target_folder_name)


def delete_email_by_number(email_number: int) -> str:
    """Delete an email by moving it to the Deleted Items folder."""
    with OutlookSessionManager() as session_manager:
        email_ops = EmailOperations(session_manager)
        return email_ops.delete_email_by_number(email_number)



