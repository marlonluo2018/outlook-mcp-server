"""
Exception classes for Outlook session operations.

This module defines custom exceptions used throughout the outlook session package.
"""


class OutlookSessionError(Exception):
    """Base exception for Outlook session related errors."""
    pass


class ConnectionError(OutlookSessionError):
    """Raised when Outlook connection fails."""
    pass


class FolderNotFoundError(OutlookSessionError):
    """Raised when a specified folder cannot be found."""
    pass


class EmailNotFoundError(OutlookSessionError):
    """Raised when a specified email cannot be found."""
    pass


class InvalidParameterError(OutlookSessionError):
    """Raised when invalid parameters are provided."""
    pass


class OperationFailedError(OutlookSessionError):
    """Raised when an Outlook operation fails."""
    pass