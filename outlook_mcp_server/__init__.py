"""Outlook MCP Server - Package initialization module.

This module handles package-level initialization and imports.
"""

# Import shared module to ensure cache is loaded on startup
from .backend import shared

# Backend imports - organized by functionality
from .backend.outlook_session import OutlookSessionManager

# Search functionality
from .backend import email_search

# Import specific functions from the email_search module
list_folders = email_search.list_folders
search_email_by_subject = email_search.search_email_by_subject
search_email_by_from = email_search.search_email_by_from
search_email_by_to = email_search.search_email_by_to
search_email_by_body = email_search.search_email_by_body
list_recent_emails = email_search.list_recent_emails
get_emails_from_folder = email_search.get_emails_from_folder

# Email data extraction and formatting
from .backend.email_data_extractor import get_email_by_number_unified, format_email_with_media

# Cache management
from .backend.shared import clear_email_cache, add_email_to_cache, save_email_cache, refresh_email_cache_with_new_data
from .backend.email_search.search_common import unified_cache_load_workflow, extract_email_info_minimal, clear_com_attribute_cache

# Performance optimizations
from .backend.email_search.parallel_extractor import extract_emails_optimized

# Email composition and operations
from .backend.email_composition import reply_to_email_by_number, compose_email

# Batch operations
from .backend.batch_operations import batch_forward_emails

# Tool registration
from .tools.registration import register_all_tools

# Version info
__version__ = "1.0.0"
__author__ = "Outlook MCP Server Team"

# Package-level exports
__all__ = [
    # Core functionality
    'OutlookSessionManager',
    
    # Search functions
    'list_folders',
    'search_email_by_subject',
    'search_email_by_from',
    'search_email_by_to',
    'search_email_by_body',
    'list_recent_emails',
    'get_emails_from_folder',
    
    # Email operations
    'get_email_by_number_unified',
    'format_email_with_media',
    'reply_to_email_by_number',
    'compose_email',
    
    # Cache management
    'clear_email_cache',
    'add_email_to_cache',
    'save_email_cache',
    'refresh_email_cache_with_new_data',
    'unified_cache_load_workflow',
    'extract_email_info_minimal',
    'clear_com_attribute_cache',
    'extract_emails_optimized',
    
    # Batch operations
    'batch_forward_emails',
    
    # Tool registration
    'register_all_tools',
]