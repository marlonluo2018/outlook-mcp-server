"""MCP Tools Package for Outlook MCP Server.

This package contains all MCP tool definitions organized by functionality:
- folder_tools: Folder management operations
- search_tools: Email search functionality  
- viewing_tools: Email viewing and cache management
- email_operations: Email composition and manipulation
- batch_operations: Batch email operations
"""

from .folder_tools import (
    move_folder_tool,
    get_folder_list_tool,
    create_folder_tool,
    remove_folder_tool,
)

from .search_tools import (
    list_recent_emails_tool,
    search_email_by_subject_tool,
    search_email_by_sender_name_tool,
    search_email_by_recipient_name_tool,
    search_email_by_body_tool,
)

from .viewing_tools import (
    view_email_cache_tool,
    get_email_by_number_tool,
    load_emails_by_folder_tool,
)

from .email_operations import (
    reply_to_email_by_number_tool,
    compose_email_tool,
    move_email_tool,
    delete_email_by_number_tool,
)

from .batch_operations import batch_forward_email_tool

__all__ = [
    # Folder tools
    'move_folder_tool',
    'get_folder_list_tool', 
    'create_folder_tool',
    'remove_folder_tool',
    
    # Search tools
    'list_recent_emails_tool',
    'search_email_by_subject_tool',
    'search_email_by_sender_name_tool',
    'search_email_by_recipient_name_tool',
    'search_email_by_body_tool',
    
    # Viewing tools
    'view_email_cache_tool',
    'get_email_by_number_tool',
    'load_emails_by_folder_tool',
    
    # Email operations
    'reply_to_email_by_number_tool',
    'compose_email_tool',
    'move_email_tool',
    'delete_email_by_number_tool',
    
    # Batch operations
    'batch_forward_email_tool',
]