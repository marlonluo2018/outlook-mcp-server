"""Tool registration module for Outlook MCP Server.

This module handles the registration of all MCP tools with the FastMCP server.
"""

from fastmcp import FastMCP
from . import (
    # Folder tools
    move_folder_tool,
    get_folder_list_tool,
    create_folder_tool,
    remove_folder_tool,
    
    # Search tools
    list_recent_emails_tool,
    search_email_by_subject_tool,
    search_email_by_sender_name_tool,
    search_email_by_recipient_name_tool,
    search_email_by_body_tool,
    
    # Viewing tools
    view_email_cache_tool,
    get_email_by_number_tool,
    load_emails_by_folder_tool,
    clear_email_cache_tool,
    
    # Email operations
    reply_to_email_by_number_tool,
    compose_email_tool,
    move_email_tool,
    delete_email_by_number_tool,
    
    # Batch operations
    batch_forward_email_tool,
)


def register_all_tools(mcp_server: FastMCP) -> None:
    """Register all MCP tools with the FastMCP server.
    
    Args:
        mcp_server: The FastMCP server instance to register tools with
    """
    # Folder management tools
    mcp_server.tool(move_folder_tool)
    mcp_server.tool(get_folder_list_tool)
    mcp_server.tool(create_folder_tool)
    mcp_server.tool(remove_folder_tool)
    
    # Search tools
    mcp_server.tool(list_recent_emails_tool)
    mcp_server.tool(search_email_by_subject_tool)
    mcp_server.tool(search_email_by_sender_name_tool)
    mcp_server.tool(search_email_by_recipient_name_tool)
    mcp_server.tool(search_email_by_body_tool)
    
    # Viewing tools
    mcp_server.tool(view_email_cache_tool)
    mcp_server.tool(get_email_by_number_tool)
    mcp_server.tool(load_emails_by_folder_tool)
    mcp_server.tool(clear_email_cache_tool)
    
    # Email operations
    mcp_server.tool(reply_to_email_by_number_tool)
    mcp_server.tool(compose_email_tool)
    mcp_server.tool(move_email_tool)
    mcp_server.tool(delete_email_by_number_tool)
    
    # Batch operations
    mcp_server.tool(batch_forward_email_tool)