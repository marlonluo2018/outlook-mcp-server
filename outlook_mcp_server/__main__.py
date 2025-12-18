"""
Main entry point for the outlook_mcp_server package when executed as a module.
This allows the package to be run with 'python -m outlook_mcp_server'.
"""

import sys
from fastmcp import FastMCP

# Import from the package
try:
    # When running as module
    from .backend.outlook_session.session_manager import OutlookSessionManager
    from .tools.registration import register_all_tools
except ImportError:
    # When running as script or direct execution
    from outlook_mcp_server.backend.outlook_session.session_manager import OutlookSessionManager
    from outlook_mcp_server.tools.registration import register_all_tools

def test_outlook_connection() -> bool:
    """Test Outlook connection before starting the server.
    
    Returns:
        bool: True if connection successful, False otherwise
    """
    try:
        with OutlookSessionManager() as session:
            # Test basic folder access
            inbox = session.get_folder()
            if inbox and hasattr(inbox, 'Name'):
                return True
    except Exception as e:
        print(f"Outlook connection test failed: {str(e)}", file=sys.stderr)
        return False
    return False

def main():
    """Main function to start the Outlook MCP Server.
    
    This function serves as the entry point for module execution.
    It tests the Outlook connection before starting the MCP server.
    """
    try:
        # Test Outlook connection first
        if not test_outlook_connection():
            print("Error: Unable to connect to Outlook. Please ensure Outlook is installed and running.", file=sys.stderr)
            sys.exit(1)
        
        print("Outlook connection successful. Starting MCP server...", file=sys.stderr)
        
        # Initialize FastMCP server
        mcp = FastMCP("outlook-assistant")
        
        # Register all MCP tools
        register_all_tools(mcp)
        
        # Run the MCP server
        mcp.run()
        
    except KeyboardInterrupt:
        print("\nMCP server stopped by user.", file=sys.stderr)
        sys.exit(0)
    except Exception as e:
        print(f"Error starting server: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()