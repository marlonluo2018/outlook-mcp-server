# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Setup and Installation
```bash
# Install dependencies
pip install -r requirements.txt

# Install development dependencies (optional)
pip install fastmcp==2.11.0 pywin32==306
```

### Running the Application
```bash
# Start MCP server (for LLM integration)
python outlook_mcp_server.py

# Start interactive CLI (for human users)
python cli_interface.py
```

### Code Quality and Testing
```bash
# Format code with Black
black .

# Type checking with MyPy
mypy .

# Run tests (if test suite exists)
pytest tests/
```

## Architecture Overview

This is a Model Context Protocol (MCP) server that provides programmatic access to Microsoft Outlook via COM automation. The architecture separates concerns between MCP server interface, backend operations, and user interaction.

### Core Components

1. **MCP Server Layer** (`outlook_mcp_server.py`):
   - FastMCP-based server providing JSON-RPC interface
   - Tool registration and parameter validation
   - Standardized MCP response formatting

2. **Backend Operations** (`backend/`):
   - `outlook_session.py`: COM session management with connection pooling
   - `email_retrieval.py`: Email search, listing, and caching operations
   - `email_composition.py`: Email creation and reply functionality
   - `batch_operations.py`: Bulk email sending operations
   - `shared.py`: Global configuration and email cache management

3. **CLI Interface** (`cli_interface.py`):
   - Interactive command-line interface for human users
   - Menu-driven workflow with email cache management
   - Separate from MCP server (human-only interface)

### Key Design Patterns

**Email Cache Workflow**:
- All operations work with in-memory email cache
- Cache is populated via `list_recent_emails_tool` or `search_emails` function (CLI only)
- Email operations reference cache entries by position number
- Cache automatically refreshes on new list/search operations

**COM Session Management**:
- Context manager pattern for proper COM object cleanup
- Connection pooling with automatic reconnection
- Thread-safe initialization via `pythoncom.CoInitialize()`

**Security Model**:
- All email sending operations require explicit user confirmation
- Input validation and sanitization for all parameters
- No persistent storage of email content

### Configuration Constants

Key constants defined in `backend/shared.py`:
- `MAX_DAYS`: 30 (maximum days to look back for emails)
- `MAX_EMAILS`: 1000 (maximum emails to process per operation)
- `MAX_LOAD_TIME`: 58 seconds (timeout for email operations)
- `CONNECT_TIMEOUT`: 30 seconds (Outlook connection timeout)

### Tool Call Patterns

MCP tools follow a specific sequence:
1. **Cache Population**: Call `list_recent_emails_tool` or `search_emails` function (CLI only)
2. **Cache Browsing**: Use `view_email_cache_tool` to identify target email
3. **Content Retrieval**: Use `get_email_by_number_tool` for full email details
4. **Email Operations**: Use composition tools (requires user confirmation)

### Error Handling Strategy

- Comprehensive try/catch blocks around all COM operations
- Retry logic with exponential backoff for connection issues
- User-friendly error messages with MCP error response format
- Graceful degradation when Outlook is unavailable

### Development Notes

- **Windows-only**: Requires pywin32 and Outlook COM access
- **COM Threading**: All COM operations must be properly initialized
- **Memory Management**: COM objects must be explicitly cleaned up
- **Performance**: Email cache limits prevent memory overload
- **Security**: Write operations always require user confirmation