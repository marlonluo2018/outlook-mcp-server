# Outlook MCP Server

An MCP (Model Context Protocol) server that provides tools for interacting with Microsoft Outlook via Python COM API.

## Features

- Connect to Outlook and manage email operations
- List folders and recent emails (with pagination)
- Search emails with advanced filtering (AND/OR logic)
- View detailed email content (including attachments)
- Reply to emails with custom recipients
- Compose new emails (HTML/plain text support)
- Email caching system for performance
- Batch processing with timeout handling
- Robust error handling and input validation

## Requirements

- Windows OS (required for Outlook COM automation)
- Python 3.8+
- Microsoft Outlook installed and configured
- Required packages: `fastmcp`, `pywin32`

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/outlook-mcp-server.git
cd outlook-mcp-server

# Install requirements
pip install -r requirements.txt

# Run the server
python outlook_mcp_server.py
```

## Configuration

Constants are defined in `backend/shared.py`:

```python
MAX_DAYS = 30          # Maximum days to look back
MAX_EMAILS = 1000      # Maximum emails to process  
MAX_LOAD_TIME = 58     # Maximum processing time (seconds)
```

## Available Tools

### 1. List Folders
```python
get_folder_list_tool()
# Returns all Outlook folders as list
```

### 2. List Recent Emails
```python
list_recent_emails_tool(
    days=7,              # Days to look back (1-30)
    folder_name="Inbox"  # Optional folder name
)
# Returns count + first page preview
```

### 3. Search Emails  
```python
search_emails_tool(
    search_term="important",  # Search term(s)
    days=7,                  # Days to look back
    folder_name="Inbox",     # Optional folder
    match_all=True           # AND/OR logic
)
# Returns count + first page preview
```

### 4. View Email Cache
```python 
view_email_cache_tool(
    page=1  # Page number (5 emails/page)
)
# Returns formatted email previews
```

### 5. Get Email Details
```python
get_email_by_number_tool(
    email_number=1  # Cache position (1-based)
)
# Returns full email content
```

### 6. Reply to Email
```python
reply_to_email_by_number_tool(
    email_number=1,
    reply_text="Thank you",
    to_recipients=None,  # Custom To: (None=original)
    cc_recipients=None   # Custom CC: (None=original) 
)
# Requires explicit user confirmation
```

### 7. Compose Email
```python
compose_email_tool(
    recipient_email="user@example.com",
    subject="Hello",
    body="Message content", 
    cc_email=None        # Optional CC
)
# Requires explicit user confirmation
```

## Project Structure

```
outlook-mcp-server/
├── outlook_mcp_server.py   # Main MCP server with tools
├── backend/
│   ├── email_retrieval.py  # Email listing/searching
│   ├── email_composition.py # Email sending
│   ├── outlook_session.py  # Outlook connection mgmt
│   ├── batch_operations.py # Batch processing
│   └── shared.py          # Shared constants
├── requirements.txt       # Dependencies
└── README.md
```

## Important Notes

1. Email sending operations REQUIRE explicit user confirmation
2. Search is limited to email subjects only
3. Default pagination shows 5 emails per page
4. Cache is cleared when listing new emails

## License

MIT License
