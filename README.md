# Outlook MCP Server

An MCP (Model Context Protocol) server that provides tools for interacting with Microsoft Outlook via Python.

## Features

- Connect to Outlook and manage email operations
- List folders and recent emails
- Search emails with advanced filtering
- View detailed email content
- Reply to emails and manage recipients
- Handle attachments
- Batch processing support
- Robust error handling and logging
- Unit test coverage

## Requirements

- Windows OS (required for Outlook COM automation)
- Python 3.8 or higher
- Microsoft Outlook installed and configured
- Required Python packages (see requirements.txt)

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/outlook-mcp-server.git
cd outlook-mcp-server

# Install required packages
pip install -r requirements.txt

# Install in development mode
pip install -e .
```

## Configuration

The server uses the following configuration constants that can be adjusted in `outlook_operations.py`:

```python
MAX_DAYS = 30          # Maximum days to look back
MAX_EMAILS = 1000      # Maximum emails to process
MAX_LOAD_TIME = 58     # Maximum processing time in seconds
CONNECT_TIMEOUT = 30   # Connection timeout in seconds
MAX_RETRIES = 3        # Maximum connection retry attempts
```

## Usage

### Starting the Server

```bash
python outlook_mcp_server.py
```

### Available Tools

#### 1. List Folders
```python
list_folders()  # Lists all available Outlook folders
```

#### 2. List Recent Emails
```python
list_recent_emails(
    days=7,              # Number of days to look back
    folder_name="Inbox"  # Optional folder name
)
```

#### 3. Search Emails
```python
search_emails(
    search_term="important",  # Search term
    days=7,                   # Number of days to look back
    folder_name="Inbox",      # Optional folder name
    match_all=True           # Match all terms (AND) or any term (OR)
)
```

#### 4. View Email Cache
```python
view_email_cache(
    page=1,     # Page number (5 emails per page)
)
```

#### 5. Get Email Details
```python
get_email_by_number(
    email_number=1  # Email number from cache
)
```

#### 6. Reply to Email
```python
reply_to_email_by_number(
    email_number=1,
    reply_text="Thank you for your email",
    to_recipients=None,  # Optional list of recipients
    cc_recipients=None   # Optional list of CC recipients
)
```

## Development

### Running Tests

```bash
# Run all tests
pytest

# Run tests with coverage
pytest --cov=.

# Run specific test file
pytest tests/test_outlook_operations.py
```

### Project Structure

```
outlook-mcp-server/
├── outlook_mcp_server.py   # Main MCP server
├── outlook_operations.py   # Core Outlook operations
├── utils.py               # Utility functions
├── requirements.txt       # Dependencies
├── setup.py              # Package configuration
├── tests/                # Test directory
│   ├── __init__.py
│   └── test_outlook_operations.py
└── README.md
```

## Error Handling

The server includes comprehensive error handling:
- Connection retry with exponential backoff
- Timeout handling for long-running operations
- Detailed error logging
- Graceful degradation for non-critical failures

## Logging

Logs are available at multiple levels:
- INFO: General operation information
- WARNING: Non-critical issues
- ERROR: Critical issues that need attention
- DEBUG: Detailed debugging information

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/your-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin feature/your-feature`)
5. Create a new Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.
