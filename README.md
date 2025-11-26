# Outlook MCP Server

A powerful MCP (Model Context Protocol) server that provides seamless integration with Microsoft Outlook through COM interface. This server enables AI assistants to interact with your Outlook emails, including searching, reading, composing, and sending emails.

## Prerequisites

Before installing and running the Outlook MCP Server, ensure you have the following:

### Required Software

1. **Python 3.8 or higher**
   - Download from [python.org](https://www.python.org/downloads/)
   - Make sure to add Python to your PATH during installation

2. **Microsoft Outlook**
   - Outlook 2016 or later (including Outlook 365)
   - Must be installed on the same machine where you'll run the server

3. **Windows Operating System**
   - Windows 10 or later
   - Required for COM interface access to Outlook

### Optional Tools

1. **UVX (Recommended)**
   - A modern Python package runner that creates isolated environments
   - Install with: `pip install uvx`
   - Benefits: Automatic dependency management, clean environments, no system pollution

2. **Git**
   - For cloning the repository
   - Download from [git-scm.com](https://git-scm.com/downloads)

3. **Build Tools (Optional)**
    - Required for most installation methods (except UVX with local source)
    - Install with: `pip install build twine`
    - Used for: Creating distribution files from source code
    - Note: Pre-built distribution files are already included in the `dist/` folder
    
    Installation options:
    1. UVX with local source - No build required
    2. Editable Installation (`pip install -e .`) - Build required
    3. Standard Installation (`pip install .`) - Build required
    4. Direct Python file execution - No build required (but dependencies needed)

### Building the Package

To install the package using methods 2 or 3, you need to build it first:

```bash
# Navigate to the project directory
cd outlook-mcp-server

# Build the package
python -m build
```

This will create distribution files in the `dist/` directory:
- `outlook_mcp_server-X.X.X-py3-none-any.whl` (wheel file)
- `outlook_mcp_server-X.X.X.tar.gz` (source distribution)

Note: Pre-built distribution files are already included in the `dist/` folder, so you can use those directly with `pip install dist/outlook_mcp_server-0.1.0-py3-none-any.whl`.

## ⚡ Quick Start: Installation Methods

Choose one of the following installation options:

### Option 1: Using UVX with Local Package (Recommended)

No build required. Run directly from source:

```bash
# Clone the repository
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server

# Run directly with UVX (no installation needed)
uvx --with "pywin32>=226" --with-editable "c:\\Project\\outlook-mcp-server" outlook-mcp-server
```

### Option 2: Editable Installation

Build required. Install in editable mode for development:

```bash
# Clone the repository
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server

# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install in editable mode
pip install -e .

# Run the server
python -m outlook_mcp_server
```

### Option 3: Standard Installation

Build required. Install the package:

```bash
# Clone the repository
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server

# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install the package
pip install .

# Run the server
python -m outlook_mcp_server
```

### Option 4: Direct Python File Execution

No build required, but dependencies needed:

```bash
# Clone the repository
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server

# Install dependencies
pip install -r requirements.txt

# Run the server
python outlook_mcp_server/__main__.py
```

## Running the Server: Usage Methods

### Prerequisites

Before running the server, you must either:
1. Install the package using one of the installation methods above, or
2. Use UVX (Option 1) which doesn't require installation

### Option 1: As an MCP Server (for LLM integration)

After installation:

```bash
# Run the MCP server
python -m outlook_mcp_server
```

Or with UVX (no installation needed):

```bash
uvx --with "pywin32>=226" --with-editable "c:\\Project\\outlook-mcp-server" outlook-mcp-server
```

### Option 2: As a CLI Interface (for human interaction)

After installation:

```bash
# Run the CLI interface
python -m outlook_mcp_server.cli_interface
```

### Option 3: Using the Package Command (after installation)

```bash
# Run using the package command
outlook-mcp-server
```

## MCP Configuration: Integration Methods

### Option 1: UVX Configuration with Local Package (Recommended)

For using the local package with UVX, use this configuration in your MCP client (e.g., Claude Desktop settings.json):

```json
{
  "mcpServers": {
    "outlook-mcp-server": {
      "command": "uvx",
      "args": [
        "--with", "pywin32>=226",
        "--with-editable", "c:\\Project\\outlook-mcp-server",
        "outlook-mcp-server"
      ]
    }
  }
}
```

### Option 2: Direct Python Command Configuration

For MCP clients that support direct Python commands or if you prefer to use your existing Python installation:

```json
{
  "mcpServers": {
    "outlook-mcp-server": {
      "command": "python",
      "args": ["-m", "outlook_mcp_server"]
    }
  }
}
```

Note: With this approach, you'll need to install the dependencies first:
```bash
pip install -r requirements.txt
```

This configuration file is provided as `mcp-config-python.json` in the project root for convenience.

### Option 3: UVX Configuration (for published package)

Once the package is published to PyPI, you can use UVX:

```json
{
  "mcpServers": {
    "outlook-mcp-server": {
      "command": "uvx",
      "args": [
        "--with", "pywin32>=226",
        "outlook-mcp-server"
      ]
    }
  }
}
```

### Option 4: Direct Python File Execution

For MCP clients that support direct Python file execution or if you prefer to run the Python file directly:

```json
{
  "mcpServers": {
    "outlook-mcp-server": {
      "command": "python",
      "args": ["outlook_mcp_server/__main__.py"]
    }
  }
}
```

Note: With this approach, you'll need to install the dependencies first:
```bash
pip install -r requirements.txt
```

This configuration file is provided as `mcp-config-python.json` in the project root for convenience.

### Human Interface: CLI

The CLI interface is designed exclusively for human users, not LLMs:

- **Purpose**: Provides a human-friendly interactive interface
- **Audience**: Only for direct human operation
- **Security**: All operations require interactive confirmation
- **Workflow**:
  1. Load emails into cache (List/Search)
  2. Operate on cached emails (View/Reply/Compose)
  3. Cache auto-refreshes on new operations

```bash
# Start interactive session (human only)
python -m outlook_mcp_server.cli_interface
```

Note: For LLM integration, use the MCP server interface instead.

### Development Installation

For development, we recommend installing in editable mode:

```bash
# Install in editable mode with development dependencies
pip install -e ".[dev]"
```

**Benefits of editable installation:**
- Changes to source code immediately affect the installed package
- No need to reinstall after making changes
- Ideal for testing and development

**Development Tools:**
- **Virtual Environment**: Always use a virtual environment for development
  ```bash
  python -m venv venv
  source venv/bin/activate  # On Windows: venv\Scripts\activate
  ```
- **Code Linting**: Install development dependencies for code quality
  ```bash
  pip install -e ".[dev]"
  ```
- **Testing**: Run tests with the development setup
  ```bash
  python -m pytest
  ```

## 🔧 Building and Distribution

### Building the Package

To build the package for distribution:

```bash
# Install build tools
pip install build twine

# Build the package
python -m build
```

This will create the distribution files in the `dist/` directory:
- `outlook_mcp_server-X.X.X-py3-none-any.whl` (wheel file)
- `outlook_mcp_server-X.X.X.tar.gz` (source distribution)

### Installing from Distribution Files

```bash
# Install from wheel file
pip install dist/outlook_mcp_server-X.X.X-py3-none-any.whl

# Or install from source distribution
pip install dist/outlook_mcp_server-X.X.X.tar.gz
```

### Publishing to PyPI (for maintainers)

```bash
# Upload to PyPI (requires credentials)
python -m twine upload dist/*
```

## 🔧 Building with UVX

### What is UVX?

UVX is a Python application runner that creates isolated environments for Python applications. It's ideal for running Python tools with their dependencies without polluting your system Python installation.

### Building with UVX

The Outlook MCP Server is fully compatible with UVX. Here's how to build and run it using UVX:

#### Method 1: Direct Execution

```bash
# Run the server directly with UVX
uvx --with "pywin32>=226" --with-editable "c:\\Project\\outlook-mcp-server" outlook-mcp-server
```

#### Method 2: MCP Configuration

For MCP clients like Claude Desktop or Trae IDE, you have two configuration options:

**Option A: UVX Configuration (Recommended)**
Use the configuration in your `mcp-config-uvx.json`:

```json
{
  "mcpServers": {
    "outlook-mcp-server": {
      "command": "uvx",
      "args": [
        "--with", "pywin32>=226",
        "--with-editable", "c:\\Project\\outlook-mcp-server",
        "outlook-mcp-server"
      ]
    }
  }
}
```

**Option B: Direct Python Command**
Use the configuration in your `mcp-config-python.json`:

```json
{
  "mcpServers": {
    "outlook-mcp-server": {
      "command": "python",
      "args": ["-m", "outlook_mcp_server"]
    }
  }
}
```

Note: With this approach, you'll need to install the dependencies first:
```bash
pip install -r requirements.txt
```

### UVX Benefits

- **Isolation**: Each application runs in its own environment
- **Dependency Management**: Automatic handling of dependencies
- **No System Pollution**: Keeps your system Python clean
- **Reproducibility**: Consistent environments across different machines
- **Editable Mode**: Supports local development with --with-editable flag

### Project Structure for UVX

The project is structured to work seamlessly with UVX:

```
outlook-mcp-server/
├── pyproject.toml          # Project metadata and dependencies
├── outlook_mcp_server/      # Package directory
│   ├── __init__.py         # Main package initialization
│   ├── __main__.py         # Enables module execution (python -m outlook_mcp_server)
│   └── backend/            # Backend modules
└── requirements.txt        # Dependencies
```

### UVX Configuration Details

The `pyproject.toml` file contains the project metadata and dependencies:

```toml
[build-system]
requires = ["setuptools>=61.0"]
build-backend = "setuptools.build_meta"

[project]
name = "outlook-mcp-server"
version = "1.3.0"
dependencies = [
    "fastmcp==2.13.1",
    "pywin32==311"
]

[project.scripts]
outlook-mcp-server = "outlook_mcp_server:main"
```

This configuration allows UVX to:
1. Identify the project dependencies
2. Create an isolated environment
3. Execute the server with all required packages
4. Support editable mode for local development

### Troubleshooting UVX

**Module Not Found Error**
```
ModuleNotFoundError: No module named 'fastmcp'
```
- Ensure `pyproject.toml` is in the project root
- Check that dependencies are correctly listed
- Verify the project path is correct in the UVX command

**Path Issues**
```
FileNotFoundError: [Errno 2] No such file or directory
```
- Use absolute paths in your configuration
- Ensure paths use proper escaping for backslashes on Windows
- Verify the project directory structure is correct

**Permission Issues**
```
Permission denied: .../outlook_mcp_server.py
```
- Ensure the Python script has execute permissions
- Run with appropriate user privileges
- Check that antivirus software isn't blocking execution

### Configuration Constants

Located in `backend/shared.py`:

| Constant | Default | Description |
|----------|---------|-------------|
| `MAX_DAYS` | 30 | Maximum days to look back for emails |
| `MAX_EMAILS` | 1000 | Maximum emails to process in one operation |
| `MAX_LOAD_TIME` | 58 | Maximum processing time per operation (seconds) |
| `CONNECT_TIMEOUT` | 30 | Connection timeout for Outlook |
| `MAX_RETRIES` | 3 | Maximum retry attempts for failed operations |

## 🎯 Usage

### MCP Server Mode (For LLM Tool Calls)

**🔄 How to Use with AI Assistants**:

**Step 1: Load Your Emails**  
Ask your AI assistant to:  
- "Show me recent emails from my Inbox"  
- "Search for email subjects about project updates"  
- "Find emails from the last 2 weeks"  
- "Search for emails from John Doe"  
- "Search for emails sent to Team Name"  
- "Search for emails with specific content in the body"

This loads your emails into a temporary cache.

**Step 2: Browse Available Emails**  
Ask your AI assistant to:  
- "Show me the emails you found"  
- "Display page 2 of the email list"  
- "What emails are in my cache?"

This shows you a list of available emails with numbers.

**Step 3: View Specific Email**  
Tell your AI assistant:  
- "Show me email number 3"  
- "Get the full content of email 5"  
- "Display email about the meeting"

This retrieves the complete email for you to review.

**Step 4: Take Action**  
Ask your AI assistant to:  
- "Reply to email 3 with: Thanks for the update!"  
- "Compose a new email to my team"  
- "Forward email 7 to my manager"

**⚠️ Important Notes**:

- **Cache System**: Each time you search or list emails, the previous cache is cleared
- **Email Numbers**: Always refer to email numbers shown in the current cache
- **Your Approval**: The AI will always ask for your confirmation before sending any emails
- **Temporary Storage**: Email cache is only in memory - nothing is saved permanently

### CLI Interface (For Direct Human Use)

**🖥️ How to Use the Interactive Menu System**:

**Getting Started**:
```bash
# Start the interactive email session
python -m outlook_mcp_server.cli_interface
```

**🔄 Step-by-Step Workflow**:

**Step 1: Load Emails into Cache**  
Start by choosing one of these options:  
- **Menu Option 2**: List recent emails (specify days and folder)  
- **Menu Option 3**: Search email subjects (enter search terms and filters)  
- **Menu Option 4**: Search emails by sender name (enter sender name)  
- **Menu Option 5**: Search emails by recipient name (enter recipient name)  
- **Menu Option 6**: Search emails by body content (enter search terms)  

This clears any previous cache and loads your selected emails.

**Step 2: Browse Available Emails**  
Use **Menu Option 7** to:  
- View all emails currently in your cache  
- See a numbered list with subjects and senders  
- Navigate through multiple pages if needed  
- Note the email number you want to work with

**Step 3: View Email Details**  
Use **Menu Option 8** to:  
- Enter the email number you want to read  
- See the complete email content  
- Check attachments and recipient details  

**Step 4: Take Action on Email**  
Choose your action:  
- **Menu Option 9**: Reply to the email (enter email number)  
- **Menu Option 10**: Compose a new email  
- **Menu Option 11**: Send batch emails (using cached email as template)  

**📋 Common Usage Patterns**:

**To reply to an email**: 2 → 7 → 8 → 9  
(List emails → View cache → Read email → Reply)

**To search and respond**: 3 → 7 → 8 → 9  
(Search email subjects → View cache → Read email → Reply)

**To search by sender name and respond**: 4 → 7 → 8 → 9  
(Search by sender name → View cache → Read email → Reply)

**To search by recipient name and respond**: 5 → 7 → 8 → 9  
(Search by recipient name → View cache → Read email → Reply)

**To search by body content and respond**: 6 → 7 → 8 → 9  
(Search by body content → View cache → Read email → Reply)

**To send batch emails**: 2 → 7 → 11  
(List emails → View cache → Send batch)

**⚠️ Important Notes**:

- **Cache Management**: Each time you use Option 2, 3, 4, 5, or 6, the cache is cleared and reloaded
- **Email Numbers**: Always use the numbers shown in the current cache (Option 7)
- **Your Confirmation**: The system asks for confirmation before sending any emails
- **Session-Based**: Cache persists until you exit or load new emails
- **Menu Navigation**: Use Option 0 to exit safely

**Best for**: Users who prefer direct, menu-driven control over email operations or need batch email capabilities.

## 🔍 Email Search System

### Search Types Overview

The Outlook MCP Server provides five distinct email search functions, each optimized for different use cases:

| Search Type | Function | Searches | Use Case |
|-------------|----------|----------|----------|
| **Subject Search** | `search_email_by_subject_tool` | Subject line, sender names, body content | Find emails by topic or keywords |
| **From Search** | `search_email_by_sender_name_tool` | Sender display names | Find emails from specific people |
| **To Search** | `search_email_by_recipient_name_tool` | Recipient display names (TO/CC) | Find emails to specific recipients |
| **Body Search** | `search_email_by_body_tool` | Full email body content | Find emails by content details |
| **General Search** | `search_emails_tool` *(CLI only)* | Subject, sender, body | Broad keyword searches *(CLI users)* |

### Search Logic System

#### Core Match Logic Parameter

All search functions use the `match_all` parameter to control matching behavior:

- **`match_all=true` (Default)**: **AND Logic** - All search terms must match
- **`match_all=false`**: **OR Logic** - Any search term can match

#### Search Term Processing

1. **Term Splitting**: Queries split by whitespace into individual terms
2. **Case Insensitive**: All matching performed in lowercase
3. **Multi-term Support**: Multiple search terms processed efficiently

### Field-Specific Search Behavior

#### 1. Subject Search (`search_email_by_subject_tool`)
- **Search Method**: Uses Outlook's built-in DASL syntax for optimal performance
- **Search Fields**: 
  - Subject line (`urn:schemas:httpmail:subject`)
  - Sender name (`urn:schemas:httpmail:fromname`) 
  - Body content (`urn:schemas:httpmail:textdescription`)
- **Logic Implementation**:
  - AND: Groups filters by term, combines with AND
  - OR: Combines all field filters with OR

#### 2. From Search (`search_email_by_sender_name_tool`)
- **Search Method**: Post-retrieval filtering
- **Search Fields**: Sender display names only
- **Important**: Searches display names (e.g., "John Smith"), NOT email addresses (e.g., "john.smith@company.com")
- **Matching**: Simple case-insensitive substring search

#### 3. To Search (`search_email_by_recipient_name_tool`)
- **Search Method**: Post-retrieval filtering
- **Search Fields**: Recipient display names in TO and CC fields
- **Important**: Searches display names only, NOT email addresses
- **Optimization**: Stops checking additional recipients once match found

#### 4. Body Search (`search_email_by_body_tool`)
- **Search Method**: Full email retrieval with advanced text analysis
- **Advanced Features**:
  - **Exact Phrase Search**: Supports quoted terms (`"exact phrase"`)
  - **Proximity Checking**: Ensures terms appear within 50 characters of each other
  - **Full Content Access**: Retrieves complete email body for accurate searching
  - **AND/OR Logic**: Same `match_all` parameter control

### Advanced Search Features

#### 1. Word-Based Search (Default)
```json
{
  "tool": "search_email_by_body_tool",
  "parameters": {
    "search_term": "project deadline meeting",
    "match_all": true
  }
}
// Searches for emails containing "project", "deadline", AND "meeting" close to each other
```

#### 2. Exact Phrase Search (With Quotes)
```json
{
  "tool": "search_email_by_body_tool",
  "parameters": {
    "search_term": "\"project deadline meeting\"",
    "match_all": true
  }
}
// Searches for emails containing the exact phrase "project deadline meeting"
```

#### 3. Broad Search (Any Term)
```json
{
  "tool": "search_email_by_subject_tool",
  "parameters": {
    "search_term": "project deadline",
    "match_all": false
  }
}
// Returns emails containing either "project" OR "deadline"
```

#### 4. Folder Hierarchy Support
All search functions support nested folder paths:

```json
{
  "tool": "search_email_by_subject_tool",
  "parameters": {
    "search_term": "meeting notes",
    "folder_name": "Projects/Subfolder/Project A"
  }
}
// Supports unlimited folder depth with "/" or "\\" separators
```

### Proximity Search Technology

**Body Search Proximity Checking**:
- **Distance**: Terms must appear within 50 characters of each other
- **Purpose**: Prevents false positives where terms appear in unrelated contexts
- **Example**: Searching "red hat partner day" won't return emails about "Redhat Incentive" programs
- **Availability**: Only in body search with `match_all=true`

### Search Result Comparison

| Search Type | Example | Logic | Results | Use Case |
|-------------|---------|-------|---------|----------|
| Word-based (AND) | "red hat partner day" | All terms close | 13 emails | High precision |
| Word-based (OR) | "red hat partner day" | Any term | 278 emails | Broad search |
| Exact phrase | "\"red hat partner day\"" | Exact match | 10 emails | Specific terms |
| From name | "John Doe" | Display name | 25 emails | Person-specific |
| To name | "Team Name" | Recipient name | 15 emails | Team/group search |

### Performance Considerations

- **Subject Search**: Fastest (uses Outlook's built-in search)
- **From/To Search**: Medium speed (retrieves all emails, then filters)
- **Body Search**: Slowest (retrieves full email content for accuracy)
- **Cache Management**: Each search clears previous cache automatically
- **Timeouts**: Protected by configurable time limits (default: 58 seconds)

### CLI vs MCP Consistency

**Unified Behavior**: Both CLI and MCP interfaces use the same search logic:
- **CLI Menu Options**: 3-6 correspond to the same search functions
- **MCP Tools**: All accept the same `match_all` parameter
- **Backend Functions**: Single implementation ensures consistency

## 📚 Available Tools

#### 1. List Folders

```json
{
  "tool": "get_folder_list_tool",
  "parameters": {}
}
// Returns: {"content": [{"type": "text", "text": "['Inbox', 'Sent Items', 'Drafts', ...]"}]}
```

#### 2. List Recent Emails

```json
{
  "tool": "list_recent_emails_tool",
  "parameters": {
    "days": 7,
    "folder_name": "Inbox"
  }
}
// Returns count and first page preview
```

#### 3. Search Email by Subject

```json
{
  "tool": "search_email_by_subject_tool",
  "parameters": {
    "search_term": "meeting notes",
    "days": 14,
    "folder_name": "Inbox",
    "match_all": true
  }
}
```

#### 4. Search Email by Sender Name

```json
{
  "tool": "search_email_by_sender_name_tool",
  "parameters": {
    "search_term": "John Doe",
    "days": 14,
    "folder_name": "Inbox",
    "match_all": true
  }
}
```

#### 5. Search Email by Recipient Name

```json
{
  "tool": "search_email_by_recipient_name_tool",
  "parameters": {
    "search_term": "Team Name",
    "days": 14,
    "folder_name": "Inbox",
    "match_all": true
  }
}
```

#### 6. Search Email by Body Content

```json
{
  "tool": "search_email_by_body_tool",
  "parameters": {
    "search_term": "project deadline",
    "days": 14,
    "folder_name": "Inbox",
    "match_all": true
  }
}
```

#### 7. View Email Cache

```json
{
  "tool": "view_email_cache_tool",
  "parameters": {
    "page": 1
  }
}
```

#### 8. Get Email Details

```json
{
  "tool": "get_email_by_number_tool",
  "parameters": {
    "email_number": 3
  }
}
// Returns full email with body and attachments
```

#### 9. Reply to Email

```json
{
  "tool": "reply_to_email_by_number_tool",
  "parameters": {
    "email_number": 5,
    "reply_text": "Thank you for your message...",
    "to_recipients": ["custom@example.com"],
    "cc_recipients": ["boss@example.com"]
  }
}
// ⚠️ Requires explicit user confirmation
```

#### 10. Compose New Email

```json
{
  "tool": "compose_email_tool",
  "parameters": {
    "recipient_email": "client@example.com",
    "subject": "Project Update",
    "body": "Dear team,\n\nHere's the latest update...",
    "cc_email": "manager@example.com"
  }
}
// ⚠️ Requires explicit user confirmation
```

#### 11. Batch Email Operations (Interactive Only)

**Workflow**:

1. First load template emails into cache via List/Search
2. Select cached email as template in interactive mode
3. Provide CSV of recipients and optional custom text
4. Confirm before sending

Note: Batch operations require working with the email cache and are only available through interactive CLI.

## 📊 Data Flow Architecture

### Email Processing Pipeline

```mermaid
Outlook COM → Session Manager → Email Parser → Cache → API Response
```

### Cache System

- **Key**: Email EntryID (unique Outlook identifier)
- **Value**: Structured email data (subject, sender, body, attachments)
- **Lifetime**: Cache cleared on new email listing/search
- **Format**: JSON-serializable Python dictionaries

### Error Handling Flow

```mermaid
Operation → Try/Catch → Retry Logic → User-friendly Error → Logging
```

## 🔒 Security Considerations

### Email Sending Protection

- **Explicit Confirmation**: All email sending operations require user approval
- **Rate Limiting**: Batch operations limited to 500 recipients per batch
- **Input Validation**: All inputs sanitized and validated before processing

### Data Privacy

- **Local Processing**: All operations performed locally on user's machine
- **No External Calls**: No data transmitted outside Outlook COM interface
- **Cache Isolation**: Email cache stored in memory only, no persistent storage

## 📈 Performance Optimization

### Caching Strategy

- **Smart Loading**: Only load requested email ranges
- **Batch Processing**: Process emails in configurable batches
- **Timeout Protection**: Automatic termination of long-running operations

### Memory Management

- **Streaming**: Large email bodies truncated in cache
- **Garbage Collection**: Automatic cleanup of COM objects
- **Resource Limits**: Configurable maximum email processing limits

## 🔧 Troubleshooting

### Common Issues

**Outlook Not Found**
```
Error: Outlook application not found or not accessible
```
- Ensure Outlook is installed and running
- Check that Outlook COM interface is enabled
- Run Outlook as administrator if needed

**COM Permission Errors**
```
Error: Access denied or COM server not available
```
- Check Windows COM security settings
- Ensure your user account has Outlook access
- Try running the script as administrator

**Email Loading Issues**
```
Error: Timeout loading emails or Operation took too long
```
- Reduce the number of days or search scope
- Check if Outlook has large folders that need indexing
- Try searching specific folders instead of entire mailbox

**Cache Not Working**
```
Error: No emails in cache or Invalid cache item
```
- Always run list/search operations first to populate cache
- Cache clears when new list/search operations are performed
- Use view_email_cache to verify cache contents

**MCP Server Communication Issues**
```
Error: Invalid JSON-RPC response or communication timeout
```
- Ensure you're using the latest version of the package
- The MCP server communicates via JSON-RPC protocol over stdio
- Any print statements to stdout will interfere with MCP communication
- Check that your MCP client is properly configured for stdio transport

## 📋 Changelog

### v1.3.0 (Current)
- **UVX Support**: Added full support for UVX application runner
- **Enhanced Installation**: Simplified installation process with UVX
- **Module Execution**: Added __main__.py for module execution support
- **Updated Documentation**: Comprehensive UVX configuration guide
- **Improved Configuration**: Updated MCP configuration examples with UVX
- **Fixed MCP Communication**: Removed stdout print statements that interfered with JSON-RPC protocol
- **Better Error Handling**: Redirected error messages to stderr to avoid protocol conflicts
- **Editable Mode Support**: Added support for --with-editable flag for local development

### v1.2.0
- **Enhanced Search Logic**: Added intelligent word proximity checking for more accurate search results
- **Improved Search Precision**: Word-based searches now check if search terms appear close to each other
- **Better Search Documentation**: Comprehensive guide on different search types and parameters
- **Default match_all=true**: All searches now default to requiring all terms to match

### v1.1.0
- **Enhanced Email Details**: Improved email data structure with recipient information
- **Better Body Formatting**: Enhanced email body formatting for replies and batch operations
- **MCP Response Refactoring**: Simplified response structure by removing unnecessary wrappers
- **Improved Error Handling**: Better input validation and error messages
- **Project Configuration**: Added pyproject.toml for proper project management
- **Code Quality**: Refactored email retrieval functions for improved readability

### v1.0.0
- **Initial Release**: Complete MCP server implementation
- **Core Features**: Email retrieval, search, composition, and batch operations
- **Security**: User confirmation for email sending
- **Performance**: Email caching and timeout handling
- **CLI**: Interactive and command-line interfaces

## ⚠️ Limitations

### Email Address Search

Due to Microsoft Exchange's Distinguished Name (DN) format for internal email addresses, the search functionality for sender and recipient fields has the following limitations:

- **Sender/Recipient Name Search**: Only searches by display name, not email address
- **Exchange Format**: Internal email addresses are stored in DN format (e.g., `/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP...`)
- **Workaround**: Use display names (e.g., "John Doe") instead of email addresses for sender/recipient name searches
- **Subject Search**: Full text search still works for email subjects

**Example**:
- ✅ Correct: Search for "John Doe" in sender name field
- ❌ Incorrect: Search for "john.doe@example.com" in sender name field

### Planned Features
- [ ] Attachment upload/download support
- [ ] Calendar integration
- [ ] Contact management
- [ ] Email rules/filters
- [ ] Multi-account support
- [ ] Webhook notifications

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

## 🆘 Support

- **Issues**: [GitHub Issues](https://github.com/marlonluo2018/outlook-mcp-server/issues)
- **Discussions**: [GitHub Discussions](https://github.com/marlonluo2018/outlook-mcp-server/discussions)
- **Documentation**: This README and [UPDATED_SEARCH_GUIDE.md](UPDATED_SEARCH_GUIDE.md)