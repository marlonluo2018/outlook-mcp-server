# Outlook MCP Server

> **⚠️ Windows-Only Application**  
> This MCP server requires Windows 10/11 and Microsoft Outlook due to COM automation dependencies.

A comprehensive Model Context Protocol (MCP) server that provides secure, efficient access to Microsoft Outlook functionality through Python COM automation. This server enables AI assistants and applications to interact with Outlook emails, folders, and contacts programmatically while maintaining enterprise-grade security and performance.

## 🚀 Overview

The Outlook MCP Server bridges the gap between AI systems and Microsoft Outlook, providing a standardized interface for email operations. Built on the Model Context Protocol, it offers both programmatic API access and interactive CLI usage patterns.

### Key Capabilities

- **Email Operations**: Search, retrieve, compose, and reply to emails
- **Advanced Search**: Intelligent word proximity search for more accurate results
- **Folder Management**: Browse and access all Outlook folders
- **Batch Processing**: Send bulk emails with CSV-based recipient lists (CLI only)
- **Caching System**: Intelligent email caching for performance optimization
- **Security**: Built-in user confirmation for email sending operations
- **Error Handling**: Comprehensive error handling with retry mechanisms

## 📋 Requirements

### System Requirements

- **Operating System**: Windows 10/11 (required for Outlook COM automation)
- **Python**: 3.8 or higher
- **Microsoft Outlook**: 2016 or later, properly configured and running
- **COM Access**: Outlook must be accessible via COM (default for most installations)

### Dependencies

- `fastmcp==2.11.0`: MCP server framework
- `pywin32==306`: Windows COM automation
- Standard library: `argparse`, `csv`, `datetime`, `logging`, `typing`

## 🛠️ Installation

### Quick Start

```bash
# Clone the repository
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server

# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install requirements
pip install -r requirements.txt
```

### MCP Configuration

Add to your MCP client configuration (e.g., Claude Desktop settings.json):

```json
{
  "mcpServers": {
    "outlook": {
      "type": "stdio",
      "command": "python",
      "args": ["C:\\Project\\outlook-mcp-server\\outlook_mcp_server.py"]
    }
  }
}
```

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
python cli_interface.py
```

Note: For LLM integration, use the MCP server interface instead.

### Development Installation

```bash
# Install development dependencies
pip install -e ".[dev]"
```

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
python cli_interface.py
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

## 📋 Changelog

### v1.2.0 (Current)
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
