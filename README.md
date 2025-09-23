# Outlook MCP Server

> **⚠️ Windows-Only Application**  
> This MCP server requires Windows 10/11 and Microsoft Outlook due to COM automation dependencies.

A comprehensive Model Context Protocol (MCP) server that provides secure, efficient access to Microsoft Outlook functionality through Python COM automation. This server enables AI assistants and applications to interact with Outlook emails, folders, and contacts programmatically while maintaining enterprise-grade security and performance.

## 🚀 Overview

The Outlook MCP Server bridges the gap between AI systems and Microsoft Outlook, providing a standardized interface for email operations. Built on the Model Context Protocol, it offers both programmatic API access and interactive CLI usage patterns.

### Key Capabilities

- **Email Operations**: Search, retrieve, compose, and reply to emails
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

- `fastmcp`: MCP server framework
- `pywin32`: Windows COM automation
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
- "Search for emails about project updates"  
- "Find emails from the last 2 weeks"

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
- **Menu Option 3**: Search emails (enter search terms and filters)  

This clears any previous cache and loads your selected emails.

**Step 2: Browse Available Emails**  
Use **Menu Option 4** to:  
- View all emails currently in your cache  
- See a numbered list with subjects and senders  
- Navigate through multiple pages if needed  
- Note the email number you want to work with

**Step 3: View Email Details**  
Use **Menu Option 5** to:  
- Enter the email number you want to read  
- See the complete email content  
- Check attachments and recipient details  

**Step 4: Take Action on Email**  
Choose your action:  
- **Menu Option 6**: Reply to the email (enter email number)  
- **Menu Option 7**: Compose a new email  
- **Menu Option 8**: Send batch emails (using cached email as template)  

**📋 Common Usage Patterns**:

**To reply to an email**: 2 → 4 → 5 → 6  
(List emails → View cache → Read email → Reply)

**To search and respond**: 3 → 4 → 5 → 6  
(Search emails → View cache → Read email → Reply)

**To send batch emails**: 2 → 4 → 8  
(List emails → View cache → Send batch)

**⚠️ Important Notes**:

- **Cache Management**: Each time you use Option 2 or 3, the cache is cleared and reloaded
- **Email Numbers**: Always use the numbers shown in the current cache (Option 4)
- **Your Confirmation**: The system asks for confirmation before sending any emails
- **Session-Based**: Cache persists until you exit or load new emails
- **Menu Navigation**: Use Option 0 to exit safely

**Best for**: Users who prefer direct, menu-driven control over email operations or need batch email capabilities.

### Available Tools

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

#### 3. Search Emails

```json
{
  "tool": "search_emails_tool",
  "parameters": {
    "search_term": "meeting notes",
    "days": 14,
    "folder_name": "Inbox",
    "match_all": true
  }
}
```

#### 4. View Email Cache

```json
{
  "tool": "view_email_cache_tool",
  "parameters": {
    "page": 1
  }
}
```

#### 5. Get Email Details

```json
{
  "tool": "get_email_by_number_tool",
  "parameters": {
    "email_number": 3
  }
}
// Returns full email with body and attachments
```

#### 6. Reply to Email

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

#### 7. Compose New Email

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

#### 8. Batch Email Operations (Interactive Only)

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

## 📋 Changelog

### v1.0.0 (Current)
- **Initial Release**: Complete MCP server implementation
- **Core Features**: Email retrieval, search, composition, and batch operations
- **Security**: User confirmation for email sending
- **Performance**: Email caching and timeout handling
- **CLI**: Interactive and command-line interfaces

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
- **Documentation**: This README and inline code documentation
