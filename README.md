# 🤖 Outlook MCP Server: Your AI Email Assistant

**Complete Outlook management with AI-powered email, folder, and policy control**

[![Star](https://img.shields.io/github/stars/marlonluo2018/outlook-mcp-server?style=for-the-badge&label=Star%20this%20project&color=yellow)](https://github.com/marlonluo2018/outlook-mcp-server)

> ⭐ **Love this project? Give it a star!** Your support helps us improve and reach more users who need AI email assistance.

### 🎯 **Core Management Features**
- **📧 Email Management**: AI-powered search, compose, reply, delete, and batch operations
- **📁 Folder Management**: Intelligent organization, creation, and workflow automation  
- **🏢 Policy Management**: Enterprise retention and compliance controls

## 🚀 What is This?

The Outlook MCP Server is your personal AI email assistant that connects Microsoft Outlook with powerful language models. It's not just another email tool - it's your complete Outlook management system with AI-powered email, folder, and policy control.

**Think of it as:**
- Your personal email analyst that understands your inbox
- A smart assistant that helps you draft perfect replies
- An AI-powered search engine for your email history
- Your complete Outlook management partner
- **Email Management**: Search, compose, reply, delete, and batch operations
- **Folder Management**: Create, move, and organize with intelligent workflows
- **Policy Management**: Enterprise-grade retention and compliance controls

## ✨ Why You'll Love It

### 🤖 **AI-Powered Email Understanding**
- **Smart Search**: "Find emails about project deadlines from last month"
- **Email Summarization**: Get quick summaries of long email threads
- **Context-Aware Replies**: AI understands the conversation and drafts appropriate responses
- **Intelligent Filtering**: Find what matters most in your crowded inbox

### 💬 **Natural Language Commands**
Talk to your AI assistant like you would to a human:
- "Show me recent emails from my boss"
- "Find all project updates from the last 2 weeks"
- "Draft a reply to John about the meeting reschedule"
- "Summarize this email thread for me"
- "Delete this spam email for me"
- "Forward this meeting invitation to all participants from my contacts.csv"

### � **Complete Folder Management**
- **Smart Organization**: Create, move, and organize folders with AI guidance
- **Intelligent Workflows**: Discover folder structure before operations
- **Nested Support**: Handle complex folder hierarchies up to 3 levels deep
- **Bulk Operations**: Move emails and folders efficiently

### 🏢 **Enterprise Policy Management**
- **Retention Policies**: Assign Exchange retention policies to emails
- **Compliance Controls**: Enterprise-grade policy assignment and verification
- **Policy Discovery**: Browse available policies before assignment
- **Multi-method Support**: Multiple assignment approaches for compatibility

### �🔌 **Seamless Integration**
- Works with any MCP-compatible AI assistant (Claude, GPT, etc.)
- Direct integration with Microsoft Outlook
- No complicated setup - just install and connect
- Windows-native performance with COM interface

## 🎯 Quick Start (5 Minutes)

### Prerequisites
- **Python 3.8+** ([Download here](https://python.org/downloads/))
- **Microsoft Outlook 2016+** (must be installed on your machine)
- **Windows 10+** (required for Outlook integration)
- **Python Dependencies**: `fastmcp` (MCP server framework) and `pywin32` (Outlook integration)
  - Automatic with UVX method, manual install with `pip install -r requirements.txt` for other methods

### Complete Setup Guide

Choose the installation method that best fits your needs:

#### 🚀 **Method 1: UVX (Recommended)** - **Automatic Dependency Management**
**Purpose**: Best for most users - UVX handles Python dependencies automatically without manual installation

**Step 1: Clone and Run**
```bash
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server
uvx --with "pywin32>=226" --with-editable "c:\Project\outlook-mcp-server" outlook-mcp-server
```

**Note**: UVX automatically installs dependencies from `requirements.txt` including `fastmcp` and `pywin32`.

**Step 2: Configure Your AI Assistant**
Use this configuration in your MCP client settings:
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

**Quick Setup**: Use the provided `mcp-config-uvx.json` file

#### 🔧 **Method 2: Standard Installation** - **Traditional Python Package**
**Purpose**: For users who prefer traditional Python package management and want to install the package permanently

**Step 1: Clone and Install Dependencies**
```bash
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server

# Install required Python dependencies
pip install -r requirements.txt

# Install build tools if needed
pip install build twine

# Build the package
python -m build

# Install the package
pip install .

# Run the server
python -m outlook_mcp_server
```

**Step 2: Configure Your AI Assistant**
Use this configuration in your MCP client settings:
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

**Quick Setup**: Use the provided `mcp-config-python.json` file

**Alternative Installation**: Pre-built distribution files are included in the `dist/` folder:
```bash
pip install dist/outlook_mcp_server-*.whl
```

#### 🔬 **Method 3: Direct Source (Development)** - **For Developers**
**Purpose**: For developers who want to modify the code or run directly from source without building

**Step 1: Clone and Run Directly**
```bash
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server

# Install dependencies
pip install -r requirements.txt

# Run directly from source
python outlook_mcp_server/__main__.py
```

**Step 2: Configure Your AI Assistant**
Use this configuration in your MCP client settings:
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

**Quick Setup**: Use the provided `mcp-config-direct.json` file

## 🤖 AI Assistant System Prompt

Your AI assistant uses a specialized prompt template that defines its role and behavior. This ensures safe and effective email management.

### Using the Prompt Template

The system prompt is defined in `agent_prompt_template.md`. To customize it for your needs:

1. **Edit the template file** to personalize the assistant's behavior
2. **Replace placeholders** like `[User Name]` and `[User Email]` with your information
3. **The assistant will follow** the structured workflows and safety constraints defined in the template

### Key Safety Features
- **Never sends emails** without explicit user confirmation
- **Always asks for clarification** when information is unclear
- **Follows structured workflows** for searching, summarizing, and drafting
- **Maintains user control** over all email actions

## 🎮 How It Works: Your AI Email Assistant in Action

Following the system prompt guidelines, your AI assistant follows a structured workflow to help you manage emails efficiently. Here's how it works from start to finish:

### 🔄 **The Complete Workflow**

#### **Phase 1: Smart Email Search & Discovery**
Your AI assistant helps you find emails using natural language commands:

**What you can ask:**
- "Find emails about budget approval from last quarter"
- "Show me urgent emails I haven't replied to"
- "Search for attachments from specific senders"
- "Find all emails related to the marketing campaign"

**How it works:**
- **Search Tools Available:**
  - `list_recent_emails_tool()` - Load recent emails from last X days
  - `search_email_by_subject_tool()` - Search email subjects only
  - `search_email_by_sender_name_tool()` - Search by sender name
  - `search_email_by_recipient_name_tool()` - Search by recipient name
  - `search_email_by_body_tool()` - Search email body content (slower)
- Results are loaded into the persistent unified cache (5 emails per page for browsing, with backend limit of 1000 entries)

#### **Phase 2: Browse & Preview Emails**
Emails are displayed in a clean, consistent format with exactly 5 emails per page:
```
Subject: [Email Subject]
From: [Sender Name]
To: [Recipient List]
Received: [Date/Time]
Status: [Read/Unread]
Attachments: [Attachment Count]
```

**Viewing Tools:**
- `view_email_cache_tool()` - View 5 emails per page
- `get_email_by_number_tool()` - Get full details of specific email with configurable modes

#### **Phase 3: AI Analysis & Summary**
When you select an email, the AI provides intelligent insights:
- **One-sentence overview** of the email content
- **Status label** (Awaiting Reply, Urgent, For Information, FYI, Completed)
- **Bulleted action items** extracted from the email
- **Deadlines/commitments** automatically identified

#### **Phase 4: AI-Powered Email Composition**
The AI helps you draft perfect responses using a 4-step approach:

**What you can ask:**
- "Draft a professional reply to this meeting invitation"
- "Help me write a follow-up email about the project status"
- "Create a polite response declining this request"
- "Summarize my unread emails from this morning"

**How it works:**
1. **Gather Information** - Confirms purpose, key points, recipients, tone
2. **Draft & Suggest** - Creates full draft + 3 improvement suggestions
3. **Iterate** - Applies changes + provides new suggestions until satisfied
4. **Send** - Requires explicit user confirmation before sending

**Email Composition Tools:**
- `reply_to_email_by_number_tool()` - Reply to an email
- `compose_email_tool()` - Compose new email

#### **Phase 5: AI-Powered Batch Email Operations**
The AI helps you send emails to multiple recipients efficiently:

**What you can ask:**
- "Forward this meeting invitation to all participants in my contact list"
- "Send this project update to my entire team using recipients.csv"
- "Broadcast this announcement to all employees with additional context"

**How it works:**
1. **Template Selection** - Uses an email from your cache as template
2. **Recipient Loading** - Reads email addresses from CSV file with 'email' column
3. **Batch Processing** - Automatically splits large lists into batches of 500
4. **Content Customization** - Adds your custom text before original email
5. **Safe Sending** - Sends via BCC to protect recipient privacy

**Batch Email Tools:**
- `batch_forward_email_tool()` - Forward template email to multiple recipients from CSV

**🔒 Key Safety Feature**: No emails are sent without your explicit approval!

## 🔧 Available Tools

### Email Search & Loading
- `get_folder_list_tool()` - Lists all Outlook mail folders
- `list_recent_emails_tool(days=7, folder_name=None)` - Load recent emails
- `search_email_by_subject_tool(search_term, days=7, folder_name=None, match_all=True)` - Search by subject
- `search_email_by_sender_name_tool(search_term, days=7, folder_name=None, match_all=True)` - Search by sender
- `search_email_by_recipient_name_tool(search_term, days=7, folder_name=None, match_all=True)` - Search by recipient
- `search_email_by_body_tool(search_term, days=7, folder_name=None, match_all=True)` - Search by body

### Folder Management
**⚠️ Important Workflow**: For folder operations, always start with `get_folder_list_tool()` to discover the folder structure first.

- `get_folder_list_tool()` - **REQUIRED FIRST STEP** - Lists all Outlook mail folders to understand structure
- `move_folder_tool(source_folder_path, target_parent_path)` - Move folders between locations (use full paths from folder list)
- `create_folder_tool(folder_name, parent_folder_name=None)` - Create new folders (supports nested paths like "Inbox/SubFolder1/SubFolder2")
- `remove_folder_tool(folder_name)` - Delete folders and their contents (use full path from folder list)
- `move_email_tool(email_number, target_folder_name)` - Move emails between folders (requires full folder path from folder list)

**Folder Operation Workflow:**
1. **Discover Structure**: Use `get_folder_list_tool()` to see available folders
2. **Identify Paths**: Note the full folder paths (e.g., "user@company.com/Inbox/FolderName")
3. **Execute Operation**: Use the appropriate tool with the correct full path

### Email Viewing & Browsing
- `view_email_cache_tool(page=1)` - View 5 emails per page
- `get_email_by_number_tool(email_number, mode="basic|enhanced|lazy")` - **Unified tool** for configurable email retrieval with media support
- `delete_email_by_number_tool(email_number)` - Move email to Deleted Items folder (soft delete)

### Email Composition (Requires User Confirmation)
- `reply_to_email_by_number_tool(email_number, reply_text, to_recipients=None, cc_recipients=None)` - Reply to email
- `compose_email_tool(recipient_email, subject, body, cc_email=None)` - Compose new email

### Email Batch Operations (Requires User Confirmation)
- `batch_forward_email_tool(email_number, csv_path, custom_text="")` - Forward email to multiple recipients from CSV file

**Batch Email Feature:**
- Uses an email from your cache as a template
- Reads recipient email addresses from a CSV file with 'email' column
- Automatically splits large lists into batches of 500 (Outlook BCC limit)
- Adds custom text before the original email content
- Preserves email formatting with consistent break lines
- Sends via BCC to protect recipient privacy

### Policy Management (Enterprise Features)
**⚠️ Important Workflow**: For policy operations, always start with `get_policies_tool()` to discover available policies first.

- `get_policies_tool()` - **REQUIRED FIRST STEP** - Discover available Exchange retention policies
- `assign_policy_tool(email_number, policy_name)` - Assign a policy to an email (use exact policy name from discovery)
- `get_email_policies_tool(email_number)` - Verify which policies are assigned to an email

**Policy Management Workflow:**
1. **Discover Policies**: Use `get_policies_tool()` to see available enterprise policies
2. **Assign Policy**: Use `assign_policy_tool()` with the exact policy name from discovery
3. **Verify Assignment**: Use `get_email_policies_tool()` to confirm the policy was applied

**Policy Operation Guidelines:**
- **Policy Names**: Use exact names from `get_policies_tool()` (e.g., "1 Year (Enterprise)", "Never Delete")
- **Email Selection**: Policies are assigned to specific emails by their cache number
- **Enterprise Requirement**: Policy features require Exchange/Office 365 enterprise accounts
- **Assignment Methods**: The system tries multiple methods (direct properties, categories, user properties)

**Search Behavior:**
- All search tools support `match_all=True` (AND logic) or `match_all=False` (OR logic)
- Email body searching is slower than other fields
- Search terms can include colons as regular text

## ⚠️ Known Issues & Limitations

### Current Limitations
- **Folder Level Limitation**: Maximum 3 folder levels supported (e.g., 'Inbox/SubFolder1/SubFolder2'). This affects nested folder creation with mailbox-specific paths.
- **Folder Deletion Delay**: Deleted folders may require multiple list checks to disappear from the folder list due to Outlook's internal caching.
- **Full Path Requirements**: Email and folder operations require full mailbox paths (e.g., "user@company.com/Inbox/FolderName") rather than relative paths.

### Recent Fixes Implemented
- **Folder Moving**: Fixed COM interface error by using `MoveTo` method instead of `Move`
- **Email Moving**: Enhanced to support full folder paths for accurate targeting
- **Error Handling**: Improved error reporting for Outlook operations

## 🎯 Real-World Examples

### Scenario 1: Busy Executive
**Problem**: Too many emails, not enough time
**Solution**: "Hey AI, summarize my urgent emails and draft replies for the top 3"

### Scenario 2: Project Manager
**Problem**: Need to track project communications
**Solution**: "Find all emails about Project Phoenix and create a status summary"

### Scenario 3: Sales Professional
**Problem**: Follow up with leads efficiently
**Solution**: "Draft personalized follow-up emails for all unread sales inquiries"

### Scenario 4: Event Organizer
**Problem**: Need to forward announcements to large contact lists
**Solution**: "Use the batch_forward_email_tool to forward this meeting invitation to all attendees from my contacts.csv file"

### Scenario 5: HR Manager
**Problem**: Forward company-wide communications to employees
**Solution**: "Use the batch_forward_email_tool to forward this policy update email to all employees in employee_list.csv"

## 💻 CLI Interface (Human Interaction)

For direct human interaction, use the CLI interface:

```bash
# Start the CLI
python cli_interface.py

# Available commands in CLI:
- search_emails - Search emails with various criteria
- view_cache - Browse cached emails (5 per page)
- get_email - View full email details
- reply_email - Reply to an email (requires confirmation)
- compose_email - Create new email (requires confirmation)
- batch_forward_email - Forward email to multiple recipients from CSV (requires confirmation)
- move_folder - Move folders between locations
- create_folder - Create new folders
- remove_folder - Delete folders
- move_email - Move emails between folders
```

### Unified Email Retrieval Architecture

The server now features a **unified email retrieval architecture** that consolidates all email access functionality into a single, configurable interface:

**🎯 Three Retrieval Modes:**
- **Basic Mode** (`"basic"`): Fast, lightweight retrieval for email listings and summaries
- **Enhanced Mode** (`"enhanced"`): Full media support with attachments, inline images, and comprehensive metadata
- **Lazy Mode** (`"lazy"`): Intelligent mode that adapts based on cached data for optimal performance

**🔧 Unified Tool Interface:**
```python
# Single tool handles all email retrieval scenarios
get_email_by_number_tool(email_number, mode="basic|enhanced|lazy", include_attachments=True, embed_images=True)
```

**✨ Key Benefits:**
- **Simplified API**: One tool handles all email retrieval scenarios
- **Performance Optimization**: Choose the right mode for your use case
- **Consolidated Architecture**: Replaces multiple legacy tools with unified interface
- **Future-Proof**: Easy to extend with new modes and features
- **Resource Efficient**: Lazy mode minimizes unnecessary data fetching

**📖 Complete Documentation:**
- See the unified tool implementation in `email_retrieval.py` for technical details

### Enhanced Email Media Support

The unified email retrieval system includes comprehensive media support:

**Features:**
- **Attachment Content Extraction**: Automatically extracts base64-encoded content for embeddable files (images, text files)
- **Inline Image Embedding**: Replaces CID references in HTML emails with embedded data URIs
- **Comprehensive Metadata**: Provides MIME types, file sizes, content IDs for all attachments
- **Content Preview**: Shows previews for small text files and images

**Use Cases:**
- View emails with embedded images without external dependencies
- Extract attachment content for analysis or processing
- Get complete email context including visual elements
- Analyze email structure and inline content

## 🚀 Recent Performance Improvements

### Email Retrieval Optimization
The latest updates include significant performance enhancements to email retrieval and filtering:

**🔧 Enhanced Date Filtering:**
- Robust date parsing with fallback mechanisms for various date formats
- Intelligent validation against different time zones and local system settings
- Graceful handling of invalid dates with automatic filtering adjustments
- Comprehensive logging for troubleshooting date-related issues

**⚡ Query Performance Improvements:**
- Optimized batch processing for large email volumes
- Configurable timeouts to prevent long-running operations (default: 60 seconds)
- Enhanced error recovery with retry mechanisms
- Detailed performance metrics and logging

**📊 Measured Performance Results:**
- 1-day queries: ~25-30 emails/second (complete in 2-3 seconds for ~70 emails)
- 3-day queries: ~20-25 emails/second (complete in 5-6 seconds for ~120 emails)
- Consistent performance across multiple query executions
- Efficient handling of date-based filtering constraints

**🛠 Backend Optimizations:**
- Improved COM interface handling for Outlook integration
- Enhanced caching mechanisms with persistent storage
- Better memory management for large email datasets
- Streamlined error handling and recovery procedures

## 🛠 Technical Details

### Architecture
- **Frontend**: MCP-compatible AI assistants (Claude, GPT, etc.)
- **Backend**: Python-based MCP server with Outlook COM integration
- **Communication**: Standard MCP protocol over stdio
- **Caching**: Unified file-based cache system (JSON files stored in `%LOCALAPPDATA%\outlook_mcp_server`)
  - Automatic persistence handles UVX process isolation by saving cache state to disk
  - 1-hour cache expiry with size management (MAX_CACHE_SIZE=1000 entries)
  - Each client runs its own isolated server instance, ensuring cache separation
  - Automatic save/load on cache operations with error handling
  - LRU-like eviction of oldest entries when cache limit is reached

### Dependencies
- **Core**: Python 3.8+, pywin32>=226
- **MCP**: FastMCP framework
- **Outlook**: Microsoft Outlook 2016+
- **Platform**: Windows 10+

### Performance
- **Search Speed**: Subject/sender search: fast, Body search: slower
- **Cache Size**: 5 emails per page for optimal browsing, with backend cache limit of 1000 entries
- **Persistence**: Automatic file-based storage ensures cache survives process restarts (especially important for UVX configuration)
- **Memory**: Efficient caching with automatic LRU cleanup when cache limit is reached
- **Optimized Retrieval**: Advanced email filtering with performance optimizations:
  - Date-based filtering with configurable time windows (1-30 days)
  - Intelligent fallback systems for invalid date formats
  - Batch processing for large email volumes
  - Performance metrics: ~25-30 emails/second for 1-day queries, ~20-25 emails/second for 3-day queries
  - Timeout protection (60 seconds default) to prevent long-running operations

## 🤝 Contributing

We welcome contributions! Here's how you can help:

1. **Report Issues**: Found a bug? Open an issue with details
2. **Suggest Features**: Have ideas for improvement? Share them!
3. **Code Contributions**: Fork the repo and submit pull requests
4. **Documentation**: Help improve docs for better user experience

## � Star History

[![Star History Chart](https://api.star-history.com/svg?repos=marlonluo2018/outlook-mcp-server&type=Date)](https://star-history.com/#marlonluo2018/outlook-mcp-server&Date)

## �📄 License

This project is open source and available under the MIT License.

## � Acknowledgments

- Built with [FastMCP](https://github.com/modelcontextprotocol/fastmcp) framework
- Inspired by the need for better AI-powered email management
- Thanks to all contributors and users who help improve this project

---

**Ready to transform your email experience? Give it a try and let us know what you think!**

> ⭐ **Don't forget to star the repository if you find this useful!**