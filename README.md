# 🤖 Outlook MCP Server

**AI-powered email management for Microsoft Outlook** - Search, compose, organize, and batch forward emails with natural language commands.

<div align="center">

⭐ **This saved you time? [Star us](https://github.com/marlonluo2018/outlook-mcp-server/stargazers) - takes 2 seconds, helps thousands of Outlook users find this AI email assistant!** ⭐

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Windows](https://img.shields.io/badge/Windows-10%2B-0078D6)](https://www.microsoft.com/windows)

</div>

## 🚀 Quick Start (2 Minutes)

### What You'll Get
- **Smart Email Search**: "Find emails about budget approval from last week"
- **AI Email Writing**: Draft replies with context-aware suggestions  
- **Easy Organization**: Create folders and move emails with simple commands
- **Batch Forwarding**: Send emails to 100s of recipients in minutes, not hours

### Requirements
- ✅ Python 3.8+
- ✅ Microsoft Outlook 2016+ (must be running)
- ✅ Windows 10+

### Installation & Setup

**Method 1: UVX (Recommended - Auto Dependencies)**
```bash
# 1. Install
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server
uvx --with "pywin32>=226" --with-editable "." outlook-mcp-server

# 2. Configure your AI assistant
# Use this in your MCP client settings:
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["--with", "pywin32>=226", "--with-editable", ".", "outlook-mcp-server"]
    }
  }
}
```

**Method 2: Standard Python**
```bash
# 1. Install
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server
pip install -r requirements.txt
python -m outlook_mcp_server

# 2. Configure your AI assistant
# Use this in your MCP client settings:
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["-m", "outlook_mcp_server"]
    }
  }
}
```

**Method 3: Direct Source (Development)**
```bash
# 1. Install
git clone https://github.com/marlonluo2018/outlook-mcp-server.git
cd outlook-mcp-server
pip install -r requirements.txt

# 2. Configure your AI assistant
# Use this in your MCP client settings:
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["C:\\Project\\outlook-mcp-server\\outlook_mcp_server\\__main__.py"]
    }
  }
}
```

### Test Your Setup ✅
Ask your AI assistant: "Show me my recent emails" - if it works, you're ready!

### Configuration Troubleshooting 🛠️

**Common Issues & Solutions:**

| Problem | Check This | Solution |
|---------|------------|----------|
| **"uvx not found"** | Is UV installed? | `pip install uv` then retry |
| **"python not found"** | Python in PATH? | Use full path like `C:\Python39\python.exe` |
| **"Outlook not running"** | Outlook window open? | Start Outlook first, then restart MCP |
| **"Permission denied"** | Admin rights? | Run terminal as administrator |
| **"Module not found"** | Dependencies installed? | `pip install -r requirements.txt` |

**Configuration Verification:**
```bash
# Test your setup before connecting to AI
python -c "import outlook_mcp_server; print('✅ Server module loaded')"

# Test Outlook connection (Windows only)
python -c "import win32com.client; outlook = win32com.client.Dispatch('Outlook.Application'); print('✅ Outlook connected')"
```

**MCP Client-Specific Setup:**

**Claude Desktop:**
1. Open Claude Desktop settings
2. Find "MCP Servers" section
3. Click "Add Server" → "Custom"
4. Paste the JSON configuration
5. Restart Claude Desktop

**Other MCP Clients:**
- Look for "MCP Configuration" or "Server Settings"
- Add the JSON to your client's config file
- Usually located at: `~/.config/[client]/mcp.json`

**Still Stuck?** [Report an issue](https://github.com/marlonluo2018/outlook-mcp-server/issues) with your error message and setup details.

## 🎯 Core Features

### Email Management
- **Search**: Find emails by subject, sender, content, or date range
- **Compose**: Write new emails with AI assistance
- **Reply**: Smart replies that understand conversation context
- **Batch Forward**: Send emails to 100s of recipients from CSV files (saves hours!)

### Folder Management  
- **List**: See all your Outlook folders
- **Create**: Make new folders with simple commands
- **Move**: Organize emails between folders
- **Delete**: Remove folders (careful - this is permanent!)

## 📧 Batch Forwarding: Save Hours on Email Distribution

### Real-World Use Cases

**🎯 Team Updates**
- Forward weekly reports to your entire team
- Send meeting notes to all participants
- Distribute project updates to stakeholders

**📊 Marketing Campaigns**  
- Send newsletters to subscriber lists
- Forward promotional emails to customer segments
- Distribute event invitations to contact groups

**🏢 Corporate Communications**
- Send policy updates to all employees
- Forward training materials to departments
- Distribute announcements to company distribution lists

### How It Works

**1. Prepare Your CSV File**
```csv
email
john@company.com
jane@company.com
team@company.com
```

**2. Use Natural Language**
```
"Forward this meeting invite to everyone in my contacts.csv"
"Send this project update to my team list"
"Distribute this newsletter to subscribers.csv"
```

**3. AI Handles the Rest**
- Automatically splits large lists (max 500 per batch)
- Sends via BCC to protect recipient privacy
- Adds your custom message before original email
- Provides delivery confirmation

### Time Savings
**Manual forwarding**: 100 emails × 30 seconds = 50 minutes
**Batch forwarding**: 30 seconds setup + 2 minutes processing = 2.5 minutes
**You save**: 47.5 minutes per batch!

## 🔄 How It Works

### Simple Workflow
1. **Load emails**: "Show me recent emails" → Emails appear in cache
2. **Browse results**: View 5 emails per page with clear formatting
3. **Take action**: Reply, move, delete, or get AI summary
4. **Confirm before sending**: AI always asks before sending emails

### AI Assistant Behavior
- **Understands natural language**: "Find urgent emails from my boss"
- **Shows email summaries**: One-line overview + key action items
- **Drafts with context**: Replies understand the conversation
- **Never sends without permission**: Always confirms before sending

## 🔧 Common Commands

Try these with your AI assistant:

```
"Show me emails from last 3 days"
"Find emails about project updates" 
"Draft a reply to John about rescheduling"
"Create a folder called 'Work Projects'"
"Move email #3 to the Archive folder"
```

## 🛠️ Essential Tools

### Email Search & Loading
- `list_recent_emails_tool(days=7)` - Load recent emails (max: 30 days)
- `search_email_by_subject_tool("search term")` - Search email subjects
- `search_email_by_sender_name_tool("sender name")` - Search by sender
- `search_email_by_body_tool("search term")` - Search email content (slower)

### Email Actions
- `view_email_cache_tool(page=1)` - Browse loaded emails (5 per page)
- `get_email_by_number_tool(email_number)` - Get full email details
- `reply_to_email_by_number_tool(email_number, "reply text")` - Reply to email
- `compose_email_tool("recipient@email.com", "subject", "body")` - Send new email

### Folder Operations
- `get_folder_list_tool()` - **Always use first** to see available folders
- `create_folder_tool("folder name")` - Create new folder
- `move_email_tool(email_number, "target folder")` - Move email between folders
- `move_folder_tool("source", "target")` - Move folders

### Safety Notes
- **Always check folder list first** before moving/deleting
- **Emails are cached** - use numbers from cache for operations
- **Never sends without confirmation** - AI always asks before sending

## ⚡ Performance

- Loads 100 emails in ~2 seconds
- Searches complete in real-time
- All processing happens locally (your data stays private)

## 🛡️ Privacy & Security

- **100% Local Processing**: No data leaves your computer
- **No Cloud Services**: Works entirely offline
- **Secure by Design**: Uses your existing Outlook installation

## 📚 Need Help?

- Check the [agent prompt template](agent_prompt_template.md) for AI assistant setup
- See [configuration examples](mcp-config-uvx.json) for different installation methods
- Report issues on the [GitHub repository](https://github.com/marlonluo2018/outlook-mcp-server/issues)

---

<div align="center">

**⭐ Love saving time on email? [Star this repo](https://github.com/marlonluo2018/outlook-mcp-server/stargazers) - helps 10,000+ Outlook users discover AI email management! ⭐**

**💡 Quick star tip: Click the ⭐ button above - it takes 2 seconds and supports open-source email AI!**

</div>