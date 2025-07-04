# Outlook MCP Server

A Model Context Protocol (MCP) server that provides access to Microsoft Outlook email functionality, allowing LLMs and other MCP clients to read, search, and manage emails through a standardized interface.

## Features

- **Folder Management**: List available mail folders in your Outlook client
- **Email Listing**: Retrieve emails from specified time periods
- **Email Search**: Search emails by contact name, keywords, or phrases with OR operators
- **Email Details**: View complete email content, including attachments
- **Email Composition**: Create and send new emails
- **Email Replies**: Reply to existing emails

## Prerequisites

- Windows operating system
- Python 3.11 or later
- Microsoft Outlook installed and configured with an active account
- Claude Desktop or another MCP-compatible client

## Installation

1. Clone or download this repository
2. Install required dependencies:

```bash
pip install mcp>=1.2.0 pywin32>=305
```

3. Configure Claude Desktop (or your preferred MCP client) to use this server

## Usage

### Configure your MCP client to start it automatically. 

Add this to your MCP client config:

```json
{
  "mcpServers": {
    "outlook": {
      "isActive": true,
      "name": "outlook",
      "description": "Outlook Tools",
      "command": "python",
      "args": [
        "${workspaceFolder}/outlook_mcp_server.py"

      ]
    }
  }
}
```

### Available Tools

The server provides the following tools with detailed functionality:

1. **list_folders**:  
   - Lists all available mail folders in Outlook  
   - Shows folder hierarchy up to 3 levels deep  
   - Returns formatted list of folders and subfolders  

2. **list_recent_emails**:  
   - Lists email titles from specified number of days (1-30)  
   - Can specify folder to search (defaults to Inbox)  
   - Caches results for detailed viewing  
   - Returns number of emails found and instructions to view them  

3. **search_emails(search_terms, match_all=False, folder="Inbox")**:
   - Searches emails by contact name, keyword, or exact phrases (use quotes)
   - Supports AND/OR operators between terms (match_all=True for AND)
   - Searches subject, sender name, and body content
   - Returns number of matches and instructions to view them

4. **view_email_cache**:  
   - Views cached emails in pages of 5  
   - Shows email subject, sender, received time, and read status  
   - Provides navigation to next/previous pages  
   - Requires prior use of list_recent_emails or search_emails  

5. **get_email_by_number**:
   - Retrieves detailed content of a specific email
   - Shows full email body (HTML or plain text), recipients, and attachments
   - Requires email number from cached listing
   - Supports overriding reply recipients when replying

6. **compose_email(recipients, subject, body, cc=None, bcc=None, html=False)**:
   - Creates and sends new emails
   - Supports multiple recipients (comma-separated)
   - HTML or plain text formatting
   - CC/BCC recipients supported

7. **reply_to_email(email_number, reply_text, html=False, recipients=None)**:
   - Replies to existing emails
   - Overrides default reply recipients when specified
   - Preserves original message formatting
   - HTML or plain text replies supported
   - Provides option to reply to the email  

6. **reply_to_email_by_number**:  
   - Replies to a specific email by its number  
   - Uses email number from cached listing  
   - Sends reply with specified text content  
   - Maintains original email thread  
   - Optional: Specify custom recipients for "To" and "CC" fields  
     - Default behavior replies to all (sender + CC)  
     - Allows overriding default recipients if needed  

7. **compose_email**:  
   - Creates and sends a new email  
   - Supports recipient, subject, and body content  
   - Optional CC field for additional recipients  
   - Handles email sending through Outlook client

### Example Workflow

1. Use `list_folders` to see all available mail folders
2. Use `list_recent_emails` to view recent emails (e.g., from last 7 days)
3. Use `view_email_cache(page=1)` to browse through matching emails in pages of 1
4. Use `get_email_by_number` to view a complete email
5. Use `reply_to_email_by_number` to respond to an email

## Examples

### Listing Recent Emails
```
Could you show me my unread emails from the last 3 days?
```

### Searching for Emails
```
Search for emails about "project update OR meeting notes" in the last week
```

### Viewing Cached Emails
```
Show me page 2 of my cached emails
```

### Reading Email Details
```
Show me the details of email #2 from the list
```

### Replying to an Email
```
Reply to email #3 with: "Thanks for the information. I'll review this and get back to you tomorrow."
```

### Composing a New Email
```
Send an email to john.doe@example.com with subject "Meeting Agenda" and body "Here's the agenda for our upcoming meeting..."
```

## Troubleshooting

- **Connection Issues**: Ensure Outlook is running and properly configured
- **Permission Errors**: Make sure the script has permission to access Outlook
- **Search Problems**: For complex searches, try using OR operators between terms
- **Email Access Errors**: Check if the email ID is valid and accessible
- **Server Crashes**: Check Outlook's connection and stability

## Security Considerations

This server has access to your Outlook email account and can read, send, and manage emails. Use it only with trusted MCP clients and in secure environments.

## Limitations

- Currently supports text emails only (not HTML)
- Maximum email history is limited to 30 days
- Search capabilities depend on Outlook's built-in search functionality
- Only supports basic email functions (no calendar, contacts, etc.)
