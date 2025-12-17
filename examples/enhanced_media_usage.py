"""
Enhanced Email Media Support - Usage Examples

This example demonstrates how to use the new get_email_with_media_tool function
to retrieve emails with full media support including inline images and attachments.
"""

# Example 1: Basic usage with default settings
"""
Get email with media support (includes attachments and embeds images):

result = get_email_with_media_tool(1)
print(result["text"])
"""

# Example 2: Get email without attachment content (faster)
"""
Get email metadata only, without extracting attachment content:

result = get_email_with_media_tool(1, include_attachments=False)
print(result["text"])
"""

# Example 3: Get email without image embedding
"""
Get email with attachment info but without inline image processing:

result = get_email_with_media_tool(1, embed_images=False)
print(result["text"])
"""

# Example 4: Complete workflow example
"""
Complete workflow for processing emails with media:

# Step 1: Load emails from inbox
from outlook_mcp_server import list_recent_emails_tool
result = list_recent_emails_tool("Inbox", days=7)
print(f"Found: {result}")

# Step 2: Get enhanced email with media
from outlook_mcp_server import get_email_with_media_tool
enhanced_result = get_email_with_media_tool(1)
print(enhanced_result["text"])

# The output will include:
# - Subject, sender, recipients
# - Body content (plain text and HTML)
# - Attachments with content preview
# - Inline images embedded in HTML
# - Enhanced metadata (importance, sensitivity, etc.)
"""

# Example 5: Comparing with basic email retrieval
"""
Comparison between basic and enhanced email retrieval:

# Basic retrieval (existing functionality)
basic_result = get_email_by_number_tool(1)
print("Basic retrieval output:")
print(basic_result["text"])
print("\n" + "="*50 + "\n")

# Enhanced retrieval (new functionality)  
enhanced_result = get_email_with_media_tool(1)
print("Enhanced retrieval output:")
print(enhanced_result["text"])

# Key differences:
# 1. Enhanced version shows attachment content for embeddable files
# 2. Enhanced version processes inline images in HTML
# 3. Enhanced version provides more detailed metadata
# 4. Enhanced version shows MIME types and file sizes
"""

# Example 6: Advanced usage for specific scenarios
"""
Advanced usage scenarios:

# Scenario 1: Analyze email with many attachments
result = get_email_with_media_tool(1)
# This will show all attachments with their types and sizes

# Scenario 2: Process email with inline images  
result = get_email_with_media_tool(2, embed_images=True)
# This will replace CID references with embedded data URIs

# Scenario 3: Quick preview without heavy content
result = get_email_with_media_tool(3, include_attachments=False)
# This shows attachment names and sizes but not content

# Scenario 4: Extract text content from attachments
result = get_email_with_media_tool(4)
# For small text files, you'll see content preview in output
"""