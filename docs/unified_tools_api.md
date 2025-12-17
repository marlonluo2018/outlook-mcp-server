# Unified Email Retrieval Tools API Documentation

## Overview

The unified email retrieval architecture provides a single, configurable interface for accessing Outlook emails with different levels of detail and functionality. This documentation describes the API specifications for the unified tools.

## Architecture Modes

The unified system supports three distinct retrieval modes:

### 1. Basic Mode (`"basic"`)
- **Purpose**: Fast, lightweight email retrieval
- **Use Case**: Simple email listing, basic information display
- **Performance**: Fastest retrieval speed
- **Data Provided**: Essential email fields (subject, sender, body, recipients, basic metadata)

### 2. Enhanced Mode (`"enhanced"`)
- **Purpose**: Comprehensive email retrieval with full media support
- **Use Case**: Detailed email analysis, media-rich content, attachment processing
- **Performance**: Slower due to additional processing
- **Data Provided**: All email fields including:
  - Full HTML body content
  - Attachment content with base64 encoding
  - Inline image embedding
  - Extended metadata (importance, sensitivity, conversation info)
  - Categories and flags

### 3. Lazy Mode (`"lazy"`)
- **Purpose**: Intelligent mode that adapts based on cached data
- **Use Case**: Optimal performance with fallback to enhanced when needed
- **Performance**: Variable - uses cached data when available, falls back to enhanced when necessary
- **Data Provided**: Basic data from cache, enhanced data when cache is insufficient

## Tool Specifications

### `get_email_by_number_tool_unified`

**Purpose**: Unified tool for retrieving emails by cache position with configurable functionality.

**Parameters:**

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `email_number` | int | Yes | - | The 1-based position of the email in the cache |
| `mode` | str | No | `"basic"` | Retrieval mode: `"basic"`, `"enhanced"`, or `"lazy"` |
| `include_attachments` | bool | No | `true` | Whether to include attachment content (enhanced mode only) |
| `embed_images` | bool | No | `true` | Whether to embed inline images (enhanced mode only) |

**Return Format:**
```json
{
  "type": "text",
  "text": "Formatted email content based on mode"
}
```

**Error Responses:**
```json
{
  "type": "text", 
  "text": "Error: [specific error message]"
}
```

**Usage Examples:**

```python
# Basic mode - fastest retrieval
result = get_email_by_number_tool_unified(1, mode="basic")

# Enhanced mode - full media support
result = get_email_by_number_tool_unified(1, mode="enhanced")

# Enhanced mode without attachments
result = get_email_by_number_tool_unified(1, mode="enhanced", include_attachments=False)

# Lazy mode - intelligent selection
result = get_email_by_number_tool_unified(1, mode="lazy")
```

### `get_email_with_media_tool` (Legacy)

**Purpose**: Legacy tool maintained for backward compatibility, internally uses enhanced mode.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `email_number` | int | Yes | The 1-based position of the email in the cache |

**Return Format:** Same as enhanced mode of unified tool.

**Note**: This tool is deprecated in favor of `get_email_by_number_tool_unified` with `mode="enhanced"`.

## Data Structures

### Email Object (Basic Mode)

```json
{
  "id": "email-id",
  "subject": "Email Subject",
  "sender": "Sender Name",
  "received_time": "2024-01-01 10:00:00",
  "unread": false,
  "has_attachments": true,
  "size": 1024,
  "body": "Email body text",
  "to": "Recipient Name <email@example.com>",
  "cc": "CC Recipient <cc@example.com>",
  "attachments": [
    {
      "name": "attachment.pdf",
      "size": 2048
    }
  ]
}
```

### Email Object (Enhanced Mode)

```json
{
  "id": "email-id",
  "subject": "Email Subject",
  "sender": "Sender Name",
  "received_time": "2024-01-01 10:00:00",
  "unread": false,
  "has_attachments": true,
  "size": 1024,
  "body": "Plain text body",
  "html_body": "<html>HTML body content</html>",
  "body_format": 2,
  "to": "Recipient Name <email@example.com>",
  "cc": "CC Recipient <cc@example.com>",
  "importance": 1,
  "sensitivity": 0,
  "conversation_topic": "Conversation Topic",
  "conversation_id": "conversation-id",
  "categories": "Category1;Category2",
  "flag_status": 0,
  "attachments": [
    {
      "name": "image.jpg",
      "size": 1024,
      "mime_type": "image/jpeg",
      "is_embeddable": true,
      "content_base64": "base64encodedcontent...",
      "content_id": "image123"
    }
  ],
  "inline_images": [
    {
      "content_id": "image123",
      "data_url": "data:image/jpeg;base64,base64encodedcontent..."
    }
  ]
}
```

## Performance Considerations

### Mode Selection Guidelines

1. **Use Basic Mode When:**
   - Displaying email lists or summaries
   - Performance is critical
   - Media content is not required
   - Working with large email volumes

2. **Use Enhanced Mode When:**
   - Full email content analysis is needed
   - Attachments must be processed
   - Inline images need to be displayed
   - Detailed metadata is required

3. **Use Lazy Mode When:**
   - Optimal performance is desired
   - Email content requirements are variable
   - Cache utilization is important
   - Fallback to enhanced is acceptable

### Attachment Size Limits

- **Default Limit**: 10MB per attachment
- **Configurable**: Set via `MAX_ATTACHMENT_SIZE` environment variable
- **Large Files**: Files exceeding limit return metadata only (no base64 content)
- **Memory Impact**: Large base64 strings significantly increase memory usage

### Caching Behavior

- **Basic Mode**: Always uses cached data
- **Enhanced Mode**: Always fetches fresh data from Outlook
- **Lazy Mode**: Uses cached data when sufficient, fetches enhanced data when needed

## Error Handling

### Common Error Scenarios

1. **Invalid Email Number**
   ```json
   {"type": "text", "text": "Error: Email number must be a positive integer"}
   ```

2. **Email Not Found**
   ```json
   {"type": "text", "text": "Error: No email found at that position. Please load emails first using list_recent_emails or search_emails."}
   ```

3. **Invalid Mode**
   ```json
   {"type": "text", "text": "Error: Invalid mode 'invalid'. Valid modes: basic, enhanced, lazy"}
   ```

4. **Outlook Connection Issues**
   ```json
   {"type": "text", "text": "Error retrieving email: [specific connection error]"}
   ```

5. **Cache Issues**
   ```json
   {"type": "text", "text": "Error: Invalid cache item type"}
   ```

### Fallback Behavior

- **Enhanced Mode**: Falls back to basic mode if Outlook connection fails
- **Lazy Mode**: Falls back to enhanced mode if cached data is insufficient
- **All Modes**: Returns basic response as last resort

## Migration Guide

### From Legacy Tools

**Before (Legacy Tools):**
```python
# Basic retrieval
result = get_email_by_number_tool(1)

# Enhanced retrieval  
result = get_email_with_media_tool(1)
```

**After (Unified Tool):**
```python
# Basic retrieval (equivalent to legacy basic)
result = get_email_by_number_tool_unified(1, mode="basic")

# Enhanced retrieval (equivalent to legacy enhanced)
result = get_email_by_number_tool_unified(1, mode="enhanced")

# Custom enhanced retrieval
result = get_email_by_number_tool_unified(
    1, 
    mode="enhanced", 
    include_attachments=True,
    embed_images=True
)
```

### Backward Compatibility

The unified architecture maintains full backward compatibility:

- **Legacy Function Names**: Preserved as wrappers around unified implementation
- **Return Formats**: Consistent with previous implementations
- **Error Messages**: Maintained for existing integrations
- **Performance**: No degradation for existing usage patterns

## Best Practices

### 1. Mode Selection
```python
# For email listings - use basic mode
for email_num in range(1, 11):
    result = get_email_by_number_tool_unified(email_num, mode="basic")

# For detailed analysis - use enhanced mode
result = get_email_by_number_tool_unified(email_num, mode="enhanced")

# For optimal performance - use lazy mode
result = get_email_by_number_tool_unified(email_num, mode="lazy")
```

### 2. Error Handling
```python
def safe_get_email(email_number, mode="basic"):
    try:
        result = get_email_by_number_tool_unified(email_number, mode=mode)
        if "Error" in result.get("text", ""):
            # Handle error appropriately
            return None
        return result
    except Exception as e:
        # Handle unexpected errors
        return None
```

### 3. Performance Optimization
```python
# Batch processing with appropriate modes
def process_emails(email_numbers):
    results = []
    for num in email_numbers:
        # Use basic mode for initial screening
        basic_result = get_email_by_number_tool_unified(num, mode="basic")
        
        # Use enhanced mode only for emails that need detailed analysis
        if needs_detailed_analysis(basic_result):
            enhanced_result = get_email_by_number_tool_unified(num, mode="enhanced")
            results.append(enhanced_result)
        else:
            results.append(basic_result)
    
    return results
```

## Implementation Notes

### Thread Safety
- All tools are thread-safe for concurrent access
- Email cache is protected against race conditions
- Outlook COM objects are properly managed per thread

### Memory Management
- Base64 content is streamed when possible
- Large attachments are handled with temporary files
- Memory usage scales with attachment content size

### Security Considerations
- No sensitive data is logged
- Attachment content is processed in secure temporary directories
- Base64 encoding prevents binary data corruption

## Future Enhancements

### Planned Features
1. **Streaming Mode**: Process large attachments without loading into memory
2. **Selective Fields**: Request only specific email fields
3. **Batch Operations**: Retrieve multiple emails in single call
4. **Caching Strategy**: Configurable cache expiration and invalidation
5. **Performance Metrics**: Built-in timing and performance monitoring

### API Evolution
- New modes may be added without breaking existing functionality
- Additional parameters will have sensible defaults
- Legacy tools will be maintained for backward compatibility