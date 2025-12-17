# Unified Email Retrieval Architecture

## Overview

The unified email retrieval architecture consolidates the previously separate `email_retrieval.py` and `email_retrieval_enhanced.py` modules into a single, cohesive design. This architecture provides configurable email retrieval modes while maintaining full backward compatibility.

## Architecture Benefits

### Before (Dual Architecture)
- **Complexity**: Two separate modules with overlapping functionality
- **Maintenance**: Changes required in multiple places
- **Confusion**: Users had to choose between basic and enhanced tools
- **Inconsistency**: Different data structures and error handling

### After (Unified Architecture)
- **Simplicity**: Single module with configurable modes
- **Maintainability**: One codebase to maintain
- **Consistency**: Unified API and data structures
- **Flexibility**: Choose functionality level based on needs
- **Performance**: Optimized for different use cases

## Core Components

### 1. Email Retrieval Modes

The unified architecture supports three retrieval modes:

#### Basic Mode (`mode="basic"`)
- **Use Case**: Quick email previews and listings
- **Performance**: Fastest
- **Features**:
  - Basic email metadata (subject, sender, date)
  - Simple body content
  - Basic attachment information (names, sizes)
  - Minimal resource usage

#### Enhanced Mode (`mode="enhanced"`)
- **Use Case**: Complete email analysis with media content
- **Performance**: Slower but comprehensive
- **Features**:
  - All basic features
  - Base64-encoded attachment content for embeddable files
  - Inline image embedding in HTML body
  - Enhanced metadata (importance, sensitivity, categories)
  - Conversation threading information
  - Text file previews

#### Lazy Mode (`mode="lazy"`)
- **Use Case**: Balanced performance with smart caching
- **Performance**: Optimized
- **Features**:
  - Uses cached data when available
  - Falls back to enhanced mode for missing data
  - Optimal for large email datasets

### 2. Core Functions

#### `get_email_by_number_unified()`
The main unified retrieval function that supports all modes:

```python
def get_email_by_number_unified(
    email_number: int, 
    mode: str = EmailRetrievalMode.BASIC,
    include_attachments: bool = True,
    embed_images: bool = True
) -> Optional[Dict[str, Any]]:
```

**Parameters:**
- `email_number`: The email position in cache (1-based)
- `mode`: Retrieval mode ("basic", "enhanced", "lazy")
- `include_attachments`: Whether to include attachment content (enhanced mode)
- `embed_images`: Whether to embed inline images (enhanced mode)

**Returns:**
- Email data dictionary based on requested mode
- `None` if email not found or invalid parameters

#### `format_email_with_media()`
Formats enhanced email data for display:

```python
def format_email_with_media(email_data: Dict[str, Any]) -> str:
```

**Features:**
- Structured formatting with sections
- Attachment information with sizes and MIME types
- Inline image indicators
- Enhanced metadata display

### 3. MCP Tool Integration

#### Unified MCP Tool
The main MCP tool provides a single interface:

```python
def get_email_by_number_tool_unified(
    email_number: int, 
    mode: str = "basic",
    include_attachments: bool = True,
    embed_images: bool = True
) -> dict:
```

#### Legacy Compatibility
Backward compatibility is maintained through wrapper functions:
- `get_email_tool_legacy_wrapper()`: Wraps basic functionality
- `get_email_with_media_tool_legacy_wrapper()`: Wraps enhanced functionality

## Usage Examples

### Basic Email Retrieval
```python
from outlook_mcp_server.backend.email_retrieval_unified import (
    get_email_by_number_unified, 
    EmailRetrievalMode
)

# Get basic email info (fast)
email_data = get_email_by_number_unified(1, mode=EmailRetrievalMode.BASIC)
print(f"Subject: {email_data['subject']}")
print(f"From: {email_data['sender']}")
```

### Enhanced Email with Media
```python
# Get enhanced email with full media support
email_data = get_email_by_number_unified(
    1, 
    mode=EmailRetrievalMode.ENHANCED,
    include_attachments=True,
    embed_images=True
)

# Access attachment content
for attachment in email_data['attachments']:
    if attachment['content_base64']:
        print(f"Attachment {attachment['name']} has embedded content")
        
# Check for inline images
if 'inline_images' in email_data:
    print(f"Found {len(email_data['inline_images'])} inline images")
```

### MCP Tool Usage
```python
from outlook_mcp_server.backend.email_tools_unified import get_email_by_number_tool_unified

# Basic mode via MCP
result = get_email_by_number_tool_unified(1, mode="basic")
print(result['text'])

# Enhanced mode via MCP
result = get_email_by_number_tool_unified(
    1, 
    mode="enhanced",
    include_attachments=True,
    embed_images=True
)
print(result['text'])
```

### Lazy Mode for Performance
```python
# Use lazy mode for optimal performance
email_data = get_email_by_number_unified(1, mode=EmailRetrievalMode.LAZY)

# This will use cached data if available, 
# or fetch enhanced data if needed
```

## Configuration Options

### Attachment Processing
- **File Size Limit**: 10MB maximum for inline content extraction
- **Embeddable Types**: Images (JPEG, PNG, GIF, BMP, ICO) and text files
- **MIME Type Detection**: Automatic based on file extension
- **Content Preview**: Available for text files ≤ 1000 characters

### Error Handling
- **Graceful Fallbacks**: Enhanced mode falls back to basic on errors
- **Session Management**: Automatic Outlook session handling
- **Validation**: Parameter validation with clear error messages
- **Logging**: Comprehensive logging for debugging

## Performance Characteristics

| Mode | Speed | Memory | Network Calls | Best For |
|------|-------|--------|---------------|----------|
| Basic | Fastest | Low | Minimal | Quick previews, large lists |
| Lazy | Fast | Medium | Minimal (cached) | General usage, mixed workloads |
| Enhanced | Slower | High | Full | Detailed analysis, media extraction |

## Migration Guide

### For Existing Users
The unified architecture maintains full backward compatibility:

```python
# Old way (still works)
from outlook_mcp_server.backend.email_retrieval import get_email_by_number

# New unified way (recommended)
from outlook_mcp_server.backend.email_retrieval_unified import (
    get_email_by_number_unified, 
    EmailRetrievalMode
)

# Both produce the same results for basic functionality
```

### For New Users
Start with the unified architecture:

```python
from outlook_mcp_server.backend.email_retrieval_unified import (
    get_email_by_number_unified, 
    EmailRetrievalMode
)

# Choose mode based on your needs
email_data = get_email_by_number_unified(1, mode=EmailRetrievalMode.BASIC)
```

## Error Handling

### Common Errors and Solutions

#### Invalid Email Number
```python
# Invalid email number
result = get_email_by_number_unified(-1, mode="basic")
# Returns: None

# Out of range
result = get_email_by_number_unified(999, mode="basic") 
# Returns: None
```

#### Invalid Mode
```python
# Invalid mode
result = get_email_by_number_unified(1, mode="invalid")
# Returns: None with error message
```

#### Session Errors
```python
# Outlook session error in enhanced mode
result = get_email_by_number_unified(1, mode="enhanced")
# Falls back to basic mode gracefully
```

## Best Practices

### 1. Choose the Right Mode
- Use **Basic** for quick previews and listings
- Use **Enhanced** for detailed analysis and media extraction
- Use **Lazy** for general usage with performance optimization

### 2. Handle Errors Gracefully
```python
try:
    email_data = get_email_by_number_unified(email_number, mode="enhanced")
    if email_data is None:
        # Handle not found case
        print("Email not found")
    else:
        # Process email data
        process_email(email_data)
except Exception as e:
    # Handle unexpected errors
    logger.error(f"Error retrieving email: {e}")
```

### 3. Use Appropriate Configuration
```python
# For large email sets, use basic or lazy mode
for email_num in range(1, 100):
    email_data = get_email_by_number_unified(email_num, mode="lazy")
    
# For detailed analysis of specific emails
email_data = get_email_by_number_unified(
    specific_email_num, 
    mode="enhanced",
    include_attachments=True,
    embed_images=True
)
```

## Future Enhancements

### Planned Features
1. **Additional Modes**: Support for specialized retrieval modes
2. **Streaming Support**: Large attachment streaming
3. **Caching Improvements**: More sophisticated caching strategies
4. **Performance Metrics**: Built-in performance monitoring
5. **Plugin Architecture**: Extensible mode system

### Extension Points
- Custom retrieval modes
- Specialized formatters
- Additional media processors
- Enhanced caching strategies

## Conclusion

The unified email retrieval architecture provides a robust, maintainable solution for Outlook email retrieval. It offers the flexibility to choose the right level of functionality for each use case while maintaining consistency and performance. The architecture is designed to be extensible and maintainable, providing a solid foundation for future enhancements.