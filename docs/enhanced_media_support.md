# Enhanced Email Media Support

This document describes the enhanced email media support functionality added to the Outlook MCP Server.

## Overview

The enhanced email retrieval functionality (`get_email_with_media_tool`) extends the basic email retrieval capabilities by providing comprehensive support for:

- **Attachment content extraction** (images, text files)
- **Inline image embedding** in HTML emails
- **Enhanced metadata** for all attachments
- **Content preview** for small files
- **Rich formatting** with media information

## Key Features

### 1. Attachment Content Extraction

The tool automatically extracts base64-encoded content for embeddable file types:

**Supported Image Formats:**
- JPEG/JPG (`image/jpeg`)
- PNG (`image/png`) 
- GIF (`image/gif`)
- BMP (`image/bmp`)
- ICO (`image/x-icon`)

**Supported Text Formats:**
- Plain text (`text/plain`)
- HTML (`text/html`)
- CSS (`text/css`)
- JSON (`application/json`)
- XML (`application/xml`)

**Other Formats:**
- PDF documents (metadata only)
- Office documents (metadata only)
- Binary files (metadata only)

### 2. Inline Image Processing

For HTML emails with inline images, the tool:

1. **Detects CID references** in HTML body (e.g., `<img src="cid:image001.jpg">`)
2. **Matches attachments** to content IDs
3. **Embeds images** as data URIs (e.g., `<img src="data:image/jpeg;base64,...">`)
4. **Preserves original formatting** while making images self-contained

### 3. Enhanced Data Structure

Each attachment includes comprehensive metadata:

```json
{
  "name": "document.pdf",
  "size": 1024576,
  "type": 1,
  "mime_type": "application/pdf",
  "is_embeddable": false,
  "content_base64": null,
  "content_size": 0,
  "content_id": null,
  "position": 0
}
```

### 4. Content Preview

For small text files (< 1000 bytes), the tool provides content preview:

```
Attachments (2):
1. config.json (245 bytes) [Embeddable: ✓, Content: ✓]
   Preview: {"server": "localhost", "port": 8080, "debug": true}
2. image.png (15234 bytes) [Embeddable: ✓, Content: ✓]
```

## Usage Examples

### Basic Usage

```python
# Get email with full media support
result = get_email_with_media_tool(1)
print(result["text"])
```

### Selective Media Processing

```python
# Get email without attachment content (faster)
result = get_email_with_media_tool(1, include_attachments=False)

# Get email without inline image processing
result = get_email_with_media_tool(1, embed_images=False)
```

### Complete Workflow

```python
# 1. Load emails
list_result = list_recent_emails_tool("Inbox", days=7)

# 2. Get enhanced email
enhanced_result = get_email_with_media_tool(1)

# 3. Process results
email_data = enhanced_result["text"]
# Contains: subject, sender, body, attachments, inline images, metadata
```

## Comparison with Basic Retrieval

| Feature | Basic (`get_email_by_number_tool`) | Enhanced (`get_email_with_media_tool`) |
|---------|-----------------------------------|----------------------------------------|
| Subject/Sender | ✓ | ✓ |
| Body Content | ✓ | ✓ (plain text + HTML) |
| Attachment Names | ✓ | ✓ |
| Attachment Sizes | ✓ | ✓ |
| Attachment Content | ✗ | ✓ (embeddable files) |
| Inline Images | ✗ | ✓ (embedded as data URIs) |
| MIME Types | ✗ | ✓ |
| Content Preview | ✗ | ✓ (small files) |
| Enhanced Metadata | ✗ | ✓ (importance, sensitivity, etc.) |

## Performance Considerations

### Speed Optimization
- **Use `include_attachments=False`** for faster retrieval when content is not needed
- **Use `embed_images=False`** when HTML processing is not required
- **Cache utilization**: Enhanced data is cached for subsequent requests

### Memory Management
- **Large files**: Content extraction is skipped for files > 10MB
- **Temporary files**: Automatically cleaned up after processing
- **Base64 encoding**: Increases data size by ~33% for embedded content

## Error Handling

The tool provides robust error handling:

- **Missing attachments**: Graceful fallback to metadata-only display
- **Corrupted files**: Skips content extraction but preserves metadata
- **COM errors**: Falls back to basic email data if enhanced retrieval fails
- **Memory issues**: Automatically limits content extraction for large files

## Use Cases

### 1. Email Analysis
```python
# Analyze email with rich media content
result = get_email_with_media_tool(email_number)
# Get complete context including images and attachments
```

### 2. Content Extraction
```python
# Extract text content from email attachments
result = get_email_with_media_tool(email_number)
attachments = result.get("attachments", [])
text_attachments = [a for a in attachments if a.get("mime_type") == "text/plain"]
```

### 3. HTML Email Processing
```python
# Process HTML emails with inline images
result = get_email_with_media_tool(email_number, embed_images=True)
# HTML body will have embedded images as data URIs
```

### 4. Archival and Export
```python
# Get complete email data for archival
result = get_email_with_media_tool(email_number, include_attachments=True)
# All embeddable content is included as base64
```

## Technical Implementation

### Attachment Processing Pipeline

1. **COM Interface**: Uses `win32com.client` to access Outlook attachments
2. **Content Extraction**: Saves attachments to temporary files
3. **Base64 Encoding**: Converts binary content to embeddable format
4. **MIME Detection**: Determines file types from extensions
5. **Cleanup**: Removes temporary files after processing

### Inline Image Processing

1. **HTML Parsing**: Uses regex to find CID references
2. **Content Matching**: Maps CIDs to attachment content IDs
3. **Data URI Generation**: Creates embeddable image URLs
4. **HTML Replacement**: Updates img src attributes

### Caching Strategy

- **Enhanced data**: Full email data cached after first retrieval
- **Metadata persistence**: Attachment info saved to cache file
- **Performance**: Subsequent requests served from cache
- **Memory efficiency**: Large content stored as base64 strings

## Future Enhancements

Potential improvements for future versions:

- **OCR support**: Extract text from image attachments
- **Thumbnail generation**: Create image thumbnails for large files
- **Content filtering**: Search within attachment content
- **Export formats**: Support for PDF, EML export with embedded media
- **Cloud storage**: Integration with cloud storage for large attachments