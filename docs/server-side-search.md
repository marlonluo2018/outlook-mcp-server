# Server-Side Search Implementation Guide

This document details the server-side search implementation for the Outlook MCP Server, focusing on Win32COM API integration and search optimization techniques.

## Overview

The server-side search implementation provides efficient email searching capabilities by leveraging Outlook's built-in search functionality rather than client-side filtering. This approach significantly improves performance and reduces memory usage.

## Search Methods

### 1. AdvancedSearch Method (Primary)

The AdvancedSearch method is the primary server-side search mechanism that provides the most efficient searching capabilities.

```python
def server_side_search(namespace, folder, search_criteria, max_results=100):
    """Perform server-side search using AdvancedSearch method."""
    try:
        # Build search scope (folder path)
        scope = f"'{folder.FolderPath}'"
        
        # Execute AdvancedSearch
        search_results = namespace.Application.AdvancedSearch(
            Scope=scope,
            Filter=search_criteria,
            SearchSubFolders=True,
            Tag="MCPSearch"
        )
        
        # Wait for search completion (with timeout)
        timeout = time.time() + 30  # 30 second timeout
        while not search_results.IsComplete:
            if time.time() > timeout:
                raise TimeoutError("Search timeout exceeded")
            time.sleep(0.1)
        
        return search_results.Results
        
    except AttributeError as e:
        logger.error(f"AdvancedSearch not available: {e}")
        # Fallback to Restrict method
        return restrict_search(folder, search_criteria)
```

**Key Features:**
- Executes search on the Exchange server (if applicable)
- Supports complex SQL-like queries
- Includes subfolder search capability
- Provides completion status tracking
- Implements timeout protection

### 2. Restrict Method (Fallback)

The Restrict method serves as a reliable fallback when AdvancedSearch encounters issues or is not available.

```python
def restrict_search(folder, filter_criteria):
    """Perform search using Restrict method as fallback."""
    try:
        # Apply filter to folder items
        filtered_items = folder.Items.Restrict(filter_criteria)
        return filtered_items
        
    except Exception as e:
        logger.error(f"Restrict search failed: {e}")
        raise
```

**Key Features:**
- Filters items within the local Outlook application
- Supports SQL-like filtering syntax
- More reliable but potentially slower than AdvancedSearch
- Works with all Outlook configurations

## Search Criteria Formatting

### SQL-Based Search Criteria

Outlook search uses SQL-like syntax for maximum flexibility and performance.

```python
def build_search_criteria(search_terms, days=7, match_all=True):
    """Build properly formatted search criteria for Outlook."""
    
    # Date filtering
    date_limit = datetime.now() - timedelta(days=days)
    date_str = date_limit.strftime("%Y-%m-%d")
    
    # Subject search terms
    subject_conditions = []
    for term in search_terms:
        # Use LIKE for partial matching
        escaped_term = term.replace("'", "''")  # Escape single quotes
        condition = f"urn:schemas:httpmail:subject LIKE '%{escaped_term}%'"
        subject_conditions.append(condition)
    
    # Combine conditions
    if match_all:
        subject_criteria = " AND ".join(subject_conditions)
    else:
        subject_criteria = " OR ".join(subject_conditions)
    
    # Full criteria with date filtering
    criteria = f"@SQL={subject_criteria} AND urn:schemas:httpmail:datereceived >= '{date_str}'"
    
    return criteria
```

### Search Schema Reference

Common Outlook search schemas:

| Schema | Description | Example |
|--------|-------------|---------|
| `urn:schemas:httpmail:subject` | Email subject | `subject LIKE '%approval%'` |
| `urn:schemas:httpmail:from` | Sender email | `from LIKE '%@company.com%'` |
| `urn:schemas:httpmail:datereceived` | Received date | `datereceived >= '2025-12-01'` |
| `urn:schemas:httpmail:hasattachment` | Has attachments | `hasattachment = 1` |
| `urn:schemas:httpmail:textdescription` | Body content | `textdescription LIKE '%meeting%'` |

### Embedded Images and Attachments

The system now provides enhanced attachment tracking with separate embedded image counting:

```python
def extract_search_results_with_attachments(search_results):
    """Extract search results with detailed attachment information."""
    
    results = []
    for item in search_results:
        # Basic email information
        email_data = {
            'subject': item.Subject,
            'sender': item.SenderName,
            'received_time': item.ReceivedTime,
            'entry_id': item.EntryID
        }
        
        # Enhanced attachment information
        attachments_count = 0
        embedded_images_count = 0
        
        if hasattr(item, 'Attachments') and item.Attachments:
            attachments_count = item.Attachments.Count
            
            # Separate embedded images from regular attachments
            for i in range(1, attachments_count + 1):
                try:
                    attachment = item.Attachments.Item(i)
                    if hasattr(attachment, 'Type') and attachment.Type == 1:  # olEmbeddeditem
                        embedded_images_count += 1
                except Exception:
                    continue
        
        email_data['attachments_count'] = attachments_count
        email_data['embedded_images_count'] = embedded_images_count
        email_data['regular_attachments_count'] = attachments_count - embedded_images_count
        
        results.append(email_data)
    
    return results
```

**Enhanced Display Format:**
- `Embedded Images: 2` (shows count or "None")
- `Attachments: 3` (regular attachments, shows count or "None")
- Clear separation for better email information clarity

### Complex Search Examples

```python
# Search for approval emails in the last 3 days
criteria = "@SQL=urn:schemas:httpmail:subject LIKE '%approval%' AND urn:schemas:httpmail:datereceived >= '2025-12-15'"

# Search for emails from specific sender with attachments
criteria = "@SQL=urn:schemas:httpmail:from LIKE '%manager@company.com%' AND urn:schemas:httpmail:hasattachment = 1"

# Search for multiple terms (OR logic)
criteria = "@SQL=(urn:schemas:httpmail:subject LIKE '%urgent%' OR urn:schemas:httpmail:subject LIKE '%important%')"
```

## Error Handling

### Common Search Errors and Solutions

```python
def handle_search_errors(func):
    """Decorator for comprehensive search error handling."""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
            
        except AttributeError as e:
            logger.error(f"COM AttributeError: {e}")
            # Fallback to Restrict method
            return fallback_to_restrict_search(*args, **kwargs)
            
        except pythoncom.com_error as e:
            logger.error(f"COM error: {e}")
            # Reinitialize COM and retry
            pythoncom.CoInitialize()
            return func(*args, **kwargs)
            
        except TimeoutError as e:
            logger.error(f"Search timeout: {e}")
            # Return partial results or empty set
            return []
            
        except Exception as e:
            logger.error(f"Unexpected search error: {e}")
            raise
    
    return wrapper
```

## Performance Considerations

### Search Optimization Tips

1. **Use Specific Schemas**: Target specific fields rather than broad searches
2. **Limit Date Ranges**: Always include date filters to reduce search scope
3. **Avoid Complex OR Conditions**: Use AND logic when possible for better performance
4. **Escape Special Characters**: Properly escape quotes and special characters
5. **Use Appropriate Methods**: Choose AdvancedSearch for server-side, Restrict for local

### Search Performance Comparison

| Method | Speed | Reliability | Server Load | Use Case |
|--------|-------|-------------|-------------|----------|
| AdvancedSearch | Fastest | Medium | Low | Large folders, Exchange server |
| Restrict | Medium | High | None | Local folders, reliable fallback |
| Client-side filtering | Slowest | High | None | Small datasets, complex logic |

## Integration Example

Complete integration example showing server-side search in action:

```python
def search_emails_server_side(folder_name, search_terms, days=7):
    """Complete server-side search implementation."""
    
    # Initialize Outlook
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Get target folder
    folder = get_folder_by_path(namespace, folder_name)
    
    # Build search criteria
    criteria = build_search_criteria(search_terms, days)
    
    # Execute server-side search
    try:
        results = server_side_search(namespace, folder, criteria)
        
        # Process results
        emails = []
        for item in results:
            email_data = {
                'subject': item.Subject,
                'sender': item.SenderName,
                'received_time': item.ReceivedTime,
                'entry_id': item.EntryID
            }
            emails.append(email_data)
        
        return emails
        
    except Exception as e:
        logger.error(f"Server-side search failed: {e}")
        # Fallback to restrict search
        return search_with_restrict(folder, criteria)
```

## Best Practices

### 1. Always Use Server-Side Search First
- Implement server-side search as the primary method
- Use client-side filtering only as a last resort
- Monitor search performance and adjust methods accordingly

### 2. Implement Proper Fallbacks
- Always have Restrict method as fallback
- Log when fallbacks are used for monitoring
- Test fallback scenarios thoroughly

### 3. Optimize Search Criteria
- Use specific schemas instead of broad searches
- Include date filters to limit scope
- Test criteria performance with different folder sizes

### 4. Handle Timeouts Gracefully
- Implement reasonable timeout limits (30 seconds)
- Return partial results when possible
- Log timeout events for analysis

## Troubleshooting

### Common Issues

1. **"AdvancedSearch not available"**
   - Solution: Fallback to Restrict method
   - Cause: Outlook configuration or permissions

2. **"Search timeout exceeded"**
   - Solution: Reduce search scope or date range
   - Cause: Large folder or complex criteria

3. **"Invalid search criteria"**
   - Solution: Validate SQL syntax and schema names
   - Cause: Malformed search criteria

4. **"No results found"**
   - Solution: Check search terms and date ranges
   - Cause: Criteria too restrictive or terms not found

This server-side search implementation provides a robust, efficient foundation for email searching that scales well with large email volumes while maintaining reliability through comprehensive error handling and fallback mechanisms.