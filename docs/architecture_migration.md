# Email Retrieval Architecture Migration Guide

## Overview

This guide explains the migration from the dual email retrieval architecture to a unified design that consolidates functionality while maintaining backward compatibility.

## Current Architecture Issues

### Problem Statement
- **Dual Modules**: Two separate modules (`email_retrieval.py` and `email_retrieval_enhanced.py`)
- **Duplicate Tools**: Two separate MCP tools (`get_email_by_number_tool` and `get_email_with_media_tool`)
- **Code Duplication**: Similar functionality implemented differently
- **Maintenance Overhead**: Changes required in multiple places
- **Confusing API**: Users need to choose between basic and enhanced tools

### Current State
```
email_retrieval.py:
├── get_email_by_number() - Basic functionality
└── Various search functions

email_retrieval_enhanced.py:
├── get_email_with_media() - Enhanced with media support
├── get_email_with_media_tool() - MCP tool
└── Media processing functions

MCP Server:
├── get_email_by_number_tool() - Uses basic function
└── get_email_with_media_tool() - Uses enhanced function
```

## Proposed Unified Architecture

### Design Goals
1. **Single Source of Truth**: One module for all email retrieval
2. **Configurable Functionality**: Different modes for different use cases
3. **Backward Compatibility**: Existing tools continue to work
4. **Performance Optimization**: Lazy loading and caching
5. **Extensibility**: Easy to add new features

### New Architecture
```
email_retrieval_unified.py:
├── get_email_by_number_unified() - Unified retrieval with modes
├── EmailRetrievalMode enum - BASIC, ENHANCED, LAZY
└── Helper functions for media processing

email_tools_unified.py:
├── get_email_by_number_tool_unified() - Single MCP tool
├── Legacy wrappers for backward compatibility
└── Response formatting functions

MCP Server:
└── get_email_by_number_tool_unified() - Single unified tool
```

## Migration Steps

### Step 1: Update MCP Server Registration
Replace the current dual tool registration with the unified approach:

```python
# OLD: In outlook_mcp_server/__init__.py
from .backend.email_retrieval import get_email_by_number
from .backend.email_retrieval_enhanced import get_email_with_media

@mcp.tool
def get_email_by_number_tool(email_number: int) -> dict:
    # Basic implementation

@mcp.tool  
def get_email_with_media_tool(email_number: int, include_attachments: bool = True, embed_images: bool = True) -> dict:
    # Enhanced implementation

# NEW: Unified approach
from .backend.email_tools_unified import (
    get_email_by_number_tool_unified,
    get_email_tool_legacy_wrapper,
    get_email_with_media_tool_legacy_wrapper
)

# Primary unified tool
@mcp.tool
def get_email_by_number_tool_unified(email_number: int, mode: str = "basic", include_attachments: bool = True, embed_images: bool = True) -> dict:
    return get_email_by_number_tool_unified(email_number, mode, include_attachments, embed_images)

# Legacy wrappers for backward compatibility (optional)
@mcp.tool
def get_email_by_number_tool(email_number: int) -> dict:
    return get_email_tool_legacy_wrapper(email_number)

@mcp.tool
def get_email_with_media_tool(email_number: int, include_attachments: bool = True, embed_images: bool = True) -> dict:
    return get_email_with_media_tool_legacy_wrapper(email_number, include_attachments, embed_images)
```

### Step 2: Update Import Statements
Update all imports to use the unified module:

```python
# OLD
from .backend.email_retrieval import get_email_by_number
from .backend.email_retrieval_enhanced import get_email_with_media

# NEW
from .backend.email_retrieval_unified import get_email_by_number_unified, EmailRetrievalMode
```

### Step 3: Update Tests
Modify existing tests to use the unified functions:

```python
# OLD
from tests.test_enhanced_media import TestEnhancedMedia

# NEW  
from tests.test_unified_retrieval import TestUnifiedRetrieval
```

### Step 4: Update Documentation
Update all documentation to reflect the new unified approach:

```markdown
# OLD
## Email Retrieval Tools
- `get_email_by_number_tool`: Basic email retrieval
- `get_email_with_media_tool`: Enhanced email retrieval with media

# NEW
## Email Retrieval Tools
- `get_email_by_number_tool_unified`: Unified email retrieval with configurable modes
  - `mode="basic"`: Fast, backward compatible
  - `mode="enhanced"`: Full media support
  - `mode="lazy"`: Performance optimized
```

## Usage Examples

### Basic Usage (Backward Compatible)
```python
# Old way (still works)
get_email_by_number_tool(1)

# New unified way
get_email_by_number_tool_unified(1, mode="basic")
```

### Enhanced Usage (With Media)
```python
# Old way (still works)
get_email_with_media_tool(1, include_attachments=True, embed_images=True)

# New unified way
get_email_by_number_tool_unified(1, mode="enhanced", include_attachments=True, embed_images=True)
```

### Performance Optimized Usage
```python
# New lazy mode for better performance
get_email_by_number_tool_unified(1, mode="lazy")
```

## Benefits of Unified Architecture

### For Users
- **Simpler API**: One tool instead of two
- **Configurable**: Choose functionality level based on needs
- **Consistent**: Same interface for all email retrieval
- **Performance**: Lazy loading and caching optimizations

### For Developers
- **Single Module**: All email logic in one place
- **DRY Principle**: No code duplication
- **Easier Testing**: One test suite to maintain
- **Better Extensibility**: Easy to add new modes or features

### For Maintenance
- **Reduced Complexity**: One codebase to maintain
- **Consistent Updates**: Changes apply everywhere
- **Better Documentation**: Single source of truth
- **Version Control**: Simpler change tracking

## Rollback Plan

If issues arise during migration:

1. **Immediate Rollback**: Revert to previous commit
2. **Gradual Rollback**: Keep legacy tools while fixing unified version
3. **Feature Flags**: Use configuration to switch between old/new implementations

## Timeline

- **Phase 1**: Deploy unified modules alongside existing ones
- **Phase 2**: Update MCP server to use unified tools with legacy wrappers
- **Phase 3**: Update documentation and examples
- **Phase 4**: Deprecate old modules (future release)
- **Phase 5**: Remove old modules (future major release)

## Testing Strategy

1. **Unit Tests**: Comprehensive tests for unified functions
2. **Integration Tests**: End-to-end testing with Outlook
3. **Backward Compatibility**: Ensure legacy tools still work
4. **Performance Tests**: Verify lazy loading and caching
5. **Edge Cases**: Error handling and boundary conditions

## Conclusion

The unified architecture provides a cleaner, more maintainable solution while preserving backward compatibility. Users get a simpler API with more options, and developers get a more maintainable codebase.