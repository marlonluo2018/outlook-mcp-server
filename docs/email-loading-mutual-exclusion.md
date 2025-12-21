# Email Loading Mutual Exclusion Design

## Overview

The `load_emails_by_folder_tool` implements a **mutual exclusion design** for its parameters to prevent user confusion and ensure predictable behavior.

## Problem Solved

Previously, users could specify both `days` and `max_emails` parameters together, leading to ambiguous behavior:
- Should it load emails from the specified days range?
- Should it load exactly the specified number of emails?
- What if there weren't enough emails in the days range?

## New Design

### Parameter Behavior

**Mutual Exclusion Rule**: You can only use **ONE** of these parameters at a time:

1. **Time-based loading** (`days` parameter only)
   - Loads emails within the specified date range
   - Respects the exact date range without expansion
   - Returns all emails found within the range
   - Example: `load_emails_by_folder_tool("Inbox", days=1)` → Returns emails from last 1 day only

2. **Number-based loading** (`max_emails` parameter only)
   - Loads exactly the specified number of most recent emails
   - Ignores date ranges completely
   - Example: `load_emails_by_folder_tool("Inbox", max_emails=50)` → Returns 50 most recent emails

### Error Handling

**Using both parameters together will raise an error:**
```python
load_emails_by_folder_tool("Inbox", days=7, max_emails=50)
# Error: Cannot specify both 'days' and 'max_emails' parameters. 
# Use either time-based (days) or number-based (max_emails) loading, not both.
```

### Default Behavior

**When neither parameter is specified:**
- Defaults to 7 days (original behavior)
- Example: `load_emails_by_folder_tool("Inbox")` → Equivalent to `days=7`

## Benefits

1. **Predictable Behavior**: Users know exactly what to expect
2. **No Date Range Expansion**: Time-based loading respects the specified range strictly
3. **Clear Intent**: Either "give me emails from this time period" or "give me this many emails"
4. **Reduced Confusion**: No ambiguous combined behavior

## Usage Examples

```python
# ✅ Correct usage - Time-based
load_emails_by_folder_tool("Inbox", days=1)      # Last 1 day
load_emails_by_folder_tool("Inbox", days=7)      # Last 7 days
load_emails_by_folder_tool("Inbox", days=30)     # Last 30 days

# ✅ Correct usage - Number-based  
load_emails_by_folder_tool("Inbox", max_emails=10)   # 10 most recent emails
load_emails_by_folder_tool("Inbox", max_emails=100)  # 100 most recent emails

# ✅ Default behavior
load_emails_by_folder_tool("Inbox")                  # Last 7 days (default)

# ❌ Incorrect usage - Will raise error
load_emails_by_folder_tool("Inbox", days=7, max_emails=50)
```

## Technical Implementation

The mutual exclusion is enforced at the parameter validation level in `viewing_tools.py`:

```python
# Enforce mutual exclusion: cannot use both days and max_emails together
if days is not None and max_emails is not None:
    raise ValueError("Cannot specify both 'days' and 'max_emails' parameters...")

# Set default behavior if neither parameter is specified
if days is None and max_emails is None:
    days = 7  # Default to 7 days if neither parameter is specified
```

This ensures the rule is applied consistently and provides clear error messages to users.