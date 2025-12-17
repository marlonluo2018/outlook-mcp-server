"""Core email utilities and helpers."""

import logging
from pathlib import Path
from typing import Any, Dict

logger = logging.getLogger(__name__)


class EmailRetrievalMode:
    """Simplified email retrieval modes."""
    COMPREHENSIVE = "comprehensive"  # Always return full text content


def get_mime_type(filename: str) -> str:
    """Determine MIME type from file extension."""
    ext = Path(filename).suffix.lower()
    mime_types = {
        '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
        '.png': 'image/png', '.gif': 'image/gif',
        '.bmp': 'image/bmp', '.ico': 'image/x-icon',
        '.txt': 'text/plain', '.html': 'text/html',
        '.htm': 'text/html', '.css': 'text/css',
        '.js': 'application/javascript', '.json': 'application/json',
        '.xml': 'application/xml', '.pdf': 'application/pdf',
        '.csv': 'text/csv', '.doc': 'application/msword', 
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.xls': 'application/vnd.ms-excel', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.ppt': 'application/vnd.ms-powerpoint', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    }
    return mime_types.get(ext, 'application/octet-stream')


def format_file_size(size_bytes: int) -> str:
    """Format file size in human-readable format."""
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB"]
    i = 0
    while size_bytes >= 1024 and i < len(size_names) - 1:
        size_bytes /= 1024.0
        i += 1
    
    return f"{size_bytes:.1f} {size_names[i]}"


def _format_recipient_for_display(recipient: Any) -> str:
    """Format recipient for display."""
    if isinstance(recipient, dict):
        name = recipient.get("name", "")
        email = recipient.get("email", "")
        if name and email:
            return f"{name} <{email}>"
        elif name:
            return name
        elif email:
            return email
        else:
            return "Unknown Recipient"
    else:
        return str(recipient) if recipient else "Unknown Recipient"