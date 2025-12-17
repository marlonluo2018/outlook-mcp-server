"""Simplified email metadata extraction."""

import logging
from typing import Dict, Any

logger = logging.getLogger(__name__)


def extract_basic_metadata(email_data: Dict[str, Any]) -> Dict[str, Any]:
    """Extract basic metadata from email data."""
    
    metadata = {
        'has_html_content': bool(email_data.get('html_body', '')),
        'has_plain_content': bool(email_data.get('body', '')),
        'text_content_length': len(email_data.get('body', '')),
        'html_content_length': len(email_data.get('html_body', '')),
        'total_recipients': 0,
        'has_attachments': email_data.get('has_attachments', False),
        'attachment_count': len(email_data.get('attachments', [])),
        'importance_level': email_data.get('importance', 1),
        'sensitivity_level': email_data.get('sensitivity', 0),
        'is_flagged': email_data.get('flag_status', 0) == 1,
        'is_unread': email_data.get('unread', False),
        'has_categories': bool(email_data.get('categories', '')),
        'conversation_id': email_data.get('conversation_id', ''),
        'conversation_topic': email_data.get('conversation_topic', ''),
    }
    
    # Count recipients
    to_recipients = email_data.get('to', '')
    cc_recipients = email_data.get('cc', '')
    
    if to_recipients:
        metadata['total_recipients'] += len(to_recipients.split(', '))
    if cc_recipients:
        metadata['total_recipients'] += len(cc_recipients.split(', '))
    
    # Basic content analysis
    text_content = email_data.get('body', '')
    if text_content:
        metadata['word_count'] = len(text_content.split())
        metadata['line_count'] = len(text_content.split('\n'))
        metadata['has_links'] = 'http://' in text_content or 'https://' in text_content
        metadata['has_email_addresses'] = '@' in text_content and '.' in text_content
    else:
        metadata['word_count'] = 0
        metadata['line_count'] = 0
        metadata['has_links'] = False
        metadata['has_email_addresses'] = False
    
    # HTML content analysis
    html_content = email_data.get('html_body', '')
    if html_content:
        metadata['html_word_count'] = len(html_content.split())
        metadata['html_has_images'] = '<img' in html_content.lower()
        metadata['html_has_tables'] = '<table' in html_content.lower()
        metadata['html_has_links'] = '<a ' in html_content.lower() and 'href=' in html_content.lower()
    else:
        metadata['html_word_count'] = 0
        metadata['html_has_images'] = False
        metadata['html_has_tables'] = False
        metadata['html_has_links'] = False
    
    # Attachment analysis
    attachments = email_data.get('attachments', [])
    if attachments:
        metadata['attachment_names'] = [attach.get('name', 'Unknown') for attach in attachments]
        metadata['total_attachment_size'] = sum(attach.get('size', 0) for attach in attachments)
        metadata['has_large_attachments'] = any(attach.get('size', 0) > 1024 * 1024 for attach in attachments)  # > 1MB
    else:
        metadata['attachment_names'] = []
        metadata['total_attachment_size'] = 0
        metadata['has_large_attachments'] = False
    
    return metadata