# -*- coding: utf-8 -*-
"""Utility functions for email MS Graph channel."""
from __future__ import annotations

import html
import re
from email.utils import parseaddr


def html_to_text(html_content: str) -> str:
    """Convert HTML email content to plain text.
    
    Args:
        html_content: HTML email body
        
    Returns:
        Plain text version with HTML tags stripped
    """
    if not html_content:
        return ""
    
    # Remove script and style tags with their content
    text = re.sub(r'<script[^>]*>.*?</script>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
    
    # Replace <br> and <p> with newlines
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</p>', '\n\n', text, flags=re.IGNORECASE)
    
    # Remove all HTML tags
    text = re.sub(r'<[^>]+>', '', text)
    
    # Unescape HTML entities
    text = html.unescape(text)
    
    # Clean up whitespace
    lines = [line.strip() for line in text.split('\n')]
    text = '\n'.join(line for line in lines if line)
    
    return text.strip()


def text_to_html(text_content: str) -> str:
    """Convert plain text to HTML for email sending.
    
    Args:
        text_content: Plain text content
        
    Returns:
        HTML formatted content
    """
    if not text_content:
        return ""
    
    # Escape HTML entities
    html_content = html.escape(text_content)
    
    # Convert newlines to <br>
    html_content = html_content.replace('\n', '<br>\n')
    
    # Wrap in basic HTML structure
    return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
</head>
<body>
    <p>{html_content}</p>
</body>
</html>"""


def parse_email_address(address: str) -> tuple[str, str]:
    """Parse email address into name and email.
    
    Args:
        address: Email address string (e.g., "John Doe <john@example.com>")
        
    Returns:
        Tuple of (name, email)
    """
    name, email = parseaddr(address)
    return name or "", email or ""


def extract_conversation_id(message: dict) -> str:
    """Extract conversation ID from MS Graph message.
    
    Args:
        message: MS Graph API message object
        
    Returns:
        Conversation ID or empty string
    """
    # MS Graph provides conversationId field
    return message.get("conversationId", "") or ""


def extract_sender_email(message: dict) -> str:
    """Extract sender email address from MS Graph message.
    
    Args:
        message: MS Graph API message object
        
    Returns:
        Sender email address
    """
    sender = message.get("from", {})
    email_address = sender.get("emailAddress", {})
    return email_address.get("address", "") or ""


def extract_sender_name(message: dict) -> str:
    """Extract sender name from MS Graph message.
    
    Args:
        message: MS Graph API message object
        
    Returns:
        Sender name or email if name not available
    """
    sender = message.get("from", {})
    email_address = sender.get("emailAddress", {})
    return email_address.get("name", "") or extract_sender_email(message)


def should_process_message(
    message: dict,
    allowed_senders: list[str],
    subject_prefix: str,
) -> bool:
    """Check if message should be processed based on filters.
    
    Args:
        message: MS Graph API message object
        allowed_senders: List of allowed sender email addresses (empty = all)
        subject_prefix: Required subject prefix (empty = all)
        
    Returns:
        True if message should be processed
    """
    # Check sender filter
    if allowed_senders:
        sender_email = extract_sender_email(message)
        if sender_email not in allowed_senders:
            return False
    
    # Check subject prefix filter
    if subject_prefix:
        subject = message.get("subject", "") or ""
        if not subject.startswith(subject_prefix):
            return False
    
    return True


# TODO: Implement attachment handling functions
def extract_attachments(message: dict) -> list:
    """Extract attachments from MS Graph message.
    
    Args:
        message: MS Graph API message object
        
    Returns:
        List of attachment objects
    """
    raise NotImplementedError("Attachment extraction not implemented yet")


def download_attachment(graph_client, message_id: str, attachment_id: str, save_path: str) -> str | None:
    """Download attachment from MS Graph.
    
    Args:
        graph_client: MSGraphClient instance
        message_id: Message ID
        attachment_id: Attachment ID
        save_path: Path to save attachment
        
    Returns:
        Local file path or None if failed
    """
    raise NotImplementedError("Attachment downloading not implemented yet")
