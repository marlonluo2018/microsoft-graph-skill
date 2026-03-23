#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Microsoft Graph Email Operations Module

Provides email operations including read, send, reply, and forward.
Enforces company policy of maximum 500 recipients per email.

Usage:
    python email_operations.py list [--folder <folder>] [--limit <n>]
    python email_operations.py get <message_id>
    python email_operations.py send --to <emails> [--cc <emails>] [--bcc <emails>] --subject <subject> --body <body>
    python email_operations.py reply <message_id> [--all] --body <body>
    python email_operations.py forward <message_id> --to <emails> [--cc <emails>] --comment <comment>
"""

import os
import sys
import json
import argparse
from pathlib import Path
from typing import List, Optional, Dict, Any
from datetime import datetime
from zoneinfo import ZoneInfo

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Add parent directory to path for config import
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import configuration and auth
from config import (
    GRAPH_API_BASE, MAX_RECIPIENTS_PER_EMAIL,
    MAX_MESSAGE_DISPLAY_LENGTH, MAX_BODY_DISPLAY_LENGTH
)
from auth import get_access_token

# Try to import requests
try:
    import requests
except ImportError:
    print("Error: requests package not found.")
    print("Install with: pip install requests")
    sys.exit(1)

# Constants
MAX_RECIPIENTS = MAX_RECIPIENTS_PER_EMAIL


def get_headers(token: str) -> Dict[str, str]:
    """Get authorization headers for API requests."""
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }


def validate_recipients(to: List[str], cc: List[str] = None, bcc: List[str] = None) -> bool:
    """
    Validate that total recipients don't exceed company limit.
    
    Args:
        to: List of To recipients
        cc: List of CC recipients
        bcc: List of BCC recipients
    
    Returns:
        bool: True if valid, raises ValueError otherwise
    """
    total = len(to) + len(cc or []) + len(bcc or [])
    
    if total > MAX_RECIPIENTS:
        raise ValueError(
            f"Total recipients ({total}) exceeds company limit of {MAX_RECIPIENTS} per email. "
            f"Please split into multiple emails."
        )
    
    if total == 0:
        raise ValueError("At least one recipient is required.")
    
    return True


def format_email_address(email: str, name: str = None) -> Dict[str, str]:
    """Format email address for Graph API."""
    return {
        "emailAddress": {
            "address": email,
            "name": name or email
        }
    }


def parse_email_list(emails: str) -> List[str]:
    """Parse comma or semicolon separated email list."""
    if not emails:
        return []
    
    # Support both comma and semicolon separators
    emails = emails.replace(';', ',')
    return [e.strip() for e in emails.split(',') if e.strip()]


# =============================================================================
# LIST MESSAGES
# =============================================================================

def list_messages(
    folder: str = "inbox",
    limit: int = 25,
    filter_query: str = None,
    order_by: str = "receivedDateTime desc",
    include_preview: bool = False,
    inference_classification: str = None,
    from_sender: str = None,
    to_recipient: str = None,
    subject: str = None,
    body: str = None,
    token: str = None
) -> List[Dict[str, Any]]:
    """
    List/search messages from a folder with optional search criteria.
    
    Args:
        folder: Folder name (inbox, sentitems, drafts, etc.) or folder ID or 'all' for all folders
        limit: Maximum number of messages to return
        filter_query: OData filter query
        order_by: Sort order
        include_preview: Include bodyPreview field
        inference_classification: Filter by classification ("focused" or "other")
        from_sender: Search by sender name or email
        to_recipient: Search by recipient name or email
        subject: Search by subject text
        body: Search by body text
        token: Access token (will obtain if not provided)
    
    Returns:
        List of message objects
    """
    if token is None:
        token = get_access_token()
    
    # Auto-detect Chinese and English patterns and adjust parameters
    if from_sender:
        # English patterns
        if any(keyword in from_sender.lower() for keyword in ["sent to", "send to", "sent"]):
            # "sent to" means searching sent folder - switch to sent folder and use to_recipient
            folder = "sentitems"
            # Remove the pattern keywords (case-insensitive)
            cleaned = from_sender
            for keyword in ["sent to", "send to", "sent"]:
                cleaned = cleaned.replace(keyword, "").replace(keyword.title(), "").replace(keyword.upper(), "")
            to_recipient = cleaned.strip()
            from_sender = None
        elif any(keyword in from_sender.lower() for keyword in ["received from", "from", "received"]):
            # "received from" means searching inbox - use inbox and clean keyword
            folder = "inbox"
            cleaned = from_sender
            for keyword in ["received from", "from", "received"]:
                cleaned = cleaned.replace(keyword, "").replace(keyword.title(), "").replace(keyword.upper(), "")
            from_sender = cleaned.strip()
        # Chinese patterns
        elif any(keyword in from_sender for keyword in ["发给", "发送给", "发送到"]):
            # "发给" means "sent to" - switch to sent folder and use to_recipient
            folder = "sentitems"
            to_recipient = from_sender.replace("发给", "").replace("发送给", "").replace("发送到", "").strip()
            from_sender = None
        elif any(keyword in from_sender for keyword in ["收到", "来自", "从"]):
            # "收到/来自" means "received from" - use inbox and clean keyword
            folder = "inbox"
            from_sender = from_sender.replace("收到", "").replace("来自", "").replace("从", "").strip()
    
    # Check if search criteria provided
    has_search_criteria = any([from_sender, to_recipient, subject, body])
    
    # Map common folder names to well-known folder IDs
    folder_map = {
        "inbox": "inbox",
        "sent": "sentitems",
        "sentitems": "sentitems",
        "drafts": "drafts",
        "deleted": "deleteditems",
        "deleteditems": "deleteditems",
        "junk": "junkemail",
        "junkemail": "junkemail",
        "outbox": "outbox"
    }
    
    # Build URL based on folder
    if folder.lower() == "all":
        url = f"{GRAPH_API_BASE}/me/messages"
    else:
        folder_id = folder_map.get(folder.lower(), folder)
        url = f"{GRAPH_API_BASE}/me/mailFolders/{folder_id}/messages"
    
    # Build select fields
    select_fields = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,inferenceClassification"
    if include_preview:
        select_fields += ",bodyPreview"
    
    # Build parameters
    params = {
        "$select": select_fields
    }
    
    # If search criteria provided, use $search
    if has_search_criteria:
        search_keywords = []
        if from_sender:
            search_keywords.append(from_sender)
        if to_recipient:
            search_keywords.append(to_recipient)
        if subject:
            search_keywords.append(subject)
        if body:
            search_keywords.append(body)
        
        search_query = " ".join(search_keywords)
        params["$search"] = f'"{search_query}"'
        params["$top"] = min(limit * 3, 100)  # Get more for client-side filtering
    else:
        params["$top"] = limit
        params["$orderby"] = order_by
    
    # Build filter query
    filters = []
    if filter_query:
        filters.append(filter_query)
    
    if inference_classification:
        classification = inference_classification.lower()
        if classification in ["focused", "other"]:
            filters.append(f"inferenceClassification eq '{classification}'")
    
    if filters:
        params["$filter"] = " and ".join(filters)
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to list messages: {response.status_code} - {response.text}")
    
    data = response.json()
    messages = data.get("value", [])
    
    # Apply client-side filtering for precise search matching
    if has_search_criteria:
        if from_sender:
            from_lower = from_sender.lower()
            messages = [
                m for m in messages
                if from_lower in m.get('from', {}).get('emailAddress', {}).get('name', '').lower()
                or from_lower in m.get('from', {}).get('emailAddress', {}).get('address', '').lower()
            ]
        
        if to_recipient:
            to_lower = to_recipient.lower()
            messages = [
                m for m in messages
                if any(
                    to_lower in r.get('emailAddress', {}).get('name', '').lower()
                    or to_lower in r.get('emailAddress', {}).get('address', '').lower()
                    for r in m.get('toRecipients', [])
                )
            ]
        
        if subject:
            subject_lower = subject.lower()
            messages = [
                m for m in messages
                if subject_lower in m.get('subject', '').lower()
            ]
        
        # Limit results after filtering
        messages = messages[:limit]
    
    return messages




# =============================================================================
# GET MESSAGE
# =============================================================================

def get_message(message_id: str, token: str = None) -> Dict[str, Any]:
    """
    Get a specific message by ID.
    
    Args:
        message_id: Message ID
        token: Access token
    
    Returns:
        Message object
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}"
    
    response = requests.get(url, headers=get_headers(token))
    
    if response.status_code != 200:
        raise Exception(f"Failed to get message: {response.status_code} - {response.text}")
    
    return response.json()


# =============================================================================
# SEND EMAIL
# =============================================================================

def send_email(
    to: List[str],
    subject: str,
    body: str,
    cc: List[str] = None,
    bcc: List[str] = None,
    body_type: str = "html",
    attachments: List[Dict] = None,
    save_to_sent: bool = True,
    token: str = None
) -> bool:
    """
    Send an email.
    
    Args:
        to: List of To recipient emails
        subject: Email subject
        body: Email body content
        cc: List of CC recipient emails
        bcc: List of BCC recipient emails
        body_type: "html" or "text"
        attachments: List of attachment objects
        save_to_sent: Whether to save to Sent Items
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    # Validate recipients
    validate_recipients(to, cc, bcc)
    
    # Build message payload
    message = {
        "subject": subject,
        "body": {
            "contentType": body_type,
            "content": body
        },
        "toRecipients": [format_email_address(e) for e in to],
        "ccRecipients": [format_email_address(e) for e in (cc or [])],
        "bccRecipients": [format_email_address(e) for e in (bcc or [])]
    }
    
    # Add attachments if provided
    if attachments:
        message["attachments"] = attachments
    
    # Build request payload
    payload = {
        "message": message,
        "saveToSentItems": save_to_sent
    }
    
    url = f"{GRAPH_API_BASE}/me/sendMail"
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code != 202:
        raise Exception(f"Failed to send email: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# HELPER: FORMAT EMAIL AS HTML
# =============================================================================

def format_email_as_html(message: Dict[str, Any]) -> str:
    """
    Format an email message as HTML for inclusion in reply/forward.
    
    Args:
        message: Message object from Graph API
    
    Returns:
        HTML string with formatted email
    """
    from_addr = message.get('from', {}).get('emailAddress', {})
    from_name = from_addr.get('name', '')
    from_email = from_addr.get('address', '')
    
    to_recipients = message.get('toRecipients', [])
    to_list = ', '.join([f"{r.get('emailAddress', {}).get('name', '')} <{r.get('emailAddress', {}).get('address', '')}>"
                         for r in to_recipients])
    
    cc_recipients = message.get('ccRecipients', [])
    cc_list = ', '.join([f"{r.get('emailAddress', {}).get('name', '')} <{r.get('emailAddress', {}).get('address', '')}>"
                         for r in cc_recipients]) if cc_recipients else None
    
    subject = message.get('subject', '(No Subject)')
    date = message.get('receivedDateTime', '')
    
    # Get body content
    body_obj = message.get('body', {})
    body_content = body_obj.get('content', '')
    body_type = body_obj.get('contentType', 'text')
    
    # If body is text, convert to HTML
    if body_type == 'text':
        body_content = body_content.replace('\n', '<br>')
    
    # Build HTML
    html = f"""<hr>
<p><strong>From:</strong> {from_name} <{from_email}><br>
<strong>Sent:</strong> {date}<br>
<strong>To:</strong> {to_list}<br>"""
    
    if cc_list:
        html += f"<strong>Cc:</strong> {cc_list}<br>"
    
    html += f"""<strong>Subject:</strong> {subject}</p>
{body_content}"""
    
    return html


# =============================================================================
# REPLY TO EMAIL
# =============================================================================

def reply_email(
    message_id: str,
    body: str,
    reply_all: bool = False,
    body_type: str = "html",
    include_history: bool = True,
    token: str = None
) -> bool:
    """
    Reply to an email with conversation history included in body.
    
    Args:
        message_id: ID of message to reply to
        body: Reply body content
        reply_all: If True, reply to all recipients; otherwise reply to sender only
        body_type: "html" or "text"
        include_history: If True, include original message in reply body (default: True)
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    # Get original message
    original_msg = get_message(message_id, token)
    
    # Check if replying to self-sent email
    from_addr = original_msg.get('from', {}).get('emailAddress', {})
    my_email = get_my_email(token)
    
    if from_addr.get('address', '').lower() == my_email.lower():
        print("⚠️  WARNING: You are replying to your own sent email.")
        print("💡 TIP: Use 'forward' command instead to forward your sent email to others.")
        print("   Example: forward <message_id> --to \"recipient@example.com\"")
        print()
    
    # Determine recipients
    from_addr = original_msg.get('from', {}).get('emailAddress', {})
    to_recipients = [from_addr.get('address')]
    cc_recipients = []
    
    if reply_all:
        # Add all original To recipients (except current user)
        my_email = get_my_email(token)
        for recipient in original_msg.get('toRecipients', []):
            email = recipient.get('emailAddress', {}).get('address', '')
            if email and email.lower() != my_email.lower() and email not in to_recipients:
                to_recipients.append(email)
        
        # Add all original CC recipients
        for recipient in original_msg.get('ccRecipients', []):
            email = recipient.get('emailAddress', {}).get('address', '')
            if email and email.lower() != my_email.lower():
                cc_recipients.append(email)
    
    # Build email body with history
    if include_history and body_type == "html":
        # Convert plain text body to HTML if needed
        if not body.strip().startswith('<'):
            body = body.replace('\n', '<br>\n')
        
        full_body = body + "\n\n" + format_email_as_html(original_msg)
    else:
        full_body = body
    
    # Get subject with RE: prefix
    subject = original_msg.get('subject', '')
    if not subject.upper().startswith('RE:'):
        subject = f"RE: {subject}"
    
    # Send email using send_email function
    return send_email(
        to=to_recipients,
        subject=subject,
        body=full_body,
        cc=cc_recipients if cc_recipients else None,
        body_type=body_type,
        token=token
    )


def get_my_email(token: str = None) -> str:
    """Get current user's email address."""
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me"
    response = requests.get(url, headers=get_headers(token))
    
    if response.status_code != 200:
        return ""
    
    data = response.json()
    return data.get('mail', '') or data.get('userPrincipalName', '')


# =============================================================================
# FORWARD EMAIL
# =============================================================================

def forward_email(
    message_id: str,
    to: List[str],
    cc: List[str] = None,
    bcc: List[str] = None,
    comment: str = "",
    body_type: str = "html",
    include_history: bool = True,
    token: str = None
) -> bool:
    """
    Forward an email with original message and attachments included.
    Uses Microsoft Graph's native forward API to preserve attachments.
    
    Args:
        message_id: ID of message to forward
        to: List of To recipient emails
        cc: List of CC recipient emails
        bcc: List of BCC recipient emails
        comment: Optional comment to add at the top
        body_type: "html" or "text"
        include_history: If True, include original message in body (default: True)
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    # Validate recipients
    validate_recipients(to, cc, bcc)
    
    # Use Microsoft Graph's native forward endpoint to preserve attachments
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}/forward"
    
    # Build recipient lists
    to_recipients = [{"emailAddress": {"address": email}} for email in to]
    cc_recipients = [{"emailAddress": {"address": email}} for email in (cc or [])]
    bcc_recipients = [{"emailAddress": {"address": email}} for email in (bcc or [])]
    
    # Convert plain text comment to HTML if needed
    # Graph API forward expects HTML in comment field
    if comment:
        if body_type == "html" and not comment.strip().startswith('<'):
            # Convert plain text to HTML with proper line breaks
            comment = comment.replace('\n', '<br>\n')
            # Wrap in HTML body tags for proper rendering
            comment = f"<html><body>{comment}</body></html>"
        elif body_type == "text":
            # For text, just use as-is
            pass
    
    # Build payload for Graph API forward
    payload = {
        "toRecipients": to_recipients,
        "comment": comment or ""
    }
    
    if cc_recipients:
        payload["ccRecipients"] = cc_recipients
    if bcc_recipients:
        payload["bccRecipients"] = bcc_recipients
    
    # Send forward request
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code not in [200, 201, 202]:
        raise Exception(f"Failed to forward email: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# MARK AS READ/UNREAD
# =============================================================================

def mark_as_read(message_id: str, token: str = None) -> bool:
    """Mark a message as read."""
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}"
    
    payload = {"isRead": True}
    
    response = requests.patch(url, headers=get_headers(token), json=payload)
    
    if response.status_code != 200:
        raise Exception(f"Failed to mark as read: {response.status_code} - {response.text}")
    
    return True


def mark_as_unread(message_id: str, token: str = None) -> bool:
    """Mark a message as unread."""
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}"
    
    payload = {"isRead": False}
    
    response = requests.patch(url, headers=get_headers(token), json=payload)
    
    if response.status_code != 200:
        raise Exception(f"Failed to mark as unread: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# GET EMAIL THREAD (CONVERSATION)
# =============================================================================

def get_message_thread(message_id: str, token: str = None) -> List[Dict[str, Any]]:
    """
    Get all messages in the same conversation thread.
    
    Args:
        message_id: ID of any message in the conversation
        token: Access token
    
    Returns:
        List of messages in the conversation, ordered by date
    """
    if token is None:
        token = get_access_token()
    
    # First get the message with conversation ID and subject
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}?$select=conversationId,subject"
    response = requests.get(url, headers=get_headers(token))
    
    if response.status_code != 200:
        raise Exception(f"Failed to get message: {response.status_code} - {response.text}")
    
    message = response.json()
    conversation_id = message.get('conversationId')
    subject = message.get('subject', '')
    
    if not conversation_id:
        raise Exception("No conversation ID found for this message")
    
    # Use conversation ID directly to get messages
    # Graph API supports /me/messages/{id} with $expand for conversation
    url = f"{GRAPH_API_BASE}/me/messages"
    
    # Try using $search with conversation index
    # Extract keywords from subject for searching
    search_terms = subject.replace('RE:', '').replace('FW:', '').replace('Fw:', '').strip()
    if len(search_terms) > 50:
        search_terms = search_terms[:50]
    
    params = {
        "$search": f'"{search_terms}"',
        "$top": 50,
        "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,conversationId,internetMessageId"
    }
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to get thread: {response.status_code} - {response.text}")
    
    # Filter to only messages in the same conversation
    all_messages = response.json().get("value", [])
    thread_messages = [m for m in all_messages if m.get('conversationId') == conversation_id]
    
    # Sort by receivedDateTime ascending (oldest first)
    thread_messages.sort(key=lambda x: x.get('receivedDateTime', ''))
    
    return thread_messages


def display_thread(messages: List[Dict]):
    """Display a conversation thread in chronological order."""
    import re
    import html as html_module
    
    if not messages:
        print("No messages in thread.")
        return
    
    print(f"\n{'='*80}")
    print(f"CONVERSATION THREAD ({len(messages)} messages)")
    print(f"Subject: {messages[0].get('subject', '(No Subject)')}")
    print(f"{'='*80}")
    
    for i, msg in enumerate(messages, 1):
        print(f"\n--- Message {i}/{len(messages)} ---")
        print(f"From: {msg.get('from', {}).get('emailAddress', {}).get('name', 'Unknown')}")
        print(f"      <{msg.get('from', {}).get('emailAddress', {}).get('address', '')}>")
        print(f"Date: {msg.get('receivedDateTime', '')}")
        
        to_list = [r.get('emailAddress', {}).get('name', '') for r in msg.get('toRecipients', [])]
        cc_list = [r.get('emailAddress', {}).get('name', '') for r in msg.get('ccRecipients', [])]
        if to_list:
            print(f"To: {', '.join(to_list)}")
        if cc_list:
            print(f"Cc: {', '.join(cc_list)}")
        
        print(f"\n")
        
        # Get body content
        body = msg.get('body', {})
        content = body.get('content', '')
        content_type = body.get('contentType', 'text')
        
        # If HTML, extract plain text
        if content_type == 'html' and content:
            content = re.sub(r'<style[^>]*>.*?</style>', '', content, flags=re.DOTALL)
            content = re.sub(r'<[^>]+>', ' ', content)
            content = html_module.unescape(content)
            content = re.sub(r'\s+', ' ', content).strip()
            # Limit length for thread view
            if len(content) > MAX_BODY_DISPLAY_LENGTH:
                content = content[:1000] + '...'
        
        # Handle encoding issues
        try:
            print(content)
        except UnicodeEncodeError:
            print(content.encode('ascii', 'replace').decode('ascii'))
    
    print(f"\n{'='*80}")
    print(f"End of thread ({len(messages)} messages)")


# =============================================================================
# DELETE EMAIL
# =============================================================================

def delete_email(message_id: str, token: str = None) -> bool:
    """Delete an email."""
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}"
    
    response = requests.delete(url, headers=get_headers(token))
    
    if response.status_code != 204:
        raise Exception(f"Failed to delete email: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# DISPLAY HELPERS
# =============================================================================

def display_message_list(messages: List[Dict], show_preview: bool = True):
    """Display a list of messages in a readable format."""
    if not show_preview:
        # Compact table format without preview
        print(f"\n{'='*100}")
        print(f"{'Date':<20} {'From':<25} {'Subject':<30} {'ID':<15}")
        print(f"{'='*100}")

        for msg in messages:
            received = msg.get('receivedDateTime', '')
            if received:
                dt = datetime.fromisoformat(received.replace('Z', '+00:00'))
                # Convert to Asia/Shanghai timezone
                dt_local = dt.astimezone(ZoneInfo('Asia/Shanghai'))
                received = dt_local.strftime('%Y-%m-%d %H:%M')

            from_addr = msg.get('from', {}).get('emailAddress', {})
            sender = from_addr.get('name', from_addr.get('address', 'Unknown'))

            subject = msg.get('subject', '(No Subject)')[:30]
            read_status = '' if msg.get('isRead', True) else '[UNREAD]'
            message_id = msg.get('id', '')[:15]

            print(f"{received:<20} {sender:<25} {subject}{read_status:<{30 - len(read_status)}} {message_id:<15}")

        print(f"{'='*100}")
        print(f"Total: {len(messages)} messages")
        print(f"💡 Tip: Use 'get <ID>' to view full email content")
    else:
        # Detailed format with preview
        print(f"\n{'='*80}")
        for i, msg in enumerate(messages, 1):
            received = msg.get('receivedDateTime', '')
            if received:
                dt = datetime.fromisoformat(received.replace('Z', '+00:00'))
                # Convert to Asia/Shanghai timezone
                dt_local = dt.astimezone(ZoneInfo('Asia/Shanghai'))
                received = dt_local.strftime('%Y-%m-%d %H:%M')

            from_addr = msg.get('from', {}).get('emailAddress', {})
            sender = from_addr.get('name', from_addr.get('address', 'Unknown'))

            subject = msg.get('subject', '(No Subject)')
            read_status = '[UNREAD] ' if not msg.get('isRead', True) else ''
            message_id = msg.get('id', '')

            print(f"\n{i}. {read_status}{subject}")
            print(f"   From: {sender}")
            print(f"   Date: {received}")
            print(f"   ID: {message_id}")
            
            # Show To recipients
            to_recipients = msg.get('toRecipients', [])
            if to_recipients:
                to_names = [r.get('emailAddress', {}).get('name', r.get('emailAddress', {}).get('address', '')) for r in to_recipients]
                to_display = ', '.join(to_names[:3])  # Show first 3 recipients
                if len(to_recipients) > 3:
                    to_display += f' (+{len(to_recipients) - 3} more)'
                print(f"   To: {to_display}")
            
            # Show CC recipients
            cc_recipients = msg.get('ccRecipients', [])
            if cc_recipients:
                cc_names = [r.get('emailAddress', {}).get('name', r.get('emailAddress', {}).get('address', '')) for r in cc_recipients]
                cc_display = ', '.join(cc_names[:3])  # Show first 3 recipients
                if len(cc_recipients) > 3:
                    cc_display += f' (+{len(cc_recipients) - 3} more)'
                print(f"   Cc: {cc_display}")
            
            # Show preview if available
            preview = msg.get('bodyPreview', '')
            if preview:
                # Limit preview to 150 characters
                preview = preview[:150].replace('\n', ' ').replace('\r', ' ')
                if len(msg.get('bodyPreview', '')) > 150:
                    preview += '...'
                print(f"   Preview: {preview}")
            
            if i < len(messages):
                print(f"   {'-'*78}")
        
        print(f"\n{'='*80}")
        print(f"Total: {len(messages)} messages")


def display_message(message: Dict):
    """Display a single message in detail."""
    import re

    print(f"\n{'='*80}")
    print(f"Subject: {message.get('subject', '(No Subject)')}")
    print(f"From: {message.get('from', {}).get('emailAddress', {})}")
    print(f"To: {[r.get('emailAddress', {}) for r in message.get('toRecipients', [])]}")
    print(f"CC: {[r.get('emailAddress', {}) for r in message.get('ccRecipients', [])]}")
    print(f"Date: {message.get('receivedDateTime', '')}")
    print(f"ID: {message.get('id', '')}")
    print(f"{'='*80}")
    
    # Get body content
    body = message.get('body', {})
    content = body.get('content', '')
    content_type = body.get('contentType', 'text')
    
    # If HTML, try to extract plain text
    if content_type == 'html' and content:
        # Remove HTML tags
        content = re.sub(r'<style[^>]*>.*?</style>', '', content, flags=re.DOTALL)
        content = re.sub(r'<[^>]+>', ' ', content)
        # Decode HTML entities
        import html
        content = html.unescape(content)
        # Clean up whitespace
        content = re.sub(r'\s+', ' ', content).strip()
        # Limit length
        if len(content) > MAX_MESSAGE_DISPLAY_LENGTH:
            content = content[:MAX_MESSAGE_DISPLAY_LENGTH] + '...'
    
    # Handle encoding issues
    try:
        print(f"\n{content}")
    except UnicodeEncodeError:
        # Fallback: encode to ASCII with replacement
        print(f"\n{content.encode('ascii', 'replace').decode('ascii')}")
    
    print(f"\n{'='*80}")


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Microsoft Graph Email Operations")
    subparsers = parser.add_subparsers(dest="command", required=True)
    
    # Global --json flag
    parser.add_argument("--json", action="store_true", help="Output in JSON format")
    
    # List command (now includes search functionality)
    list_parser = subparsers.add_parser("list", help="List/search messages")
    list_parser.add_argument("--folder", default="inbox", help="Folder name (or 'all' for all folders)")
    list_parser.add_argument("--limit", "--top", type=int, default=25, dest="limit", help="Max messages to return (--top is alias)")
    list_parser.add_argument("--filter", dest="filter_query", help="OData filter query")
    list_parser.add_argument("--unread", action="store_true", help="Show unread only")
    list_parser.add_argument("--preview", action="store_true", help="Show email body preview")
    list_parser.add_argument("--focused", action="store_true", help="Show Focused inbox only")
    list_parser.add_argument("--other", action="store_true", help="Show Other inbox only")
    # Search parameters
    list_parser.add_argument("--from", dest="from_sender", help="Search by sender name or email")
    list_parser.add_argument("--to", dest="to_recipient", help="Search by recipient name or email")
    list_parser.add_argument("--subject", help="Search by subject text")
    list_parser.add_argument("--body", help="Search by body text")
    
    # Search command (complete alias for list, supports all parameters)
    search_parser = subparsers.add_parser("search", help="Search/list messages (alias for list)")
    search_parser.add_argument("--folder", default="inbox", help="Folder name (or 'all' for all folders)")
    search_parser.add_argument("--limit", "--top", type=int, default=25, dest="limit", help="Max messages to return (--top is alias)")
    search_parser.add_argument("--filter", dest="filter_query", help="OData filter query")
    search_parser.add_argument("--unread", action="store_true", help="Show unread only")
    search_parser.add_argument("--preview", action="store_true", help="Show email body preview")
    search_parser.add_argument("--focused", action="store_true", help="Show Focused inbox only")
    search_parser.add_argument("--other", action="store_true", help="Show Other inbox only")
    # Search parameters
    search_parser.add_argument("--from", dest="from_sender", help="Search by sender name or email")
    search_parser.add_argument("--to", dest="to_recipient", help="Search by recipient name or email")
    search_parser.add_argument("--subject", help="Search by subject text")
    search_parser.add_argument("--body", help="Search by body text")
    
    # Find command (search + display first result)
    find_parser = subparsers.add_parser("find", help="Find and display a message (combines search + get)")
    find_parser.add_argument("--from", dest="from_sender", help="Sender name or email")
    find_parser.add_argument("--to", dest="to_recipient", help="Recipient name or email")
    find_parser.add_argument("--subject", help="Subject contains")
    find_parser.add_argument("--body", help="Body contains")
    find_parser.add_argument("--folder", default="inbox", help="Folder name (or 'all' for all folders)")
    
    # Get command
    get_parser = subparsers.add_parser("get", help="Get a message")
    get_parser.add_argument("message_id", help="Message ID")
    
    # Thread command
    thread_parser = subparsers.add_parser("thread", help="Get conversation thread")
    thread_parser.add_argument("message_id", help="Message ID (any message in the thread)")
    
    # Send command
    send_parser = subparsers.add_parser("send", help="Send an email")
    send_parser.add_argument("--to", required=True, help="To recipients (comma-separated)")
    send_parser.add_argument("--cc", help="CC recipients (comma-separated)")
    send_parser.add_argument("--bcc", help="BCC recipients (comma-separated)")
    send_parser.add_argument("--subject", required=True, help="Email subject")
    send_parser.add_argument("--body", required=True, help="Email body")
    send_parser.add_argument("--body-type", choices=["html", "text"], default="html")
    
    # Reply command
    reply_parser = subparsers.add_parser("reply", help="Reply to an email")
    reply_parser.add_argument("message_id", help="Message ID to reply to")
    reply_parser.add_argument("--body", required=True, help="Reply body")
    reply_parser.add_argument("--all", dest="reply_all", action="store_true", help="Reply to all")
    
    # Forward command
    forward_parser = subparsers.add_parser("forward", help="Forward an email")
    forward_parser.add_argument("message_id", help="Message ID to forward")
    forward_parser.add_argument("--to", required=True, help="To recipients (comma-separated)")
    forward_parser.add_argument("--cc", help="CC recipients (comma-separated)")
    forward_parser.add_argument("--bcc", help="BCC recipients (comma-separated)")
    forward_parser.add_argument("--comment", default="", help="Comment to add")
    
    # Mark read/unread
    read_parser = subparsers.add_parser("read", help="Mark message as read/unread")
    read_parser.add_argument("message_id", help="Message ID")
    read_parser.add_argument("--unread", action="store_true", help="Mark as unread")
    
    # Delete command
    delete_parser = subparsers.add_parser("delete", help="Delete an email")
    delete_parser.add_argument("message_id", help="Message ID to delete")
    
    args = parser.parse_args()
    
    try:
        if args.command in ["list", "search"]:
            # Unified handler for list and search commands (identical functionality)
            filter_query = args.filter_query
            if args.unread:
                filter_query = (filter_query + " and " if filter_query else "") + "isRead eq false"
            
            # Determine inference classification
            inference_classification = None
            if args.focused:
                inference_classification = "focused"
            elif args.other:
                inference_classification = "other"
            
            messages = list_messages(
                folder=args.folder,
                limit=args.limit,
                filter_query=filter_query,
                include_preview=True,  # Always include bodyPreview in API response for flexibility
                inference_classification=inference_classification,
                from_sender=getattr(args, 'from_sender', None),
                to_recipient=getattr(args, 'to_recipient', None),
                subject=getattr(args, 'subject', None),
                body=getattr(args, 'body', None)
            )
            if args.json:
                print(json.dumps({"success": True, "messages": messages, "total": len(messages)}, indent=2, default=str))
            else:
                display_message_list(messages, show_preview=True)
        
        elif args.command == "find":
            # Find uses list_messages with limit=1
            try:
                messages = list_messages(
                    folder=args.folder,
                    limit=1,
                    from_sender=args.from_sender,
                    to_recipient=args.to_recipient,
                    subject=args.subject,
                    body=args.body
                )
                if not messages:
                    print("No matching messages found.")
                    sys.exit(1)
                
                # Get full message details
                message = get_message(messages[0]['id'])
                if args.json:
                    print(json.dumps({"success": True, "message": message}, indent=2, default=str))
                else:
                    display_message(message)
            except Exception as e:
                # Auto-fallback to list --preview if find fails (e.g., due to search indexing delays)
                if "search" in str(e).lower() or "index" in str(e).lower():
                    print("⚠️  Search API failed (likely due to indexing delays)")
                    print("💡 Auto-falling back to list --preview...")
                    print()
                    messages = list_messages(
                        folder=args.folder,
                        limit=10,
                        from_sender=args.from_sender,
                        to_recipient=args.to_recipient,
                        subject=args.subject,
                        body=args.body,
                        include_preview=True
                    )
                    if args.json:
                        print(json.dumps({"success": True, "messages": messages, "total": len(messages)}, indent=2, default=str))
                    else:
                        display_message_list(messages, show_preview=True)
                else:
                    raise
        
        elif args.command == "get":
            message = get_message(args.message_id)
            if args.json:
                print(json.dumps({"success": True, "message": message}, indent=2, default=str))
            else:
                display_message(message)
        
        elif args.command == "thread":
            messages = get_message_thread(args.message_id)
            if args.json:
                print(json.dumps({"success": True, "messages": messages, "total": len(messages)}, indent=2, default=str))
            else:
                display_thread(messages)
        
        elif args.command == "send":
            send_email(
                to=parse_email_list(args.to),
                subject=args.subject,
                body=args.body,
                cc=parse_email_list(args.cc) if args.cc else None,
                bcc=parse_email_list(args.bcc) if args.bcc else None,
                body_type=args.body_type
            )
            if args.json:
                print(json.dumps({"success": True, "message": "Email sent successfully"}))
            else:
                print("✓ Email sent successfully")
        
        elif args.command == "reply":
            reply_email(
                message_id=args.message_id,
                body=args.body,
                reply_all=args.reply_all
            )
            if args.json:
                print(json.dumps({"success": True, "message": "Reply sent successfully"}))
            else:
                print("✓ Reply sent successfully")
        
        elif args.command == "forward":
            forward_email(
                message_id=args.message_id,
                to=parse_email_list(args.to),
                cc=parse_email_list(args.cc) if args.cc else None,
                bcc=parse_email_list(args.bcc) if args.bcc else None,
                comment=args.comment
            )
            if args.json:
                print(json.dumps({"success": True, "message": "Email forwarded successfully"}))
            else:
                print("✓ Email forwarded successfully")
        
        elif args.command == "read":
            if args.unread:
                mark_as_unread(args.message_id)
                if args.json:
                    print(json.dumps({"success": True, "message": "Marked as unread"}))
                else:
                    print("✓ Marked as unread")
            else:
                mark_as_read(args.message_id)
                if args.json:
                    print(json.dumps({"success": True, "message": "Marked as read"}))
                else:
                    print("✓ Marked as read")
        
        elif args.command == "delete":
            delete_email(args.message_id)
            if args.json:
                print(json.dumps({"success": True, "message": "Email deleted"}))
            else:
                print("✓ Email deleted")
    
    except ValueError as e:
        if args.json:
            print(json.dumps({"success": False, "error": str(e)}))
        else:
            print(f"Error: {e}")
        sys.exit(1)
    except Exception as e:
        if args.json:
            print(json.dumps({"success": False, "error": str(e)}))
        else:
            print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
