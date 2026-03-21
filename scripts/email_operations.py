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
    token: str = None
) -> List[Dict[str, Any]]:
    """
    List messages from a folder.
    
    Args:
        folder: Folder name (inbox, sentitems, drafts, etc.) or folder ID
        limit: Maximum number of messages to return
        filter_query: OData filter query
        order_by: Sort order
        token: Access token (will obtain if not provided)
    
    Returns:
        List of message objects
    """
    if token is None:
        token = get_access_token()
    
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
    
    folder_id = folder_map.get(folder.lower(), folder)
    
    url = f"{GRAPH_API_BASE}/me/mailFolders/{folder_id}/messages"
    params = {
        "$top": limit,
        "$orderby": order_by,
        "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments"
    }
    
    if filter_query:
        params["$filter"] = filter_query
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to list messages: {response.status_code} - {response.text}")
    
    data = response.json()
    return data.get("value", [])


# =============================================================================
# SEARCH MESSAGES
# =============================================================================

def search_messages(
    from_sender: str = None,
    to_recipient: str = None,
    subject: str = None,
    body: str = None,
    folder: str = "inbox",
    limit: int = 25,
    token: str = None
) -> List[Dict[str, Any]]:
    """
    Search messages by various criteria.
    
    Uses $search for initial filtering, then applies client-side filtering
    for more specific criteria like sender name matching.
    
    Args:
        from_sender: Sender name or email to search for
        to_recipient: Recipient name or email to search for
        subject: Subject contains this text
        body: Body contains this text
        folder: Folder name or 'all' for all folders
        limit: Maximum number of messages to return
        token: Access token
    
    Returns:
        List of message objects
    """
    if token is None:
        token = get_access_token()
    
    # Build search keywords (for $search parameter)
    search_keywords = []
    if from_sender:
        search_keywords.append(from_sender)
    if to_recipient:
        search_keywords.append(to_recipient)
    if subject:
        search_keywords.append(subject)
    if body:
        search_keywords.append(body)
    
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
    
    if folder.lower() == "all":
        url = f"{GRAPH_API_BASE}/me/messages"
    else:
        folder_id = folder_map.get(folder.lower(), folder)
        url = f"{GRAPH_API_BASE}/me/mailFolders/{folder_id}/messages"
    
    # Use $search if we have keywords, otherwise just list
    if search_keywords:
        search_query = " ".join(search_keywords)
        params = {
            "$search": f'"{search_query}"',
            "$top": min(limit * 3, 100),  # Get more results for client-side filtering
            "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments"
        }
    else:
        params = {
            "$top": limit,
            "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments",
            "$orderby": "receivedDateTime desc"
        }
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to search messages: {response.status_code} - {response.text}")
    
    data = response.json()
    messages = data.get("value", [])
    
    # Apply client-side filtering for more precise matching
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
    
    # Limit results
    return messages[:limit]


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
# REPLY TO EMAIL
# =============================================================================

def reply_email(
    message_id: str,
    body: str,
    reply_all: bool = False,
    body_type: str = "html",
    token: str = None
) -> bool:
    """
    Reply to an email.
    
    Args:
        message_id: ID of message to reply to
        body: Reply body content
        reply_all: If True, reply to all recipients; otherwise reply to sender only
        body_type: "html" or "text"
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    endpoint = "replyAll" if reply_all else "reply"
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}/{endpoint}"
    
    payload = {
        "message": {
            "body": {
                "contentType": body_type,
                "content": body
            }
        }
    }
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code != 202:
        raise Exception(f"Failed to reply to email: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# FORWARD EMAIL
# =============================================================================

def forward_email(
    message_id: str,
    to: List[str],
    cc: List[str] = None,
    bcc: List[str] = None,
    comment: str = "",
    token: str = None
) -> bool:
    """
    Forward an email.
    
    Args:
        message_id: ID of message to forward
        to: List of To recipient emails
        cc: List of CC recipient emails
        bcc: List of BCC recipient emails
        comment: Optional comment to add
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    # Validate recipients
    validate_recipients(to, cc, bcc)
    
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}/forward"
    
    payload = {
        "toRecipients": [format_email_address(e) for e in to],
        "ccRecipients": [format_email_address(e) for e in (cc or [])],
        "bccRecipients": [format_email_address(e) for e in (bcc or [])],
        "comment": comment
    }
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code != 202:
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

def display_message_list(messages: List[Dict]):
    """Display a list of messages in a readable format."""
    print(f"\n{'='*80}")
    print(f"{'Date':<25} {'From':<30} {'Subject':<40}")
    print(f"{'='*80}")
    
    for msg in messages:
        received = msg.get('receivedDateTime', '')
        if received:
            dt = datetime.fromisoformat(received.replace('Z', '+00:00'))
            received = dt.strftime('%Y-%m-%d %H:%M')
        
        from_addr = msg.get('from', {}).get('emailAddress', {})
        sender = from_addr.get('name', from_addr.get('address', 'Unknown'))
        
        subject = msg.get('subject', '(No Subject)')[:40]
        read_status = '' if msg.get('isRead', True) else '[UNREAD]'
        
        print(f"{received:<25} {sender:<30} {subject}{read_status}")
    
    print(f"{'='*80}")
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
    
    # List command
    list_parser = subparsers.add_parser("list", help="List messages")
    list_parser.add_argument("--folder", default="inbox", help="Folder name")
    list_parser.add_argument("--limit", type=int, default=25, help="Max messages to return")
    list_parser.add_argument("--filter", dest="filter_query", help="OData filter query")
    list_parser.add_argument("--unread", action="store_true", help="Show unread only")
    
    # Search command
    search_parser = subparsers.add_parser("search", help="Search messages")
    search_parser.add_argument("--from", dest="from_sender", help="Sender name or email")
    search_parser.add_argument("--to", dest="to_recipient", help="Recipient name or email")
    search_parser.add_argument("--subject", help="Subject contains")
    search_parser.add_argument("--body", help="Body contains")
    search_parser.add_argument("--folder", default="inbox", help="Folder name (or 'all' for all folders)")
    search_parser.add_argument("--limit", type=int, default=25, help="Max messages to return")
    
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
        if args.command == "list":
            filter_query = args.filter_query
            if args.unread:
                filter_query = (filter_query + " and " if filter_query else "") + "isRead eq false"
            
            messages = list_messages(
                folder=args.folder,
                limit=args.limit,
                filter_query=filter_query
            )
            if args.json:
                print(json.dumps({"success": True, "messages": messages, "total": len(messages)}, indent=2, default=str))
            else:
                display_message_list(messages)
        
        elif args.command == "search":
            messages = search_messages(
                from_sender=args.from_sender,
                to_recipient=args.to_recipient,
                subject=args.subject,
                body=args.body,
                folder=args.folder,
                limit=args.limit
            )
            if args.json:
                print(json.dumps({"success": True, "messages": messages, "total": len(messages)}, indent=2, default=str))
            else:
                display_message_list(messages)
        
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
