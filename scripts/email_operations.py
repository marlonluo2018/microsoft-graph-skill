#!/usr/bin/env python3
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
from typing import List, Optional, Dict, Any
from datetime import datetime

# Try to import requests
try:
    import requests
except ImportError:
    print("Error: requests package not found.")
    print("Install with: pip install requests")
    sys.exit(1)

# Import auth module
from auth import get_access_token, DEFAULT_SCOPES

# Constants
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
MAX_RECIPIENTS = 500  # Company policy limit


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
    print(f"\n{'='*80}")
    print(f"Subject: {message.get('subject', '(No Subject)')}")
    print(f"From: {message.get('from', {}).get('emailAddress', {})}")
    print(f"To: {[r.get('emailAddress', {}) for r in message.get('toRecipients', [])]}")
    print(f"CC: {[r.get('emailAddress', {}) for r in message.get('ccRecipients', [])]}")
    print(f"Date: {message.get('receivedDateTime', '')}")
    print(f"{'='*80}")
    print(f"\n{message.get('body', {}).get('content', '')}")
    print(f"\n{'='*80}")


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Microsoft Graph Email Operations")
    subparsers = parser.add_subparsers(dest="command", required=True)
    
    # List command
    list_parser = subparsers.add_parser("list", help="List messages")
    list_parser.add_argument("--folder", default="inbox", help="Folder name")
    list_parser.add_argument("--limit", type=int, default=25, help="Max messages to return")
    list_parser.add_argument("--filter", dest="filter_query", help="OData filter query")
    list_parser.add_argument("--unread", action="store_true", help="Show unread only")
    
    # Get command
    get_parser = subparsers.add_parser("get", help="Get a message")
    get_parser.add_argument("message_id", help="Message ID")
    
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
            display_message_list(messages)
        
        elif args.command == "get":
            message = get_message(args.message_id)
            display_message(message)
        
        elif args.command == "send":
            send_email(
                to=parse_email_list(args.to),
                subject=args.subject,
                body=args.body,
                cc=parse_email_list(args.cc) if args.cc else None,
                bcc=parse_email_list(args.bcc) if args.bcc else None,
                body_type=args.body_type
            )
            print("✓ Email sent successfully")
        
        elif args.command == "reply":
            reply_email(
                message_id=args.message_id,
                body=args.body,
                reply_all=args.reply_all
            )
            print("✓ Reply sent successfully")
        
        elif args.command == "forward":
            forward_email(
                message_id=args.message_id,
                to=parse_email_list(args.to),
                cc=parse_email_list(args.cc) if args.cc else None,
                bcc=parse_email_list(args.bcc) if args.bcc else None,
                comment=args.comment
            )
            print("✓ Email forwarded successfully")
        
        elif args.command == "read":
            if args.unread:
                mark_as_unread(args.message_id)
                print("✓ Marked as unread")
            else:
                mark_as_read(args.message_id)
                print("✓ Marked as read")
        
        elif args.command == "delete":
            delete_email(args.message_id)
            print("✓ Email deleted")
    
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
