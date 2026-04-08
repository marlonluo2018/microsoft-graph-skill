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
import io
import json
import argparse
import csv
import time
from pathlib import Path
from typing import List, Optional, Dict, Any
from datetime import datetime
from zoneinfo import ZoneInfo

# Force UTF-8 encoding for stdout to handle emojis on Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Fix Windows console encoding (alternative method)
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

# NOTE: Time filtering is supported via --start/--end arguments.
# Timestamp MUST include timezone (e.g., "2026-03-26T12:00:00+08:00" or "2026-03-26T04:00:00Z").
# Special value "now" can be used for current time.
# If not specified, all messages are returned (subject to --limit).

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





def api_request(
    method: str,
    url: str,
    token: str = None,
    max_retries: int = 3,
    base_delay: float = 1.0,
    **kwargs
) -> "requests.Response":
    """
    Make an API request with automatic retry on rate limiting (429).
    
    Args:
        method: HTTP method ('get', 'post', or 'patch')
        url: API endpoint URL
        token: Access token (will obtain if not provided)
        max_retries: Maximum number of retries for 429 errors
        base_delay: Base delay in seconds for exponential backoff
        **kwargs: Additional arguments passed to requests
        
    Returns:
        requests.Response object
        
    Raises:
        Exception: If request fails after all retries
    """
    if token is None:
        token = get_access_token()
    
    headers = kwargs.pop('headers', get_headers(token))
    params = kwargs.pop('params', None)
    json_data = kwargs.pop('json', None)
    
    for attempt in range(max_retries + 1):
        if method.lower() == 'get':
            response = requests.get(url, headers=headers, params=params, **kwargs)
        elif method.lower() == 'post':
            response = requests.post(url, headers=headers, json=json_data, params=params, **kwargs)
        elif method.lower() == 'patch':
            response = requests.patch(url, headers=headers, json=json_data, params=params, **kwargs)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")
        
        # Check for rate limiting
        if response.status_code == 429:
            if attempt < max_retries:
                # Get retry-after header or use exponential backoff
                retry_after = response.headers.get('Retry-After')
                if retry_after:
                    delay = float(retry_after)
                else:
                    delay = base_delay * (2 ** attempt)  # Exponential backoff
                
                print(f"⚠️ Rate limited (429). Retrying in {delay:.1f}s... (attempt {attempt + 1}/{max_retries})")
                time.sleep(delay)
                continue
            else:
                raise Exception(
                    f"Rate limit exceeded after {max_retries} retries. "
                    f"Please wait a few minutes before trying again."
                )
        
        # For other errors, raise immediately
        if response.status_code >= 400:
            raise Exception(f"API request failed: {response.status_code} - {response.text}")
        
        return response
    
    raise Exception("Unexpected error in api_request")


def validate_recipients(to: List[str], cc: List[str] = None, bcc: List[str] = None) -> bool:
    """
    Validate that total recipients don't exceed company limit.
    
    Args:
        to: List of To recipients (can be None)
        cc: List of CC recipients
        bcc: List of BCC recipients
    
    Returns:
        bool: True if valid, raises ValueError otherwise
    """
    total = len(to or []) + len(cc or []) + len(bcc or [])
    
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


def detect_outlook_syntax(value: str, param_name: str) -> tuple:
    """
    Detect and convert Outlook search syntax to CLI parameters.
    
    Outlook syntax: "from:beng", "to:john@example.com", "subject:meeting"
    Returns: (converted_value, warning_message)
    
    Examples:
        detect_outlook_syntax("from:beng", "--from") -> ("beng", "⚠️  Auto-converted 'from:beng' -> --from 'beng'")
        detect_outlook_syntax("beng", "--from") -> ("beng", None)
    """
    if not value or ':' not in value:
        return value, None
    
    # Outlook syntax patterns
    outlook_patterns = {
        'from': 'from_sender',
        'to': 'to_recipient',
        'subject': 'subject',
        'body': 'body',
    }
    
    # Check if value matches Outlook syntax
    for pattern, target_param in outlook_patterns.items():
        if value.lower().startswith(f"{pattern}:"):
            # Extract the actual value
            actual_value = value[len(pattern)+1:].strip()
            # Remove surrounding quotes if present
            if actual_value.startswith('"') and actual_value.endswith('"'):
                actual_value = actual_value[1:-1]
            if actual_value.startswith("'") and actual_value.endswith("'"):
                actual_value = actual_value[1:-1]
            
            # Check if the detected param matches the CLI param
            if param_name.replace('--', '').replace('_', '') in [target_param, pattern]:
                warning = f"⚠️  Auto-converted Outlook syntax: '{value}' -> {param_name} '{actual_value}'"
                return actual_value, warning
            else:
                # User used wrong param, but we detected what they meant
                warning = f"⚠️  Detected Outlook syntax in {param_name}: '{value}'"
                warning += f"\n   💡 Did you mean --{pattern} '{actual_value}'?"
                return value, warning
    
    return value, None


def convert_outlook_syntax_args(args) -> list:
    """
    Process args and convert any Outlook syntax found.
    Returns list of warning messages.
    """
    warnings = []
    
    params_to_check = [
        ('from_sender', '--from'),
        ('to_recipient', '--to'),
        ('subject', '--subject'),
        ('body', '--body'),
    ]
    
    for attr, param_name in params_to_check:
        value = getattr(args, attr, None)
        if value:
            converted, warning = detect_outlook_syntax(value, param_name)
            if converted != value:
                setattr(args, attr, converted)
            if warning:
                warnings.append(warning)
    
    return warnings


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
    token: str = None,
    display_timezone: str = None,
    start: str = None,
    end: str = None,
    message_type: str = "all"
) -> tuple:
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
        start: ISO 8601 timestamp to filter messages received AFTER this time (requires timezone, or "now")
        end: ISO 8601 timestamp to filter messages received BEFORE this time (requires timezone, or "now")
        message_type: Filter by message type - "all" (default), "emails", or "events"
        display_timezone: Timezone for display and "now" calculation (default: Asia/Shanghai)
    
    Returns:
        Tuple of (list of message objects, time_info dict)
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
    # Note: @odata.type is automatically included in responses, no need to select it
    select_fields = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,inferenceClassification"
    if include_preview:
        select_fields += ",bodyPreview"
    
    # Build parameters
    params = {
        "$select": select_fields
    }
    
    # Add $expand for attachments and event messages
    expand_clauses = []
    
    # Always expand attachments to show file names
    expand_clauses.append("attachments($select=name,contentType,size,isInline)")
    
    # Add event expansion when we might show events (message_type is "all" or "events")
    if message_type in ["all", "events"]:
        expand_clauses.append("microsoft.graph.eventMessage/event($select=responseStatus,start,end,location)")
    
    if expand_clauses:
        params["$expand"] = ",".join(expand_clauses)
    
    # If search criteria provided, use $search
    # Track if start is handled in $search (KQL syntax) to avoid duplicate filtering
    start_in_search = False
    
    if has_search_criteria:
        search_keywords = []
        if from_sender:
            search_keywords.append(f"from:{from_sender}")
        if to_recipient:
            search_keywords.append(f"to:{to_recipient}")
        if subject:
            search_keywords.append(f"subject:{subject}")
        if body:
            search_keywords.append(body)
        
        # KQL: Add start to $search whenever $search is used
        # This avoids $search + $filter conflict (Graph API doesn't allow both)
        # NOTE: KQL only supports date format (YYYY-MM-DD)
        if start:
            # Handle "now" special case - convert to actual timestamp first
            start_for_validation = start
            if start.lower() == "now":
                tz = ZoneInfo(display_timezone or "Asia/Shanghai")
                start_for_validation = datetime.now(tz).strftime('%Y-%m-%dT%H:%M:%S%z')
                # Format offset as +08:00
                offset = start_for_validation[-5:]
                start_for_validation = start_for_validation[:-2] + ':' + start_for_validation[-2:]
            else:
                # For plain timestamps, validate that --timezone is provided
                has_embedded_tz = (start_for_validation.endswith('Z') or
                                 '+' in start_for_validation or
                                 (start_for_validation.count('-') > 2 and '-' in start_for_validation[-6:]))
                
                if not has_embedded_tz:
                    # Plain timestamp - require --timezone parameter
                    if not display_timezone:
                        raise ValueError(
                            f"TIMEZONE_REQUIRED: '{start}' is missing timezone information.\n"
                            f"--timezone parameter is required when using plain timestamps.\n"
                            f"Valid formats:\n"
                            f"  - Plain with --timezone: '2026-03-26T12:00:00' --timezone 'Asia/Shanghai'\n"
                            f"  - UTC time: '2026-03-26T04:00:00Z'\n"
                            f"  - With timezone offset: '2026-03-26T12:00:00+08:00'\n"
                            f"  - Current time: 'now'"
                        )
                    # Convert plain timestamp to timezone-aware for validation
                    try:
                        tz = ZoneInfo(display_timezone)
                        local_dt = datetime.fromisoformat(start_for_validation)
                        local_dt = local_dt.replace(tzinfo=tz)
                        start_for_validation = local_dt.strftime('%Y-%m-%dT%H:%M:%S%z')
                        # Format offset as +08:00
                        offset = start_for_validation[-5:]
                        start_for_validation = start_for_validation[:-2] + ':' + start_for_validation[-2:]
                    except Exception as e:
                        raise ValueError(f"Invalid timestamp or timezone: {start} - {str(e)}")
            
            # Convert start to KQL date format (YYYY-MM-DD) - only date is supported by KQL
            start_date = start_for_validation.split('T')[0] if 'T' in start_for_validation else start_for_validation
            search_keywords.append(f"received>={start_date}")
            start_in_search = True  # Mark as handled
        
        search_query = " ".join(search_keywords)
        params["$search"] = f'"{search_query}"'
        params["$top"] = limit  # No need for extra results, server filters
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
    
    # Add time filter for incremental sync (messages received within --start/--end range)
    # STRICT TIMEZONE-REQUIRED design:
    #   - UTC time (e.g., "2026-03-26T00:00:00Z") → ACCEPTED → use directly
    #   - With timezone offset (e.g., "2026-03-26T08:00:00+08:00") → ACCEPTED → convert to UTC
    #   - "now" → ACCEPTED → use current time with timezone from display_timezone param
    #   - Pure date (e.g., "2026-03-26") → REJECTED (ambiguous)
    #   - Time WITHOUT timezone (e.g., "2026-03-26T15:00:00") → REJECTED (ambiguous)
    # NOTE: If start is already in $search (KQL), skip $filter to avoid conflict
    start_info = None  # Track timezone conversion info for display
    end_info = None  # Track end timezone conversion info
    
    def convert_timestamp_to_utc(timestamp: str, param_name: str, display_timezone: str = None) -> tuple:
        """
        Convert timestamp to UTC. Returns (utc_timestamp, timezone_offset, error).
        
        Args:
            timestamp: Plain datetime string (no timezone) or "now"
            param_name: Parameter name for error messages
            display_timezone: REQUIRED - timezone for parsing and display
        """
        if not timestamp:
            return None, None, None
        
        # display_timezone is always required
        if not display_timezone:
            return None, None, (
                f"TIMEZONE_REQUIRED: --timezone is required.\n"
                f"Example: --{param_name} '2026-03-26T12:00:00' --timezone 'Asia/Shanghai'"
            )
        
        # Calculate display offset from display_timezone
        tz = ZoneInfo(display_timezone)
        sample_dt = datetime.now(tz)
        offset = sample_dt.strftime('%z')
        tz_offset = f"{offset[:3]}:{offset[3:]}"  # Format as +08:00
        
        # Handle "now" special value
        if timestamp.lower() == "now":
            now_dt = datetime.now(tz)
            utc_time = now_dt.astimezone(ZoneInfo('UTC')).strftime('%Y-%m-%dT%H:%M:%SZ')
            return utc_time, tz_offset, None
        
        # Reject embedded timezone (Z or +08:00)
        if timestamp.endswith('Z') or '+' in timestamp or (timestamp.count('-') > 2):
            return None, None, (
                f"INVALID_FORMAT: '{timestamp}' contains embedded timezone.\n"
                f"Use plain datetime with --timezone instead:\n"
                f"  --{param_name} '2026-03-26T12:00:00' --timezone 'Asia/Shanghai'"
            )
        
        try:
            # Parse plain datetime and add timezone
            local_dt = datetime.fromisoformat(timestamp)
            local_dt = local_dt.replace(tzinfo=tz)
            utc_time = local_dt.astimezone(ZoneInfo('UTC')).strftime('%Y-%m-%dT%H:%M:%SZ')
            return utc_time, tz_offset, None
                
        except (ValueError, TypeError) as e:
            return None, None, f"Invalid timestamp format: {timestamp} - {str(e)}"
    
    # Process --start parameter
    if start and not start_in_search:
        start_utc, tz_offset, error = convert_timestamp_to_utc(start, 'start', display_timezone)
        if error:
            raise ValueError(error)
        
        start_info = {
            'original': start,
            'converted_utc': start_utc,
            'timezone': tz_offset,
            'timezone_offset': tz_offset
        }
        filters.append(f"receivedDateTime gt {start_utc}")
    
    # Process --end parameter
    if end:
        end_utc, tz_offset, error = convert_timestamp_to_utc(end, 'end', display_timezone)
        if error:
            raise ValueError(error)
        
        end_info = {
            'original': end,
            'converted_utc': end_utc,
            'timezone': tz_offset,
            'timezone_offset': tz_offset
        }
        filters.append(f"receivedDateTime lt {end_utc}")
    
    # IMPORTANT: Microsoft Graph API doesn't allow $search and $filter together
    # If using $search (has_search_criteria), we must skip $filter entirely
    # Time filters should be handled via KQL in $search (already done above)
    if filters and not has_search_criteria:
        params["$filter"] = " and ".join(filters)
    
    response = api_request('get', url, token, params=params)
    
    data = response.json()
    messages = data.get("value", [])
    
    # Server-side filtering is now complete via KQL $search
    # No client-side filtering needed
    
    # Filter by message type (client-side)
    # Meeting invites are identified by @odata.type containing "eventMessage"
    if message_type == "emails":
        # Only show regular emails (exclude meeting invites)
        messages = [m for m in messages if 'eventMessage' not in m.get('@odata.type', '')]
    elif message_type == "events":
        # Only show meeting invites
        messages = [m for m in messages if 'eventMessage' in m.get('@odata.type', '')]
    # "all" = no filtering, show everything
    
    # Build time info for display
    time_info = None
    if start_info or end_info:
        time_info = {
            'start': start_info,
            'end': end_info
        }
    
    return messages, time_info




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
    
    # Expand attachments to get full details including IDs
    params = {
        "$expand": "attachments($select=id,name,contentType,size,isInline)"
    }
    
    response = api_request('get', url, token, params=params)
    
    return response.json()

def get_event_from_message(message_id: str, token: str = None) -> Dict[str, Any]:
    """
    Get the event associated with a meeting invite message.
    
    Args:
        message_id: Message ID of the meeting invite
        token: Access token
    
    Returns:
        Event object with event ID
    
    Raises:
        Exception: If message is not a meeting invite or event not found
    """
    if token is None:
        token = get_access_token()
    
    # Get the message with expanded event data
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}"
    params = {"$expand": "microsoft.graph.eventMessage/event"}
    
    response = api_request('get', url, token, params=params)
    message = response.json()
    
    # Check if this is a meeting invite
    if 'eventMessage' not in message.get('@odata.type', ''):
        raise ValueError("This message is not a meeting invite")
    
    # Extract the event from the expanded data
    if 'event' not in message:
        raise ValueError("Could not retrieve event data from meeting invite")
    
    return message['event']


def accept_meeting_invite(
    message_id: str,
    comment: str = None,
    send_response: bool = True,
    token: str = None
) -> bool:
    """
    Accept a meeting invite from a message.
    
    Args:
        message_id: Message ID of the meeting invite
        comment: Optional comment to include in response
        send_response: Whether to send response to organizer
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    # Get the event ID from the message
    event = get_event_from_message(message_id, token)
    event_id = event['id']
    
    # Accept the event using calendar API
    url = f"{GRAPH_API_BASE}/me/events/{event_id}/accept"
    
    payload = {
        "sendResponse": send_response
    }
    
    if comment:
        payload["comment"] = comment
    
    response = api_request('post', url, token, json=payload)
    
    return response.status_code == 202


def decline_meeting_invite(
    message_id: str,
    comment: str = None,
    send_response: bool = True,
    token: str = None
) -> bool:
    """
    Decline a meeting invite from a message.
    
    Args:
        message_id: Message ID of the meeting invite
        comment: Optional comment to include in response
        send_response: Whether to send response to organizer
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    # Get the event ID from the message
    event = get_event_from_message(message_id, token)
    event_id = event['id']
    
    # Decline the event using calendar API
    url = f"{GRAPH_API_BASE}/me/events/{event_id}/decline"
    
    payload = {
        "sendResponse": send_response
    }
    
    if comment:
        payload["comment"] = comment
    
    response = api_request('post', url, token, json=payload)
    
    return response.status_code == 202



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
    importance: str = None,
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
        importance: "low", "normal", or "high"
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    # Validate recipients
    validate_recipients(to, cc, bcc)
    
    # Convert plain text body to HTML if needed (for proper line breaks)
    if body_type == "html" and not body.strip().startswith('<'):
        body = body.replace('\n', '<br>')
    
    # Build message payload
    message = {
        "subject": subject,
        "body": {
            "contentType": body_type,
            "content": body
        },
        "toRecipients": [format_email_address(e) for e in (to or [])],
        "ccRecipients": [format_email_address(e) for e in (cc or [])],
        "bccRecipients": [format_email_address(e) for e in (bcc or [])]
    }
    
    # Add importance if specified
    if importance and importance.lower() in ['low', 'normal', 'high']:
        message["importance"] = importance.lower()
    
    # Add attachments if provided
    if attachments:
        message["attachments"] = attachments
    
    # Build request payload
    payload = {
        "message": message,
        "saveToSentItems": save_to_sent
    }
    
    url = f"{GRAPH_API_BASE}/me/sendMail"
    
    response = api_request('post', url, token, json=payload)
    
    return True


def batch_send_email(
    to: List[str],
    subject: str,
    body: str,
    cc: List[str] = None,
    bcc: List[str] = None,
    csv_path: str = None,
    email_column: str = None,
    body_type: str = "html",
    attachments: List[Dict] = None,
    importance: str = None,
    token: str = None
) -> Dict[str, Any]:
    """
    Send an email to multiple recipients with automatic batching.
    
    When total recipients exceed MAX_RECIPIENTS_PER_EMAIL (500),
    automatically splits them into multiple batches and sends multiple emails.
    
    BCC recipients can be loaded from a CSV file for mass mailing.
    
    Args:
        to: List of To recipient emails
        subject: Email subject
        body: Email body content
        cc: List of CC recipient emails
        bcc: List of BCC recipient emails
        csv_path: Path to CSV file containing BCC email addresses
        email_column: Column name in CSV for email addresses (auto-detect if not specified)
        body_type: "html" or "text"
        attachments: List of attachment objects
        importance: "low", "normal", or "high"
        token: Access token
    
    Returns:
        Dictionary with batch processing results
    """
    if token is None:
        token = get_access_token()
    
    # Get BCC recipients from CSV if provided
    if csv_path:
        print(f"📖 Reading BCC recipients from CSV: {csv_path}")
        csv_bcc = read_recipients_from_csv(csv_path, email_column)
        print(f"✓ Found {len(csv_bcc)} BCC recipients in CSV")
        # Merge with any manually specified BCC
        bcc = (bcc or []) + csv_bcc
    
    # Calculate total recipients
    to_count = len(to) if to else 0
    cc_count = len(cc) if cc else 0
    bcc_count = len(bcc) if bcc else 0
    total_recipients = to_count + cc_count + bcc_count
    
    # If within limit, use regular send_email
    if total_recipients <= MAX_RECIPIENTS:
        send_email(to, subject, body, cc, bcc, body_type, attachments, True, importance, token)
        return {
            "success": True,
            "message": "Email sent successfully",
            "total_recipients": total_recipients,
            "total_batches": 1,
            "sent_count": total_recipients,
            "failed_count": 0,
            "batch_size": MAX_RECIPIENTS,
            "errors": []
        }
    
    # Split recipients into batches
    # IMPORTANT: To and CC recipients must be in EVERY batch (they're not distributed)
    # Only BCC recipients are distributed across batches
    # This ensures we stay under the 500 recipient limit per batch
    sent_count = 0
    failed_count = 0
    errors = []
    
    # Calculate how many BCC recipients can fit in each batch
    # Reserve space for To and CC recipients that will be in every batch
    fixed_recipients_count = to_count + cc_count
    bcc_per_batch = MAX_RECIPIENTS - fixed_recipients_count
    
    if bcc_per_batch <= 0:
        raise ValueError(
            f"To ({to_count}) + CC ({cc_count}) recipients already exceed or equal the limit of {MAX_RECIPIENTS}. "
            f"Cannot add BCC recipients. Please reduce To/CC recipients."
        )
    
    # Split BCC recipients into batches
    batches = []
    bcc_list = bcc or []
    
    for i in range(0, len(bcc_list), bcc_per_batch):
        batch_bcc = bcc_list[i:i + bcc_per_batch]
        batches.append({
            "to": to if to else None,  # Same To in every batch
            "cc": cc if cc else None,  # Same CC in every batch
            "bcc": batch_bcc
        })
    
    # Send each batch
    for idx, batch in enumerate(batches, 1):
        batch_to_count = len(batch["to"]) if batch["to"] else 0
        batch_cc_count = len(batch["cc"]) if batch["cc"] else 0
        batch_bcc_count = len(batch["bcc"]) if batch["bcc"] else 0
        batch_total = batch_to_count + batch_cc_count + batch_bcc_count
        
        try:
            send_email(
                to=batch["to"],
                subject=subject,
                body=body,
                cc=batch["cc"],
                bcc=batch["bcc"],
                body_type=body_type,
                attachments=attachments,
                importance=importance,
                token=token
            )
            sent_count += batch_total
            print(f"✓ Batch {idx}/{len(batches)} sent successfully ({batch_total} recipients)")
        except Exception as e:
            failed_count += batch_total
            error_msg = f"Batch {idx} ({batch_total} recipients): {str(e)}"
            errors.append(error_msg)
            print(f"✗ Batch {idx}/{len(batches)} failed: {error_msg}")
    
    # Prepare response
    response = {
        "success": True,
        "message": f"Email sending completed in {len(batches)} batches",
        "total_recipients": total_recipients,
        "total_batches": len(batches),
        "sent_count": sent_count,
        "failed_count": failed_count,
        "batch_size": MAX_RECIPIENTS,
    }
    
    if errors:
        response["errors"] = errors
    
    return response


def read_recipients_from_csv(csv_path: str, email_column: str = None) -> List[str]:
    """
    Read email addresses from a CSV file.
    
    Args:
        csv_path: Path to the CSV file
        email_column: Column name containing email addresses (auto-detect if not specified)
    
    Returns:
        List of email addresses
    """
    emails = []
    
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        
        # Auto-detect email column if not specified
        if not email_column:
            # Common email column names
            email_columns = ['email', 'Email', 'EMAIL', 'email_address', 'EmailAddress', 
                           'Email Address', '邮箱', '邮件', '地址', 'e-mail', 'E-mail']
            
            available_columns = reader.fieldnames or []
            for col in email_columns:
                if col in available_columns:
                    email_column = col
                    break
            
            # If still not found, use the first column
            if not email_column and available_columns:
                email_column = available_columns[0]
                print(f"⚠️  Auto-detected column '{email_column}' as email column")
        
        if not email_column:
            raise ValueError("Could not determine email column in CSV file")
        
        for row in reader:
            email = row.get(email_column, '').strip()
            if email and '@' in email:
                emails.append(email)
    
    return emails


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
    importance: str = None,
    token: str = None
) -> bool:
    """
    Reply to an email using Microsoft Graph's native reply endpoint.
    
    This preserves inline attachments (images with cid: references) automatically
    because Microsoft handles the attachment copying internally.
    
    Args:
        message_id: ID of message to reply to
        body: Reply body content (comment to add at the top)
        reply_all: If True, reply to all recipients; otherwise reply to sender only
        body_type: "html" or "text"
        include_history: If True, include original message in reply body (default: True)
                        Note: Graph API native reply always includes original message
                        in the conversation thread format.
        importance: Not supported by native reply endpoint (will show warning)
        token: Access token
    
    Returns:
        bool: True if successful
    
    Note:
        Unlike the old implementation that created a new email, this uses Graph API's
        native /reply or /replyAll endpoint which automatically:
        - Preserves inline attachments (cid: references work correctly)
        - Sets RE: prefix on subject
        - Adds to the same conversation thread
        - Maintains proper email threading
    """
    if token is None:
        token = get_access_token()
    
    # Get original message to check if replying to self-sent email
    original_msg = get_message(message_id, token)
    
    # Check if replying to self-sent email
    from_field = original_msg.get('from', {})
    from_addr = from_field.get('emailAddress', from_field)  # Fallback to direct format
    my_email = get_my_email(token)
    
    if from_addr.get('address', '').lower() == my_email.lower():
        print("⚠️  WARNING: You are replying to your own sent email.")
        print("💡 TIP: Use 'forward' command instead to forward your sent email to others.")
        print("   Example: forward <message_id> --to \"recipient@example.com\"")
        print()
    
    # Warn if importance is specified (not supported by native reply)
    if importance:
        print("⚠️  Note: Importance level is not supported by the native reply endpoint.")
        print("   The reply will be sent with normal importance.")
    
    # Prepare the comment/body
    # Graph API expects the comment in a specific format
    comment = body
    
    # Convert plain text to HTML if needed
    if body_type == "html" and not comment.strip().startswith('<'):
        # Convert literal \n strings to actual newlines (for CLI convenience)
        comment = comment.replace('\\n', '\n')
        # Convert plain text to HTML with proper line breaks
        comment = comment.replace('\n', '<br>')
    
    # Determine which endpoint to use
    # /reply - reply to sender only
    # /replyAll - reply to all recipients
    endpoint = "replyAll" if reply_all else "reply"
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}/{endpoint}"
    
    # Build payload
    # Graph API reply endpoint expects a "comment" field
    # The original message will be automatically included below the comment
    payload = {
        "comment": comment
    }
    
    # Send the reply request
    response = api_request('post', url, token, json=payload)
    
    return True


def batch_reply_email(
    message_id: str,
    body: str,
    to: List[str] = None,
    cc: List[str] = None,
    bcc: List[str] = None,
    csv_path: str = None,
    email_column: str = None,
    body_type: str = "html",
    include_history: bool = True,
    attachments: List[Dict] = None,
    importance: str = None,
    token: str = None
) -> Dict[str, Any]:
    """
    Reply to an email with custom recipients and automatic batching.
    
    When total recipients exceed MAX_RECIPIENTS_PER_EMAIL (500),
    automatically splits them into multiple batches and sends multiple emails.
    
    BCC recipients can be loaded from a CSV file for mass mailing.
    
    Args:
        message_id: ID of message to reply to
        body: Reply body content
        to: Custom To recipients (overrides default reply behavior)
        cc: Custom CC recipients
        bcc: Custom BCC recipients
        csv_path: Path to CSV file containing BCC email addresses
        email_column: Column name in CSV for email addresses (auto-detect if not specified)
        body_type: "html" or "text"
        include_history: If True, include original message in reply body
        attachments: List of attachment objects (prepared by prepare_file_attachments)
        importance: "low", "normal", or "high"
        token: Access token
    
    Returns:
        Dictionary with batch processing results
    """
    if token is None:
        token = get_access_token()
    
    # Get BCC recipients from CSV if provided
    if csv_path:
        print(f"📖 Reading BCC recipients from CSV: {csv_path}")
        csv_bcc = read_recipients_from_csv(csv_path, email_column)
        print(f"✓ Found {len(csv_bcc)} BCC recipients in CSV")
        # Merge with any manually specified BCC
        bcc = (bcc or []) + csv_bcc
    
    # Calculate total recipients
    to_count = len(to) if to else 0
    cc_count = len(cc) if cc else 0
    bcc_count = len(bcc) if bcc else 0
    total_recipients = to_count + cc_count + bcc_count
    
    # Get original message for subject and recipients
    original_msg = get_message(message_id, token)
    subject = original_msg.get('subject', '')
    if not subject.upper().startswith('RE:'):
        subject = f"RE: {subject}"
    
    # If no custom recipients specified, extract from original message (reply-all behavior)
    if not to and not cc and not bcc:
        # Get current user's email to exclude from recipients
        my_email = get_my_email(token)
        
        # Extract sender from original message
        from_addr = original_msg.get('from', {}).get('emailAddress', {})
        sender_email = from_addr.get('address', '')
        
        # Extract other To recipients (excluding current user)
        to_recipients = original_msg.get('toRecipients', [])
        other_to_emails = [
            r.get('emailAddress', {}).get('address', '')
            for r in to_recipients
            if r.get('emailAddress', {}).get('address', '').lower() != my_email.lower()
        ]
        
        # Extract CC recipients
        cc_recipients = original_msg.get('ccRecipients', [])
        cc_emails = [
            r.get('emailAddress', {}).get('address', '')
            for r in cc_recipients
        ]
        
        # Build To list: sender + other To recipients (excluding self)
        to = []
        if sender_email:
            to.append(sender_email)
        to.extend(other_to_emails)
        
        # Build CC list from original CC recipients
        if cc_emails:
            cc = cc_emails
        
        if to:
            print(f"ℹ️  Auto-extracted recipients from original email:")
            print(f"   To: {', '.join(to)}")
            if cc:
                print(f"   CC: {', '.join(cc)}")
    
    # Build email body with history
    if include_history and body_type == "html":
        if not body.strip().startswith('<'):
            body = body.replace('\n', '<br>')
        full_body = body + "\n\n" + format_email_as_html(original_msg)
    else:
        full_body = body
    
    # If within limit, use regular send_email
    if total_recipients <= MAX_RECIPIENTS:
        send_email(to, subject, full_body, cc, bcc, body_type, attachments, True, importance, token)
        return {
            "success": True,
            "message": "Reply sent successfully",
            "total_recipients": total_recipients,
            "total_batches": 1,
            "sent_count": total_recipients,
            "failed_count": 0,
            "batch_size": MAX_RECIPIENTS,
            "errors": []
        }
    
    # Split recipients into batches
    # IMPORTANT: To and CC recipients must be in EVERY batch (they're not distributed)
    # Only BCC recipients are distributed across batches
    # This ensures we stay under the 500 recipient limit per batch
    sent_count = 0
    failed_count = 0
    errors = []
    
    # Calculate how many BCC recipients can fit in each batch
    # Reserve space for To and CC recipients that will be in every batch
    fixed_recipients_count = to_count + cc_count
    bcc_per_batch = MAX_RECIPIENTS - fixed_recipients_count
    
    if bcc_per_batch <= 0:
        raise ValueError(
            f"To ({to_count}) + CC ({cc_count}) recipients already exceed or equal the limit of {MAX_RECIPIENTS}. "
            f"Cannot add BCC recipients. Please reduce To/CC recipients."
        )
    
    # Split BCC recipients into batches
    batches = []
    bcc_list = bcc or []
    
    for i in range(0, len(bcc_list), bcc_per_batch):
        batch_bcc = bcc_list[i:i + bcc_per_batch]
        batches.append({
            "to": to if to else None,  # Same To in every batch
            "cc": cc if cc else None,  # Same CC in every batch
            "bcc": batch_bcc
        })
    
    # Send each batch
    for idx, batch in enumerate(batches, 1):
        batch_to_count = len(batch["to"]) if batch["to"] else 0
        batch_cc_count = len(batch["cc"]) if batch["cc"] else 0
        batch_bcc_count = len(batch["bcc"]) if batch["bcc"] else 0
        batch_total = batch_to_count + batch_cc_count + batch_bcc_count
        
        try:
            send_email(
                to=batch["to"],
                subject=subject,
                body=full_body,
                cc=batch["cc"],
                bcc=batch["bcc"],
                body_type=body_type,
                importance=importance,
                token=token
            )
            sent_count += batch_total
            print(f"✓ Batch {idx}/{len(batches)} sent successfully ({batch_total} recipients)")
        except Exception as e:
            failed_count += batch_total
            error_msg = f"Batch {idx} ({batch_total} recipients): {str(e)}"
            errors.append(error_msg)
            print(f"✗ Batch {idx}/{len(batches)} failed: {error_msg}")
    
    # Prepare response
    response = {
        "success": True,
        "message": f"Reply sending completed in {len(batches)} batches",
        "total_recipients": total_recipients,
        "total_batches": len(batches),
        "sent_count": sent_count,
        "failed_count": failed_count,
        "batch_size": MAX_RECIPIENTS,
    }
    
    if errors:
        response["errors"] = errors
    
    return response


def get_my_email(token: str = None) -> str:
    """Get current user's email address from cached token data."""
    from config import TOKEN_CACHE_FILE
    
    # Read from token cache file to avoid API call
    try:
        if TOKEN_CACHE_FILE.exists():
            with open(TOKEN_CACHE_FILE, "r", encoding='utf-8') as f:
                token_data = json.load(f)
                username = token_data.get("username")
                if username:
                    return username
    except Exception:
        pass
    
    # Fallback: call API
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me"
    response = api_request('get', url, token)
    
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
    to_recipients = [{"emailAddress": {"address": email}} for email in (to or [])]
    cc_recipients = [{"emailAddress": {"address": email}} for email in (cc or [])]
    bcc_recipients = [{"emailAddress": {"address": email}} for email in (bcc or [])]
    
    # Convert plain text comment to HTML if needed
    # Graph API forward expects HTML in comment field
    if comment:
        # First, convert literal \n strings to actual newlines (for CLI convenience)
        comment = comment.replace('\\n', '\n')
        
        if body_type == "html" and not comment.strip().startswith('<'):
            # Convert plain text to HTML with proper line breaks
            comment = comment.replace('\n', '<br>')
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
    response = api_request('post', url, token, json=payload)
    
    return True


def batch_forward_email(
    message_id: str,
    to: List[str] = None,
    cc: List[str] = None,
    bcc: List[str] = None,
    csv_path: str = None,
    email_column: str = None,
    comment: str = "",
    body_type: str = "html",
    token: str = None
) -> Dict[str, Any]:
    """
    Forward an email to multiple recipients with automatic batching.
    
    When total recipients exceed MAX_RECIPIENTS_PER_EMAIL (500), 
    automatically splits them into multiple batches and sends multiple emails.
    
    BCC recipients can be loaded from a CSV file for mass mailing.
    
    Args:
        message_id: ID of message to forward
        to: List of To recipient emails
        cc: List of CC recipient emails
        bcc: List of BCC recipient emails
        csv_path: Path to CSV file containing BCC email addresses
        email_column: Column name in CSV for email addresses (auto-detect if not specified)
        comment: Optional comment to add at the top
        body_type: "html" or "text"
        token: Access token
    
    Returns:
        Dictionary with batch processing results:
        {
            "success": True,
            "message": "Email forwarding completed in X batches",
            "total_recipients": total,
            "total_batches": num_batches,
            "sent_count": sent,
            "failed_count": failed,
            "batch_size": MAX_RECIPIENTS_PER_EMAIL,
            "errors": []  # List of error messages if any
        }
    """
    if token is None:
        token = get_access_token()
    
    # Get BCC recipients from CSV if provided
    if csv_path:
        print(f"📖 Reading BCC recipients from CSV: {csv_path}")
        csv_bcc = read_recipients_from_csv(csv_path, email_column)
        print(f"✓ Found {len(csv_bcc)} BCC recipients in CSV")
        # Merge with any manually specified BCC
        bcc = (bcc or []) + csv_bcc
    
    # Microsoft Graph API requires at least one To recipient for forward
    # If no To recipient specified but has BCC/CC, add current user as To
    if not to and (bcc or cc):
        to = [get_my_email(token)]
        print(f"ℹ️  Graph API requires To recipient - using current user: {to[0]}")
        print(f"💡 All other recipients will be in BCC (hidden from each other)")
    
    # Calculate total recipients
    to_count = len(to) if to else 0
    cc_count = len(cc) if cc else 0
    bcc_count = len(bcc) if bcc else 0
    total_recipients = to_count + cc_count + bcc_count
    
    # If within limit, use regular forward_email
    if total_recipients <= MAX_RECIPIENTS:
        forward_email(message_id, to, cc, bcc, comment, body_type, True, token)
        return {
            "success": True,
            "message": "Email forwarded successfully",
            "total_recipients": total_recipients,
            "total_batches": 1,
            "sent_count": total_recipients,
            "failed_count": 0,
            "batch_size": MAX_RECIPIENTS,
            "errors": []
        }
    
    # Split recipients into batches
    # IMPORTANT: To and CC recipients must be in EVERY batch (they're not distributed)
    # Only BCC recipients are distributed across batches
    # This ensures we stay under the 500 recipient limit per batch
    sent_count = 0
    failed_count = 0
    errors = []
    
    # Calculate how many BCC recipients can fit in each batch
    # Reserve space for To and CC recipients that will be in every batch
    fixed_recipients_count = to_count + cc_count
    bcc_per_batch = MAX_RECIPIENTS - fixed_recipients_count
    
    if bcc_per_batch <= 0:
        raise ValueError(
            f"To ({to_count}) + CC ({cc_count}) recipients already exceed or equal the limit of {MAX_RECIPIENTS}. "
            f"Cannot add BCC recipients. Please reduce To/CC recipients."
        )
    
    # Split BCC recipients into batches
    batches = []
    bcc_list = bcc or []
    
    for i in range(0, len(bcc_list), bcc_per_batch):
        batch_bcc = bcc_list[i:i + bcc_per_batch]
        batches.append({
            "to": to if to else None,  # Same To in every batch
            "cc": cc if cc else None,  # Same CC in every batch
            "bcc": batch_bcc
        })
    
    # Send each batch
    for idx, batch in enumerate(batches, 1):
        batch_to_count = len(batch["to"]) if batch["to"] else 0
        batch_cc_count = len(batch["cc"]) if batch["cc"] else 0
        batch_bcc_count = len(batch["bcc"]) if batch["bcc"] else 0
        batch_total = batch_to_count + batch_cc_count + batch_bcc_count
        
        try:
            forward_email(
                message_id=message_id,
                to=batch["to"],
                cc=batch["cc"],
                bcc=batch["bcc"],
                comment=comment,
                body_type=body_type,
                token=token
            )
            sent_count += batch_total
            print(f"✓ Batch {idx}/{len(batches)} sent successfully ({batch_total} recipients)")
        except Exception as e:
            failed_count += batch_total
            error_msg = f"Batch {idx} ({batch_total} recipients): {str(e)}"
            errors.append(error_msg)
            print(f"✗ Batch {idx}/{len(batches)} failed: {error_msg}")
    
    # Prepare response
    response = {
        "success": True,
        "message": f"Email forwarding completed in {len(batches)} batches",
        "total_recipients": total_recipients,
        "total_batches": len(batches),
        "sent_count": sent_count,
        "failed_count": failed_count,
        "batch_size": MAX_RECIPIENTS,
    }
    
    if errors:
        response["errors"] = errors
    
    return response


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
    
    response = api_request('patch', url, token, json=payload)
    
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
    response = api_request('get', url, token)
    
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
    
    response = api_request('get', url, token, params=params)
    
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
# ATTACHMENT OPERATIONS
# =============================================================================

def list_attachments(message_id: str, token: str = None) -> List[Dict[str, Any]]:
    """
    List all attachments for a message.
    
    Args:
        message_id: Message ID
        token: Access token
    
    Returns:
        List of attachment objects
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}/attachments"
    
    response = api_request('get', url, token)
    
    data = response.json()
    attachments = data.get("value", [])
    
    return attachments


def download_attachment(
    message_id: str,
    attachment_id: str,
    save_dir: str = None,
    token: str = None
) -> Dict[str, Any]:
    """
    Download a specific attachment.
    
    Args:
        message_id: Message ID
        attachment_id: Attachment ID
        save_dir: Directory to save the file (default: current directory)
        token: Access token
    
    Returns:
        Dictionary with download result
    """
    if token is None:
        token = get_access_token()
    
    # Get attachment metadata
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}/attachments/{attachment_id}"
    response = api_request('get', url, token)
    
    attachment = response.json()
    
    # Get file name and content
    file_name = attachment.get('name', 'attachment')
    content_bytes = attachment.get('contentBytes')
    content_type = attachment.get('contentType', 'application/octet-stream')
    
    if not content_bytes:
        # For file attachments, content is in contentBytes
        # For item attachments (like emails), we need to handle differently
        if attachment.get('@odata.type') == '#microsoft.graph.itemAttachment':
            raise Exception("Item attachments (embedded emails) are not supported for direct download")
        raise Exception("No content found in attachment")
    
    # Decode base64 content
    import base64
    file_content = base64.b64decode(content_bytes)
    
    # Determine save path
    if save_dir:
        save_path = Path(save_dir)
        save_path.mkdir(parents=True, exist_ok=True)
    else:
        save_path = Path.cwd()
    
    file_path = save_path / file_name
    
    # Write file
    with open(file_path, 'wb') as f:
        f.write(file_content)
    
    return {
        "success": True,
        "file_name": file_name,
        "file_path": str(file_path),
        "size": len(file_content),
        "content_type": content_type
    }


def download_all_attachments(
    message_id: str,
    save_dir: str = None,
    token: str = None
) -> List[Dict[str, Any]]:
    """
    Download all attachments from a message.
    
    Args:
        message_id: Message ID
        save_dir: Directory to save files (default: current directory)
        token: Access token
    
    Returns:
        List of download results
    """
    if token is None:
        token = get_access_token()
    
    # Get all attachments
    attachments = list_attachments(message_id, token)
    
    if not attachments:
        return []
    
    results = []
    for attachment in attachments:
        attachment_id = attachment.get('id')
        try:
            result = download_attachment(message_id, attachment_id, save_dir, token)
            results.append(result)
        except Exception as e:
            results.append({
                "success": False,
                "file_name": attachment.get('name', 'unknown'),
                "error": str(e)
            })
    
    return results


def display_attachments(attachments: List[Dict]):
    """Display a list of attachments."""
    if not attachments:
        print("No attachments found.")
        return
    
    print(f"\n{'='*80}")
    print(f"{'#':<5} {'Name':<40} {'Size':<15} {'Type':<20}")
    print(f"{'='*80}")
    
    for i, att in enumerate(attachments, 1):
        name = att.get('name', 'Unknown')[:40]
        size = att.get('size', 0)
        if size:
            if size < 1024:
                size_str = f"{size} B"
            elif size < 1024 * 1024:
                size_str = f"{size / 1024:.1f} KB"
            else:
                size_str = f"{size / (1024 * 1024):.1f} MB"
        else:
            size_str = "-"
        
        content_type = att.get('contentType', 'Unknown')[:20]
        att_type = att.get('@odata.type', '').replace('#microsoft.graph.', '')
        
        print(f"{i:<5} {name:<40} {size_str:<15} {content_type:<20}")
    
    print(f"{'='*80}")
    print(f"Total: {len(attachments)} attachments")

def prepare_file_attachments(file_paths: List[str]) -> tuple[List[Dict[str, Any]], List[str]]:
    """
    Prepare file attachments for sending via Microsoft Graph API.
    
    Supports files up to 150 MB (Microsoft Graph limit).
    Files under 3 MB use inline attachment.
    Files 3-150 MB require draft message + upload session.
    
    Args:
        file_paths: List of file paths to attach
    
    Returns:
        Tuple of (inline_attachments, large_file_paths)
        - inline_attachments: List of attachment objects for files < 3 MB
        - large_file_paths: List of file paths for files >= 3 MB (need upload session)
    
    Raises:
        FileNotFoundError: If any file doesn't exist
        Exception: If file is too large or can't be read
    """
    import base64
    import mimetypes
    
    inline_attachments = []
    large_file_paths = []
    max_inline_size = 3 * 1024 * 1024  # 3 MB limit for inline attachments
    max_total_size = 150 * 1024 * 1024  # 150 MB total limit per message
    
    for file_path in file_paths:
        path = Path(file_path)
        
        if not path.exists():
            raise FileNotFoundError(f"Attachment file not found: {file_path}")
        
        file_size = path.stat().st_size
        
        # Check absolute maximum
        if file_size > max_total_size:
            raise Exception(
                f"File '{path.name}' is too large ({file_size / (1024*1024):.1f} MB). "
                f"Maximum size per file is 150 MB."
            )
        
        # For files > 3 MB, we need to use upload session (create draft + upload)
        if file_size > max_inline_size:
            large_file_paths.append(file_path)
            print(f"📎 Large file detected: {path.name} ({file_size / (1024*1024):.1f} MB) - will use upload session")
            continue
        
        # Read file content for inline attachment
        try:
            with open(path, 'rb') as f:
                file_content = f.read()
        except Exception as e:
            raise Exception(f"Failed to read file '{file_path}': {str(e)}")
        
        # Encode to base64
        content_bytes = base64.b64encode(file_content).decode('utf-8')
        
        # Detect MIME type
        mime_type, _ = mimetypes.guess_type(str(path))
        if not mime_type:
            mime_type = 'application/octet-stream'
        
        # Create attachment object
        attachment = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": path.name,
            "contentType": mime_type,
            "contentBytes": content_bytes
        }
        
        inline_attachments.append(attachment)
        print(f"📎 Inline attachment prepared: {path.name} ({file_size / 1024:.1f} KB)")
    
    return inline_attachments, large_file_paths


def create_draft_message(
    to: List[str],
    subject: str,
    body: str,
    cc: List[str] = None,
    bcc: List[str] = None,
    body_type: str = "html",
    importance: str = None,
    token: str = None
) -> str:
    """
    Create a draft message and return its ID.
    
    Args:
        to: List of To recipient emails
        subject: Email subject
        body: Email body content
        cc: List of CC recipient emails
        bcc: List of BCC recipient emails
        body_type: "html" or "text"
        importance: "low", "normal", or "high"
        token: Access token
    
    Returns:
        str: Draft message ID
    """
    if token is None:
        token = get_access_token()
    
    # Convert plain text body to HTML if needed
    if body_type == "html" and not body.strip().startswith('<'):
        body = body.replace('\n', '<br>')
    
    # Build message payload
    message = {
        "subject": subject,
        "body": {
            "contentType": body_type,
            "content": body
        },
        "toRecipients": [format_email_address(e) for e in (to or [])],
        "ccRecipients": [format_email_address(e) for e in (cc or [])],
        "bccRecipients": [format_email_address(e) for e in (bcc or [])]
    }
    
    # Add importance if specified
    if importance and importance.lower() in ['low', 'normal', 'high']:
        message["importance"] = importance.lower()
    
    url = f"{GRAPH_API_BASE}/me/messages"
    
    response = api_request('post', url, token, json=message)
    draft = response.json()
    
    return draft['id']


def upload_large_attachment(
    message_id: str,
    file_path: str,
    token: str = None
) -> bool:
    """
    Upload a large attachment (>3MB) to a draft message using upload session.
    
    Args:
        message_id: Draft message ID
        file_path: Path to file to attach
        token: Access token
    
    Returns:
        bool: True if successful
    """
    import mimetypes
    
    if token is None:
        token = get_access_token()
    
    path = Path(file_path)
    
    if not path.exists():
        raise FileNotFoundError(f"Attachment file not found: {file_path}")
    
    file_size = path.stat().st_size
    file_name = path.name
    
    # Detect MIME type
    mime_type, _ = mimetypes.guess_type(str(path))
    if not mime_type:
        mime_type = 'application/octet-stream'
    
    # Create upload session
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}/attachments/createUploadSession"
    
    attachment_item = {
        "attachmentType": "file",
        "name": file_name,
        "size": file_size
    }
    
    payload = {
        "AttachmentItem": attachment_item
    }
    
    response = api_request('post', url, token, json=payload)
    upload_session = response.json()
    upload_url = upload_session['uploadUrl']
    
    # Upload file in chunks (recommended chunk size: 5-10 MB)
    chunk_size = 5 * 1024 * 1024  # 5 MB chunks
    
    with open(path, 'rb') as f:
        file_data = f.read()
    
    bytes_uploaded = 0
    total_bytes = len(file_data)
    
    print(f"📤 Uploading {file_name} ({total_bytes / (1024*1024):.1f} MB)...")
    
    while bytes_uploaded < total_bytes:
        chunk_end = min(bytes_uploaded + chunk_size, total_bytes)
        chunk = file_data[bytes_uploaded:chunk_end]
        
        # Prepare headers for this chunk
        headers = {
            'Content-Length': str(len(chunk)),
            'Content-Range': f'bytes {bytes_uploaded}-{chunk_end-1}/{total_bytes}',
            'Content-Type': 'application/octet-stream'
        }
        
        # Upload chunk (don't use api_request as it adds auth headers we don't need for upload URL)
        response = requests.put(upload_url, headers=headers, data=chunk)
        
        if response.status_code not in [200, 201, 202]:
            raise Exception(f"Upload failed: {response.status_code} - {response.text}")
        
        bytes_uploaded = chunk_end
        progress = (bytes_uploaded / total_bytes) * 100
        print(f"  Progress: {progress:.1f}% ({bytes_uploaded / (1024*1024):.1f} MB / {total_bytes / (1024*1024):.1f} MB)")
    
    print(f"✓ Upload complete: {file_name}")
    return True


def send_draft_message(message_id: str, token: str = None) -> bool:
    """
    Send a draft message.
    
    Args:
        message_id: Draft message ID
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/messages/{message_id}/send"
    
    response = api_request('post', url, token)
    
    return True



# =============================================================================
# LIST MAIL FOLDERS
# =============================================================================

def list_mail_folders(
    include_hidden: bool = False,
    token: str = None
) -> List[Dict[str, Any]]:
    """
    List all mail folders for the user.
    
    Args:
        include_hidden: Include hidden folders
        token: Access token
    
    Returns:
        List of folder objects with id, displayName, totalItemCount, unreadItemCount
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/mailFolders"
    
    params = {
        "$select": "id,displayName,totalItemCount,unreadItemCount,isHidden,childFolderCount"
    }
    
    response = api_request('get', url, token, params=params)
    
    data = response.json()
    folders = data.get("value", [])
    
    # Filter out hidden folders if not requested
    if not include_hidden:
        folders = [f for f in folders if not f.get('isHidden', False)]
    
    return folders


def display_folder_list(folders: List[Dict]):
    """Display a list of mail folders."""
    print(f"\n{'='*100}")
    print(f"{'Folder Name':<40} {'Total':<10} {'Unread':<10} {'ID':<40}")
    print(f"{'='*100}")
    
    for folder in folders:
        name = folder.get('displayName', 'Unknown')[:40]
        total = folder.get('totalItemCount', 0)
        unread = folder.get('unreadItemCount', 0)
        folder_id = folder.get('id', '')[:40]
        
        unread_str = f"[{unread}]" if unread > 0 else ""
        print(f"{name:<40} {total:<10} {unread_str:<10} {folder_id:<40}")
    
    print(f"{'='*100}")
    print(f"Total: {len(folders)} folders")
    print(f"\n💡 Tip: Use --folder <name> or --folder <ID> to search in a specific folder")
    print(f"💡 Tip: Use --folder all to search across all folders")


# =============================================================================
# DISPLAY HELPERS
# =============================================================================

def display_message_list(messages: List[Dict], show_preview: bool = True, show_detail: bool = False, display_timezone: str = None, message_type: str = "all"):
    """Display a list of messages in a readable format.
    
    Args:
        messages: List of message dictionaries
        show_preview: Show bodyPreview (first few lines)
        show_detail: Show full body content (requires --detail flag to fetch full messages)
        display_timezone: Timezone for display (e.g., "+08:00" or "UTC")
        message_type: Message type filter - "all" (show categorized), "emails", or "events"
    """
    # If message_type is "all", categorize messages into emails and events
    # Meeting invites are identified by @odata.type containing "eventMessage"
    if message_type == "all":
        emails = [m for m in messages if 'eventMessage' not in m.get('@odata.type', '')]
        events = [m for m in messages if 'eventMessage' in m.get('@odata.type', '')]
        
        # Display emails first
        if emails:
            print(f"\n{'='*80}")
            print(f"📧 EMAILS ({len(emails)} messages)")
            print(f"{'='*80}")
            display_message_list(emails, show_preview, show_detail, display_timezone, "emails")
        
        # Display events second
        if events:
            print(f"\n{'='*80}")
            print(f"📅 MEETING INVITES ({len(events)} messages)")
            print(f"{'='*80}")
            display_message_list(events, show_preview, show_detail, display_timezone, "events")
        
        # Display total
        print(f"\n{'='*80}")
        print(f"Total: {len(messages)} messages ({len(emails)} emails, {len(events)} events)")
        return
    
    # Determine display timezone
    display_tz = None
    if display_timezone:
        try:
            # Try to use timezone name directly (e.g., "Asia/Shanghai")
            display_tz = ZoneInfo(display_timezone)
        except Exception:
            try:
                # Parse timezone offset like "+08:00" or "-05:00"
                import re
                match = re.match(r'^([+-])(\d{2}):?(\d{2})?$', display_timezone)
                if match:
                    sign = match.group(1)
                    hours = int(match.group(2))
                    # ZoneInfo uses Etc/GMT convention (inverted sign)
                    if sign == '+':
                        display_tz = ZoneInfo(f"Etc/GMT-{hours}")
                    else:
                        display_tz = ZoneInfo(f"Etc/GMT+{hours}")
            except Exception:
                pass  # Fallback to UTC display
    
    # If show_detail is True, show full body content
    if show_detail:
        print(f"\n{'='*80}")
        for i, msg in enumerate(messages, 1):
            received = msg.get('receivedDateTime', '')
            if received:
                dt = datetime.fromisoformat(received.replace('Z', '+00:00'))
                if display_tz:
                    dt = dt.astimezone(display_tz)
                # Simple clean format without timezone suffix
                received = dt.strftime('%Y-%m-%d %H:%M')

            from_addr = msg.get('from', {}).get('emailAddress', {})
            sender = from_addr.get('name', from_addr.get('address', 'Unknown'))

            subject = msg.get('subject', '(No Subject)')
            read_status = '[UNREAD] ' if not msg.get('isRead', True) else ''
            message_id = msg.get('id', '')

            print(f"\n{i}. {read_status}{subject}")
            print(f"   From: {sender}")
            print(f"   Received: {received}")
            print(f"   ID: {message_id}")
            
            # Show attachments
            attachments = msg.get('attachments', [])
            if attachments:
                # Separate regular attachments from embedded images
                regular_attachments = [att for att in attachments if not att.get('isInline', False)]
                embedded_images = [att for att in attachments if att.get('isInline', False)]
                
                if regular_attachments:
                    att_names = [att.get('name', 'Unknown') for att in regular_attachments]
                    try:
                        print(f"   📎 Attachments ({len(regular_attachments)}): {', '.join(att_names)}")
                    except:
                        print(f"   [Attachments] ({len(regular_attachments)}): {', '.join(att_names)}")
                
                if embedded_images:
                    img_names = [att.get('name', 'Unknown') for att in embedded_images]
                    try:
                        print(f"   🖼️  Embedded images ({len(embedded_images)}): {', '.join(img_names)}")
                    except:
                        print(f"   [Embedded Images] ({len(embedded_images)}): {', '.join(img_names)}")
            
            # Show To recipients
            to_recipients = msg.get('toRecipients', [])
            if to_recipients:
                to_names = [r.get('emailAddress', {}).get('name', r.get('emailAddress', {}).get('address', '')) for r in to_recipients]
                to_display = ', '.join(to_names[:3])
                if len(to_recipients) > 3:
                    to_display += f' (+{len(to_recipients) - 3} more)'
                print(f"   To: {to_display}")
            
            # Show CC recipients
            cc_recipients = msg.get('ccRecipients', [])
            if cc_recipients:
                cc_names = [r.get('emailAddress', {}).get('name', r.get('emailAddress', {}).get('address', '')) for r in cc_recipients]
                cc_display = ', '.join(cc_names[:3])
                if len(cc_recipients) > 3:
                    cc_display += f' (+{len(cc_recipients) - 3} more)'
                print(f"   Cc: {cc_display}")
            
            # Show full body content
            body_content = msg.get('body', {})
            body_text = body_content.get('content', '') if body_content else ''
            if body_text:
                # Clean up HTML if needed
                if body_content.get('contentType') == 'html':
                    # Remove script and style tags with their content
                    import re
                    body_text = re.sub(r'<(script|style)[^>]*>.*?</\1>', '', body_text, flags=re.DOTALL | re.IGNORECASE)
                    # Remove HTML tags
                    body_text = re.sub(r'<[^>]+>', '', body_text)
                    # Decode HTML entities
                    body_text = body_text.replace('&nbsp;', ' ').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
                    body_text = body_text.replace('&quot;', '"').replace('&#39;', "'").replace('&apos;', "'")
                    # Clean up whitespace
                    body_text = re.sub(r'[ \t]+', ' ', body_text)  # Multiple spaces to single
                    body_text = re.sub(r'\n\s*\n\s*\n+', '\n\n', body_text)  # Multiple blank lines to double
                    body_text = body_text.strip()
                # Limit display to 2000 chars for readability
                if len(body_text) > 2000:
                    body_text = body_text[:2000] + '\n... [truncated, use --json for full content]'
                print(f"\n   Body:\n   {'-'*40}")
                for line in body_text.split('\n'):
                    print(f"   {line}")
                print(f"   {'-'*40}")
            
            if i < len(messages):
                print(f"\n{'~'*80}")
        
        print(f"\n{'='*80}")
        print(f"Total: {len(messages)} messages")
        return
    
    if not show_preview:
        # Compact table format without preview
        print(f"\n{'='*120}")
        print(f"{'Date':<20} {'From':<25} {'Subject':<30} {'Attachments':<20} {'ID':<15}")
        print(f"{'='*120}")

        for msg in messages:
            received = msg.get('receivedDateTime', '')
            if received:
                dt = datetime.fromisoformat(received.replace('Z', '+00:00'))
                if display_tz:
                    dt = dt.astimezone(display_tz)
                # Simple clean format without timezone suffix
                received = dt.strftime('%Y-%m-%d %H:%M')

            from_addr = msg.get('from', {}).get('emailAddress', {})
            sender = from_addr.get('name', from_addr.get('address', 'Unknown'))

            subject = msg.get('subject', '(No Subject)')[:30]
            read_status = '' if msg.get('isRead', True) else '[UNREAD]'
            message_id = msg.get('id', '')[:15]
            
            # Format attachment info
            attachments = msg.get('attachments', [])
            att_info = ''
            if attachments:
                regular_attachments = [att for att in attachments if not att.get('isInline', False)]
                embedded_images = [att for att in attachments if att.get('isInline', False)]
                parts = []
                if regular_attachments:
                    parts.append(f"📎{len(regular_attachments)}")
                if embedded_images:
                    parts.append(f"🖼️{len(embedded_images)}")
                att_info = ' '.join(parts)

            print(f"{received:<20} {sender:<25} {subject}{read_status:<{30 - len(read_status)}} {att_info:<20} {message_id:<15}")

        print(f"{'='*120}")
        print(f"Total: {len(messages)} messages")
        print(f"💡 Tip: Use 'get <ID>' to view full email content")
    else:
        # Detailed format with preview
        print(f"\n{'='*80}")
        for i, msg in enumerate(messages, 1):
            received = msg.get('receivedDateTime', '')
            if received:
                dt = datetime.fromisoformat(received.replace('Z', '+00:00'))
                if display_tz:
                    dt = dt.astimezone(display_tz)
                # Simple clean format without timezone suffix
                # Timezone is already shown at the top, no need to repeat
                received = dt.strftime('%Y-%m-%d %H:%M')

            from_addr = msg.get('from', {}).get('emailAddress', {})
            sender = from_addr.get('name', from_addr.get('address', 'Unknown'))

            subject = msg.get('subject', '(No Subject)')
            read_status = '[UNREAD] ' if not msg.get('isRead', True) else ''
            message_id = msg.get('id', '')

            print(f"\n{i}. {read_status}{subject}")
            print(f"   From: {sender}")
            print(f"   Received: {received}")
            print(f"   ID: {message_id}")
            
            # Show response status and meeting time for meeting invites
            if 'eventMessage' in msg.get('@odata.type', ''):
                # Try to get event data - it might be in the message or need to be fetched
                event = msg.get('event')
                if not event:
                    # Event data not expanded, try to fetch it
                    try:
                        event = get_event_from_message(message_id, token=None)
                    except:
                        event = None
                
                if event:
                    response_status = event.get('responseStatus', {})
                    response = response_status.get('response', 'none')
                    # Map response to emoji and text
                    status_map = {
                        'accepted': '✅ Accepted',
                        'declined': '❌ Declined',
                        'tentativelyAccepted': '❓ Tentative',
                        'notResponded': '⏳ Not Responded',
                        'organizer': '👤 Organizer',
                        'none': '⏳ Not Responded'
                    }
                    status_display = status_map.get(response, f'❔ {response}')
                    print(f"   Status: {status_display}")
                    
                    # Only show meeting time and availability for meetings not yet responded to
                    # For accepted/declined/tentative meetings, just show the status
                    if response in ['notResponded', 'none']:
                        # Show meeting time and availability
                        start_info = event.get('start', {})
                        end_info = event.get('end', {})
                    else:
                        start_info = None
                        end_info = None
                    
                    if start_info and end_info:
                        start_dt_str = start_info.get('dateTime', '')
                        end_dt_str = end_info.get('dateTime', '')
                        start_tz = start_info.get('timeZone', 'UTC')
                        
                        if start_dt_str and end_dt_str:
                            try:
                                # Parse datetime (handle Microsoft's format with extra precision)
                                start_str = start_dt_str.split('.')[0]
                                end_str = end_dt_str.split('.')[0]
                                
                                start_dt = datetime.fromisoformat(start_str).replace(tzinfo=ZoneInfo('UTC'))
                                end_dt = datetime.fromisoformat(end_str).replace(tzinfo=ZoneInfo('UTC'))
                                
                                # Convert to user's timezone if display_tz is set, otherwise use UTC+8
                                user_tz = display_tz if display_tz else ZoneInfo('Asia/Shanghai')
                                start_local = start_dt.astimezone(user_tz)
                                end_local = end_dt.astimezone(user_tz)
                                
                                duration_mins = (end_dt - start_dt).total_seconds() / 60
                                
                                # Check availability
                                # For accepted meetings: show actual busy status (don't exclude itself)
                                # For not responded meetings: exclude itself to check for conflicts
                                availability_status = "❓ Unknown"
                                recommendation = ""
                                
                                # Only check availability for meetings not yet responded to
                                if response in ['notResponded', 'none']:
                                    try:
                                        # Import calendar operations to check availability
                                        import sys
                                        from pathlib import Path
                                        sys.path.insert(0, str(Path(__file__).parent))
                                        from calendar_operations import get_availability
                                        
                                        # Get current user's email
                                        my_email = get_my_email()
                                        
                                        # Check if user is free during this time
                                        # Convert datetime to ISO format for API
                                        avail_result = get_availability(
                                            emails=[my_email],
                                            start=start_dt.isoformat(),
                                            end=end_dt.isoformat(),
                                            timezone='UTC'
                                        )
                                        
                                        # Parse availability - check schedule items to see if there are OTHER meetings
                                        schedules = avail_result.get('value', [])
                                        if schedules:
                                            schedule = schedules[0]
                                            schedule_items = schedule.get('scheduleItems', [])
                                            
                                            # Get this event's subject - try event first, then fall back to message subject
                                            current_subject = event.get('subject', '') or msg.get('subject', '')
                                            current_subject = current_subject.strip()
                                            
                                            # Check if there are any OTHER events (excluding this one by subject)
                                            other_events = []
                                            for item in schedule_items:
                                                item_subject = item.get('subject', '').strip()
                                                
                                                # Skip if this is the same event (match by subject)
                                                # Use case-insensitive comparison
                                                if item_subject.lower() == current_subject.lower():
                                                    continue
                                                
                                                other_events.append(item)
                                            
                                            if not other_events:
                                                availability_status = "✅ Free"
                                                recommendation = " → Recommend Accept"
                                            else:
                                                availability_status = f"❌ Busy ({len(other_events)} conflict{'s' if len(other_events) > 1 else ''})"
                                                recommendation = " → Not Recommended"
                                    except Exception as e:
                                        # If availability check fails, just show unknown
                                        pass
                                
                                print(f"   Time: {start_local.strftime('%Y-%m-%d %H:%M')} - {end_local.strftime('%H:%M')} ({duration_mins:.0f} min) {availability_status}{recommendation}")
                            except Exception:
                                pass  # If parsing fails, just don't show time
            
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
            
            # Show attachments
            attachments = msg.get('attachments', [])
            if attachments:
                # Separate regular attachments from embedded images
                regular_attachments = [att for att in attachments if not att.get('isInline', False)]
                embedded_images = [att for att in attachments if att.get('isInline', False)]
                
                if regular_attachments:
                    att_names = [att.get('name', 'Unknown') for att in regular_attachments]
                    try:
                        print(f"   📎 Attachments ({len(regular_attachments)}): {', '.join(att_names)}")
                    except:
                        print(f"   [Attachments] ({len(regular_attachments)}): {', '.join(att_names)}")
                
                if embedded_images:
                    img_names = [att.get('name', 'Unknown') for att in embedded_images]
                    try:
                        print(f"   🖼️  Embedded images ({len(embedded_images)}): {', '.join(img_names)}")
                    except:
                        print(f"   [Embedded Images] ({len(embedded_images)}): {', '.join(img_names)}")
            
            if i < len(messages):
                print(f"   {'-'*78}")
        
        print(f"\n{'='*80}")
        print(f"Total: {len(messages)} messages")


def display_message(message: Dict):
    """Display a single message in detail with thread separation."""
    import re
    import html as html_module

    print(f"\n{'='*80}")
    print(f"Subject: {message.get('subject', '(No Subject)')}")
    print(f"From: {message.get('from', {}).get('emailAddress', {})}")
    print(f"To: {[r.get('emailAddress', {}) for r in message.get('toRecipients', [])]}")
    print(f"CC: {[r.get('emailAddress', {}) for r in message.get('ccRecipients', [])]}")
    print(f"Date: {message.get('receivedDateTime', '')}")
    print(f"ID: {message.get('id', '')}")
    
    # Show attachments
    attachments = message.get('attachments', [])
    if attachments:
        regular_attachments = [att for att in attachments if not att.get('isInline', False)]
        embedded_images = [att for att in attachments if att.get('isInline', False)]
        
        if regular_attachments:
            print(f"📎 Attachments ({len(regular_attachments)}):")
            for att in regular_attachments:
                att_id = att.get('id', 'N/A')
                name = att.get('name', 'Unknown')
                size = att.get('size', 0)
                size_kb = size / 1024 if size else 0
                content_type = att.get('contentType', 'Unknown')
                print(f"   - {name} ({size_kb:.1f} KB, {content_type})")
                print(f"     ID: {att_id}")
        
        if embedded_images:
            print(f"🖼️  Embedded Images ({len(embedded_images)}):")
            for att in embedded_images:
                att_id = att.get('id', 'N/A')
                name = att.get('name', 'Unknown')
                size = att.get('size', 0)
                size_kb = size / 1024 if size else 0
                content_type = att.get('contentType', 'Unknown')
                print(f"   - {name} ({size_kb:.1f} KB, {content_type})")
                print(f"     ID: {att_id}")
    
    print(f"{'='*80}")
    
    # Get body content
    body = message.get('body', {})
    content = body.get('content', '')
    content_type = body.get('contentType', 'text')
    
    # If HTML, extract and parse thread structure
    if content_type == 'html' and content:
        # Remove style and script tags first
        content = re.sub(r'<style[^>]*>.*?</style>', '', content, flags=re.DOTALL | re.IGNORECASE)
        content = re.sub(r'<script[^>]*>.*?</script>', '', content, flags=re.DOTALL | re.IGNORECASE)
        
        # Clean up HTML tags FIRST to make pattern matching easier
        clean_content = re.sub(r'<[^>]+>', '\n', content)  # Replace tags with newlines to preserve structure
        clean_content = html_module.unescape(clean_content)
        # Normalize whitespace but keep newlines
        clean_content = re.sub(r'[ \t]+', ' ', clean_content)
        clean_content = re.sub(r'\n\s*\n\s*\n+', '\n\n', clean_content)
        
        # Now look for email thread headers in the cleaned text
        # Pattern: "From: Name <email> Sent: date" (common in forwards/replies)
        thread_header_pattern = r'From:\s+[^<\n]+<[^>]+>\s+Sent:\s+[^\n]+'
        
        # Find all thread boundaries
        matches = list(re.finditer(thread_header_pattern, clean_content, re.IGNORECASE))
        
        threads = []
        if matches:
            # Split content by thread headers
            last_pos = 0
            for match in matches:
                # Get content before this thread header (the previous thread)
                if match.start() > last_pos:
                    threads.append(clean_content[last_pos:match.start()])
                # Mark the start of next thread (include the header)
                last_pos = match.start()
            # Add the last thread
            if last_pos < len(clean_content):
                threads.append(clean_content[last_pos:])
        else:
            # No thread headers found - single thread
            threads = [clean_content]
        
        # Display each thread
        for i, thread_content in enumerate(threads):
            if not thread_content or not thread_content.strip():
                continue
            
            # Final cleanup
            thread_content = thread_content.strip()
            
            if i > 0:
                # Add bread line separator between threads
                print(f"\n{'─'*80}")
                print(f"📧 Thread {i+1}")
                print(f"{'─'*80}")
            
            # Handle encoding issues
            try:
                print(f"\n{thread_content}")
            except UnicodeEncodeError:
                print(f"\n{thread_content.encode('ascii', 'replace').decode('ascii')}")
    else:
        # Plain text content - also look for thread patterns
        thread_header_pattern = r'(?:From|发件人):\s*[^\n]+\s+(?:Sent|发送时间):\s*[^\n]+'
        matches = list(re.finditer(thread_header_pattern, content, re.MULTILINE))
        
        if matches:
            threads = []
            last_pos = 0
            for match in matches:
                if match.start() > last_pos:
                    threads.append(content[last_pos:match.start()])
                last_pos = match.start()
            if last_pos < len(content):
                threads.append(content[last_pos:])
            
            # Display threads with separators
            for i, thread_content in enumerate(threads):
                if not thread_content.strip():
                    continue
                    
                if i > 0:
                    print(f"\n{'─'*80}")
                    print(f"📧 Thread {i+1}")
                    print(f"{'─'*80}")
                
                try:
                    print(f"\n{thread_content.strip()}")
                except UnicodeEncodeError:
                    print(f"\n{thread_content.strip().encode('ascii', 'replace').decode('ascii')}")
        else:
            # Single thread - show as is
            try:
                print(f"\n{content}")
            except UnicodeEncodeError:
                print(f"\n{content.encode('ascii', 'replace').decode('ascii')}")
    
    print(f"\n{'='*80}")


# =============================================================================
# CLI
# =============================================================================

def unescape_body(body: str) -> str:
    """
    Convert escape sequences in body text to actual characters.
    Handles: \\n -> newline, \\t -> tab, \\\\ -> backslash
    This allows using \\n in command line arguments across all shells.
    """
    if not body:
        return body
    return body.replace('\\n', '\n').replace('\\t', '\t').replace('\\\\', '\\')


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
    list_parser.add_argument("--detail", action="store_true", help="Show full email body content (alias for --preview with more detail)")
    list_parser.add_argument("--focused", action="store_true", help="Show Focused inbox only")
    list_parser.add_argument("--other", action="store_true", help="Show Other inbox only")
    # Message type filters
    list_parser.add_argument("--emails-only", action="store_true", help="Show only emails (exclude meeting invites)")
    list_parser.add_argument("--events-only", action="store_true", help="Show only meeting invites (exclude regular emails)")
    # Search parameters
    list_parser.add_argument("--from", dest="from_sender", help="Search by sender name or email")
    list_parser.add_argument("--to", dest="to_recipient", help="Search by recipient name or email")
    list_parser.add_argument("--subject", help="Search by subject text")
    list_parser.add_argument("--body", help="Search by body text")
    list_parser.add_argument("--start", help="Filter messages received after this timestamp (format: '2026-03-26T12:00:00+08:00', '2026-03-26T04:00:00Z', or 'now')")
    list_parser.add_argument("--end", help="Filter messages received before this timestamp (format: '2026-03-26T12:00:00+08:00', '2026-03-26T04:00:00Z', or 'now')")
    list_parser.add_argument("--timezone", required=True, help="Display timezone (e.g., 'Asia/Shanghai', 'UTC', '+08:00')")
    
    # Search command (complete alias for list, supports all parameters)
    search_parser = subparsers.add_parser("search", help="Search/list messages (alias for list)")
    search_parser.add_argument("--folder", default="inbox", help="Folder name (or 'all' for all folders)")
    search_parser.add_argument("--limit", "--top", type=int, default=25, dest="limit", help="Max messages to return (--top is alias)")
    search_parser.add_argument("--filter", dest="filter_query", help="OData filter query")
    search_parser.add_argument("--unread", action="store_true", help="Show unread only")
    search_parser.add_argument("--preview", action="store_true", help="Show email body preview")
    search_parser.add_argument("--detail", action="store_true", help="Show full email body content (alias for --preview with more detail)")
    search_parser.add_argument("--focused", action="store_true", help="Show Focused inbox only")
    search_parser.add_argument("--other", action="store_true", help="Show Other inbox only")
    # Message type filters
    search_parser.add_argument("--emails-only", action="store_true", help="Show only emails (exclude meeting invites)")
    search_parser.add_argument("--events-only", action="store_true", help="Show only meeting invites (exclude regular emails)")
    # Search parameters
    search_parser.add_argument("--from", dest="from_sender", help="Search by sender name or email")
    search_parser.add_argument("--to", dest="to_recipient", help="Search by recipient name or email")
    search_parser.add_argument("--subject", help="Search by subject text")
    search_parser.add_argument("--body", help="Search by body text")
    search_parser.add_argument("--start", help="Filter messages received after this timestamp (format: '2026-03-26T12:00:00+08:00', '2026-03-26T04:00:00Z', or 'now')")
    search_parser.add_argument("--end", help="Filter messages received before this timestamp (format: '2026-03-26T12:00:00+08:00', '2026-03-26T04:00:00Z', or 'now')")
    search_parser.add_argument("--timezone", required=True, help="Display timezone (e.g., 'Asia/Shanghai', 'UTC', '+08:00')")
    
    # Find command (complete alias for list, supports all parameters)
    find_parser = subparsers.add_parser("find", help="Find/list messages (alias for list)")
    find_parser.add_argument("--folder", default="inbox", help="Folder name (or 'all' for all folders)")
    find_parser.add_argument("--limit", "--top", type=int, default=25, dest="limit", help="Max messages to return (--top is alias)")
    find_parser.add_argument("--filter", dest="filter_query", help="OData filter query")
    find_parser.add_argument("--unread", action="store_true", help="Show unread only")
    find_parser.add_argument("--preview", action="store_true", help="Show email body preview")
    find_parser.add_argument("--detail", action="store_true", help="Show full email body content (alias for --preview with more detail)")
    find_parser.add_argument("--focused", action="store_true", help="Show Focused inbox only")
    find_parser.add_argument("--other", action="store_true", help="Show Other inbox only")
    # Message type filters
    find_parser.add_argument("--emails-only", action="store_true", help="Show only emails (exclude meeting invites)")
    find_parser.add_argument("--events-only", action="store_true", help="Show only meeting invites (exclude regular emails)")
    # Search parameters
    find_parser.add_argument("--from", dest="from_sender", help="Search by sender name or email")
    find_parser.add_argument("--to", dest="to_recipient", help="Search by recipient name or email")
    find_parser.add_argument("--subject", help="Search by subject text")
    find_parser.add_argument("--body", help="Search by body text")
    find_parser.add_argument("--start", help="Filter messages received after this timestamp (format: '2026-03-26T12:00:00+08:00', '2026-03-26T04:00:00Z', or 'now')")
    find_parser.add_argument("--end", help="Filter messages received before this timestamp (format: '2026-03-26T12:00:00+08:00', '2026-03-26T04:00:00Z', or 'now')")
    find_parser.add_argument("--timezone", required=True, help="Display timezone (e.g., 'Asia/Shanghai', 'UTC', '+08:00')")
    
    # Get command
    get_parser = subparsers.add_parser("get", help="Get a message")
    get_parser.add_argument("message_id", help="Message ID")
    
    # Thread command
    thread_parser = subparsers.add_parser("thread", help="Get conversation thread")
    thread_parser.add_argument("message_id", help="Message ID (any message in the thread)")
    
    # Send command (supports batching automatically)
    send_parser = subparsers.add_parser("send", help="Send an email (auto-batch for large recipient lists)")
    send_parser.add_argument("--to", action="append", help="To recipients (comma-separated or use multiple --to)")
    send_parser.add_argument("--cc", action="append", help="CC recipients (comma-separated or use multiple --cc)")
    send_parser.add_argument("--bcc", action="append", help="BCC recipients (comma-separated or use multiple --bcc)")
    send_parser.add_argument("--csv", dest="csv_path", help="CSV file containing BCC email addresses")
    send_parser.add_argument("--email-column", help="Column name in CSV for emails (auto-detected if not specified)")
    send_parser.add_argument("--subject", required=True, help="Email subject")
    send_parser.add_argument("--body", required=True, help="Email body")
    send_parser.add_argument("--attachments", action="append", help="File paths to attach (can use multiple --attachments)")
    send_parser.add_argument("--body-type", choices=["html", "text"], default="html")
    
    # Reply command (supports batching automatically)
    # Default behavior: reply to all recipients (sender + original To/CC)
    # Use --sender-only to reply only to the original sender
    reply_parser = subparsers.add_parser("reply", help="Reply to an email (auto-batch for large recipient lists)")
    reply_parser.add_argument("message_id", help="Message ID to reply to")
    reply_parser.add_argument("--body", required=True, help="Reply body")
    reply_parser.add_argument("--sender-only", dest="reply_all", action="store_false", 
                              help="Reply only to sender (default: reply to all)")
    reply_parser.add_argument("--to", help="Additional To recipients (comma-separated)")
    reply_parser.add_argument("--cc", help="Additional CC recipients (comma-separated)")
    reply_parser.add_argument("--bcc", help="BCC recipients (comma-separated)")
    reply_parser.add_argument("--csv", dest="csv_path", help="CSV file containing BCC email addresses")
    reply_parser.add_argument("--email-column", help="Column name in CSV for emails (auto-detected if not specified)")
    reply_parser.add_argument("--attachments", action="append", help="File paths to attach (can use multiple --attachments)")
    reply_parser.add_argument("--importance", choices=["low", "normal", "high"], help="Email importance level")
    
    # Forward command (supports batching automatically)
    forward_parser = subparsers.add_parser("forward", help="Forward an email (auto-batch for large recipient lists)")
    forward_parser.add_argument("message_id", help="Message ID to forward")
    forward_parser.add_argument("--to", help="To recipients (comma-separated)")
    forward_parser.add_argument("--cc", help="CC recipients (comma-separated)")
    forward_parser.add_argument("--bcc", help="BCC recipients (comma-separated)")
    forward_parser.add_argument("--csv", dest="csv_path", help="CSV file containing BCC email addresses")
    forward_parser.add_argument("--email-column", help="Column name in CSV for emails (auto-detected if not specified)")
    forward_parser.add_argument("--comment", default="", help="Comment to add")
    
    # Mark read/unread
    read_parser = subparsers.add_parser("read", help="Mark message as read/unread")
    read_parser.add_argument("message_id", help="Message ID")
    read_parser.add_argument("--unread", action="store_true", help="Mark as unread")
    
    # Delete command
    delete_parser = subparsers.add_parser("delete", help="Delete an email")
    delete_parser.add_argument("message_id", help="Message ID to delete")
    
    # List folders command
    folders_parser = subparsers.add_parser("folders", help="List all mail folders")
    folders_parser.add_argument("--all", dest="include_hidden", action="store_true", help="Include hidden folders")
    
    # Attachments command
    att_parser = subparsers.add_parser("attachments", help="List or download attachments")
    att_parser.add_argument("message_id", help="Message ID")
    att_parser.add_argument("--download", "-d", action="store_true", help="Download all attachments")
    att_parser.add_argument("--save-dir", help="Directory to save attachments (default: current directory)")
    att_parser.add_argument("--id", dest="attachment_id", help="Download specific attachment by ID")
    
    # Accept meeting invite command
    accept_parser = subparsers.add_parser("accept-invite", help="Accept a meeting invite")
    accept_parser.add_argument("message_id", help="Message ID of the meeting invite")
    accept_parser.add_argument("--comment", help="Optional comment to include in response")
    accept_parser.add_argument("--no-send", action="store_true", help="Don't send response to organizer")
    
    # Decline meeting invite command
    decline_parser = subparsers.add_parser("decline-invite", help="Decline a meeting invite")
    decline_parser.add_argument("message_id", help="Message ID of the meeting invite")
    decline_parser.add_argument("--comment", help="Optional comment to include in response")
    decline_parser.add_argument("--no-send", action="store_true", help="Don't send response to organizer")
    
    args = parser.parse_args()
    
    # Pre-process body/comment arguments: convert escape sequences like \n to actual newlines
    # This allows using \n in any shell (PowerShell, bash, cmd) for line breaks
    if hasattr(args, 'body') and args.body:
        args.body = unescape_body(args.body)
    if hasattr(args, 'comment') and args.comment:
        args.comment = unescape_body(args.comment)
    
    # Auto-convert Outlook search syntax (e.g., "from:beng" -> --from "beng")
    syntax_warnings = convert_outlook_syntax_args(args)
    for warning in syntax_warnings:
        print(warning)
    
    try:
        if args.command in ["list", "search", "find"]:
            # Unified handler for list, search and find commands (identical functionality)
            filter_query = args.filter_query
            if args.unread:
                filter_query = (filter_query + " and " if filter_query else "") + "isRead eq false"
            
            # Determine inference classification
            inference_classification = None
            if args.focused:
                inference_classification = "focused"
            elif args.other:
                inference_classification = "other"
            
            # Check if --detail flag is set
            show_detail = getattr(args, 'detail', False)
            
            # Determine message_type based on flags
            message_type = "all"  # default: show all with categorization
            if getattr(args, 'emails_only', False):
                message_type = "emails"
            elif getattr(args, 'events_only', False):
                message_type = "events"
            
            messages, time_info = list_messages(
                folder=args.folder,
                limit=args.limit,
                filter_query=filter_query,
                include_preview=True,  # Always include bodyPreview in API response for flexibility
                inference_classification=inference_classification,
                from_sender=getattr(args, 'from_sender', None),
                to_recipient=getattr(args, 'to_recipient', None),
                subject=getattr(args, 'subject', None),
                body=getattr(args, 'body', None),
                start=getattr(args, 'start', None),
                end=getattr(args, 'end', None),
                message_type=message_type,
                display_timezone=args.timezone
            )
            
            # If --detail flag is set, fetch full message content for each result
            if show_detail and messages:
                detailed_messages = []
                for msg in messages:
                    full_msg = get_message(msg['id'])
                    detailed_messages.append(full_msg)
                messages = detailed_messages
            
            if args.json:
                result = {"success": True, "messages": messages, "total": len(messages)}
                if time_info:
                    if time_info.get('start'):
                        result["start_info"] = {
                            **time_info['start'],
                            "_description": f"Emails received after {time_info['start']['original']}"
                        }
                    if time_info.get('end'):
                        result["end_info"] = {
                            **time_info['end'],
                            "_description": f"Emails received before {time_info['end']['original']}"
                        }
                print(json.dumps(result, indent=2, default=str))
            else:
                # Use user-specified timezone for display
                display_tz = getattr(args, 'timezone', None)
                
                # Display timezone first
                if display_tz:
                    print(f"   Display Timezone: {display_tz}")
                
                # Display time info in user-friendly format
                if time_info:
                    if time_info.get('start'):
                        # Parse and format the start timestamp in a clean way
                        start_original = time_info['start']['original']
                        try:
                            # Parse the timestamp
                            start_dt = datetime.fromisoformat(start_original.replace('Z', '+00:00'))
                            # Convert to display timezone if available
                            if display_tz:
                                try:
                                    if isinstance(display_tz, str):
                                        display_tz_obj = ZoneInfo(display_tz)
                                    else:
                                        display_tz_obj = display_tz
                                    start_dt = start_dt.astimezone(display_tz_obj)
                                except:
                                    pass
                            # Format cleanly without timezone suffix
                            start_formatted = start_dt.strftime('%Y-%m-%d %H:%M')
                            print(f"\n📅 Start: {start_formatted}")
                        except:
                            # Fallback to original if parsing fails
                            print(f"\n📅 Start: {start_original}")
                    
                    if time_info.get('end'):
                        # Parse and format the end timestamp in a clean way
                        end_original = time_info['end']['original']
                        try:
                            # Parse the timestamp
                            end_dt = datetime.fromisoformat(end_original.replace('Z', '+00:00'))
                            # Convert to display timezone if available
                            if display_tz:
                                try:
                                    if isinstance(display_tz, str):
                                        display_tz_obj = ZoneInfo(display_tz)
                                    else:
                                        display_tz_obj = display_tz
                                    end_dt = end_dt.astimezone(display_tz_obj)
                                except:
                                    pass
                            # Format cleanly without timezone suffix
                            end_formatted = end_dt.strftime('%Y-%m-%d %H:%M')
                            print(f"\n📅 End: {end_formatted}")
                        except:
                            # Fallback to original if parsing fails
                            print(f"\n📅 End: {end_original}")
                    
                display_message_list(messages, show_preview=True, show_detail=show_detail, display_timezone=display_tz, message_type=message_type)
        
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
            # Handle multiple --to/--cc/--bcc arguments (action="append" creates a list)
            to_list = []
            if args.to:
                for to_arg in args.to:
                    to_list.extend(parse_email_list(to_arg))
            
            cc_list = []
            if args.cc:
                for cc_arg in args.cc:
                    cc_list.extend(parse_email_list(cc_arg))
            
            bcc_list = []
            if args.bcc:
                for bcc_arg in args.bcc:
                    bcc_list.extend(parse_email_list(bcc_arg))
            
            # Handle attachments
            inline_attachments = None
            large_files = []
            if hasattr(args, 'attachments') and args.attachments:
                # Parse comma-separated file paths (similar to email parsing)
                file_list = []
                for att_arg in args.attachments:
                    # Split by comma and strip whitespace
                    files = [f.strip() for f in att_arg.split(',') if f.strip()]
                    file_list.extend(files)
                
                try:
                    inline_attachments, large_files = prepare_file_attachments(file_list)
                    if inline_attachments:
                        print(f"✓ Prepared {len(inline_attachments)} inline attachment(s)")
                    if large_files:
                        print(f"✓ Detected {len(large_files)} large file(s) for upload session")
                except Exception as e:
                    print(f"✗ Error preparing attachments: {str(e)}")
                    sys.exit(1)
            
            # If we have large files, use draft + upload session workflow
            if large_files:
                print("📧 Creating draft message for large attachments...")
                draft_id = create_draft_message(
                    to=to_list if to_list else [],
                    subject=args.subject,
                    body=args.body,
                    cc=cc_list if cc_list else None,
                    bcc=bcc_list if bcc_list else None,
                    body_type=args.body_type,
                    importance=getattr(args, 'importance', None)
                )
                print(f"✓ Draft created: {draft_id}")
                
                # Add inline attachments to draft if any
                if inline_attachments:
                    print(f"📎 Adding {len(inline_attachments)} inline attachment(s) to draft...")
                    url = f"{GRAPH_API_BASE}/me/messages/{draft_id}"
                    token = get_access_token()
                    # Update draft with inline attachments
                    for attachment in inline_attachments:
                        attach_url = f"{url}/attachments"
                        api_request('post', attach_url, token, json=attachment)
                    print(f"✓ Inline attachments added")
                
                # Upload large files
                for file_path in large_files:
                    upload_large_attachment(draft_id, file_path)
                
                # Send the draft
                print("📤 Sending message...")
                send_draft_message(draft_id)
                print("✓ Email sent successfully")
            else:
                # Normal send with inline attachments only
                result = batch_send_email(
                    to=to_list if to_list else None,
                    subject=args.subject,
                    body=args.body,
                    cc=cc_list if cc_list else None,
                    bcc=bcc_list if bcc_list else None,
                    csv_path=getattr(args, 'csv_path', None),
                    email_column=getattr(args, 'email_column', None),
                    body_type=args.body_type,
                    attachments=inline_attachments
                )
                if args.json:
                    print(json.dumps(result, indent=2))
                elif result['total_batches'] > 1:
                    print(f"\n{'='*60}")
                    print(f"SEND RESULTS (Batched)")
                    print(f"{'='*60}")
                    print(f"Total recipients: {result['total_recipients']}")
                    print(f"Total batches: {result['total_batches']}")
                    print(f"Sent: {result['sent_count']}")
                    print(f"Failed: {result['failed_count']}")
                    if result.get('errors'):
                        print(f"\nErrors:")
                        for error in result['errors']:
                            print(f"  - {error}")
                    print(f"{'='*60}")
                else:
                    print("✓ Email sent successfully")
        
        elif args.command == "reply":
            # Handle attachments
            inline_attachments = None
            large_files = []
            if hasattr(args, 'attachments') and args.attachments:
                # Parse comma-separated file paths (similar to email parsing)
                file_list = []
                for att_arg in args.attachments:
                    # Split by comma and strip whitespace
                    files = [f.strip() for f in att_arg.split(',') if f.strip()]
                    file_list.extend(files)
                
                try:
                    inline_attachments, large_files = prepare_file_attachments(file_list)
                    if inline_attachments:
                        print(f"✓ Prepared {len(inline_attachments)} inline attachment(s)")
                    if large_files:
                        print(f"✓ Detected {len(large_files)} large file(s) for upload session")
                except Exception as e:
                    print(f"✗ Error preparing attachments: {str(e)}")
                    sys.exit(1)
            
            # Determine which function to use:
            # - Standard reply_email: when replying to original sender only (no extra recipients/attachments)
            # - batch_reply_email: when adding custom recipients via --to/--cc/--bcc/--csv or attachments
            # - Draft workflow: when large files are present
            has_extra_recipients = args.to or args.cc or args.bcc or getattr(args, 'csv_path', None)
            has_attachments = inline_attachments is not None or large_files
            
            # Large files require draft workflow
            if large_files:
                print("📧 Creating reply draft for large attachments...")
                token = get_access_token()
                
                # Create reply draft
                url = f"{GRAPH_API_BASE}/me/messages/{args.message_id}/createReply"
                response = api_request('post', url, token)
                draft = response.json()
                draft_id = draft['id']
                print(f"✓ Reply draft created: {draft_id}")
                
                # Get the draft to see existing body (which includes thread)
                update_url = f"{GRAPH_API_BASE}/me/messages/{draft_id}"
                draft_response = api_request('get', update_url, token)
                draft_data = draft_response.json()
                existing_body = draft_data.get('body', {}).get('content', '')
                
                # Prepend our new content to existing body (which has the thread)
                new_content = args.body
                if not new_content.strip().startswith('<'):
                    new_content = new_content.replace('\n', '<br>')
                
                # Combine: new content + existing thread
                combined_body = f"{new_content}<br><br>{existing_body}"
                
                update_payload = {
                    "body": {
                        "contentType": "html",
                        "content": combined_body
                    }
                }
                api_request('patch', update_url, token, json=update_payload)
                print("✓ Reply body updated with thread preserved")
                
                # Add inline attachments if any
                if inline_attachments:
                    print(f"📎 Adding {len(inline_attachments)} inline attachment(s)...")
                    for attachment in inline_attachments:
                        attach_url = f"{update_url}/attachments"
                        api_request('post', attach_url, token, json=attachment)
                    print(f"✓ Inline attachments added")
                
                # Upload large files
                for file_path in large_files:
                    upload_large_attachment(draft_id, file_path)
                
                # Send the draft
                print("📤 Sending reply...")
                send_draft_message(draft_id)
                
                if args.json:
                    print(json.dumps({"success": True, "message": "Reply sent successfully"}))
                else:
                    print("✓ Reply sent successfully")
            
            elif not has_extra_recipients and not has_attachments:
                # Standard reply: use reply_email which auto-extracts sender from original message
                reply_email(
                    message_id=args.message_id,
                    body=args.body,
                    reply_all=args.reply_all,
                    importance=getattr(args, 'importance', None)
                )
                if args.json:
                    print(json.dumps({"success": True, "message": "Reply sent successfully"}))
                else:
                    print("✓ Reply sent successfully")
            else:
                # Note: If attachments are specified, we must use batch_reply_email
                # because native reply endpoint doesn't support attachments
                if has_attachments and not has_extra_recipients:
                    print("ℹ️  Note: Using custom reply mode to support attachments")
                
                result = batch_reply_email(
                    message_id=args.message_id,
                    body=args.body,
                    to=parse_email_list(args.to) if args.to else None,
                    cc=parse_email_list(args.cc) if args.cc else None,
                    bcc=parse_email_list(args.bcc) if args.bcc else None,
                    csv_path=getattr(args, 'csv_path', None),
                    email_column=getattr(args, 'email_column', None),
                    attachments=inline_attachments
                )
                if args.json:
                    print(json.dumps(result, indent=2))
                elif result['total_batches'] > 1:
                    print(f"\n{'='*60}")
                    print(f"REPLY RESULTS (Batched)")
                    print(f"{'='*60}")
                    print(f"Total recipients: {result['total_recipients']}")
                    print(f"Total batches: {result['total_batches']}")
                    print(f"Sent: {result['sent_count']}")
                    print(f"Failed: {result['failed_count']}")
                    if result.get('errors'):
                        print(f"\nErrors:")
                        for error in result['errors']:
                            print(f"  - {error}")
                    print(f"{'='*60}")
                else:
                    print("✓ Reply sent successfully")
        
        elif args.command == "forward":
            result = batch_forward_email(
                message_id=args.message_id,
                to=parse_email_list(args.to) if args.to else None,
                cc=parse_email_list(args.cc) if args.cc else None,
                bcc=parse_email_list(args.bcc) if args.bcc else None,
                csv_path=getattr(args, 'csv_path', None),
                email_column=getattr(args, 'email_column', None),
                comment=args.comment
            )
            if args.json:
                print(json.dumps(result, indent=2))
            elif result['total_batches'] > 1:
                print(f"\n{'='*60}")
                print(f"FORWARD RESULTS (Batched)")
                print(f"{'='*60}")
                print(f"Total recipients: {result['total_recipients']}")
                print(f"Total batches: {result['total_batches']}")
                print(f"Sent: {result['sent_count']}")
                print(f"Failed: {result['failed_count']}")
                if result.get('errors'):
                    print(f"\nErrors:")
                    for error in result['errors']:
                        print(f"  - {error}")
                print(f"{'='*60}")
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
        
        elif args.command == "folders":
            folders = list_mail_folders(include_hidden=getattr(args, 'include_hidden', False))
            if args.json:
                print(json.dumps({"success": True, "folders": folders, "total": len(folders)}, indent=2, default=str))
            else:
                display_folder_list(folders)
        
        elif args.command == "attachments":
            if args.download or args.attachment_id:
                # Download mode
                if args.attachment_id:
                    # Download specific attachment
                    result = download_attachment(args.message_id, args.attachment_id, args.save_dir)
                    if args.json:
                        print(json.dumps({"success": True, "result": result}, indent=2))
                    else:
                        print(f"✓ Downloaded: {result['file_name']}")
                        print(f"  Path: {result['file_path']}")
                        print(f"  Size: {result['size']} bytes")
                else:
                    # Download all attachments
                    results = download_all_attachments(args.message_id, args.save_dir)
                    if args.json:
                        print(json.dumps({"success": True, "results": results}, indent=2))
                    else:
                        print(f"\n{'='*60}")
                        print(f"ATTACHMENT DOWNLOAD RESULTS")
                        print(f"{'='*60}")
                        for r in results:
                            if r.get('success'):
                                print(f"✓ {r['file_name']} -> {r['file_path']}")
                            else:
                                print(f"✗ {r['file_name']}: {r.get('error', 'Unknown error')}")
                        print(f"{'='*60}")
            else:
                # List mode
                attachments = list_attachments(args.message_id)
                if args.json:
                    print(json.dumps({"success": True, "attachments": attachments, "total": len(attachments)}, indent=2, default=str))
                else:
                    display_attachments(attachments)
        
        elif args.command == "accept-invite":
            try:
                accept_meeting_invite(
                    message_id=args.message_id,
                    comment=args.comment,
                    send_response=not args.no_send
                )
                if args.json:
                    print(json.dumps({"success": True, "message": "Meeting invite accepted"}))
                else:
                    print("✓ Meeting invite accepted")
            except ValueError as e:
                if args.json:
                    print(json.dumps({"success": False, "error": str(e)}))
                else:
                    print(f"✗ Error: {e}")
                sys.exit(1)
        
        elif args.command == "decline-invite":
            try:
                decline_meeting_invite(
                    message_id=args.message_id,
                    comment=args.comment,
                    send_response=not args.no_send
                )
                if args.json:
                    print(json.dumps({"success": True, "message": "Meeting invite declined"}))
                else:
                    print("✓ Meeting invite declined")
            except ValueError as e:
                if args.json:
                    print(json.dumps({"success": False, "error": str(e)}))
                else:
                    print(f"✗ Error: {e}")
                sys.exit(1)
    
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
