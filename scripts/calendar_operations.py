#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Microsoft Graph Calendar Operations Module

Provides calendar operations including event creation, management,
and availability queries.

Usage:
    python calendar_operations.py list [--limit <n>]
    python calendar_operations.py get <event_id>
    python calendar_operations.py create --subject <subject> --start <datetime> --end <datetime> [--attendees <emails>]
    python calendar_operations.py availability --emails <emails> --start <datetime> --end <datetime>
    python calendar_operations.py update <event_id> [--subject <subject>] ...
    python calendar_operations.py delete <event_id>
"""

import os
import sys
import json
import argparse
from pathlib import Path
from typing import List, Optional, Dict, Any
from datetime import datetime, timedelta

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Add parent directory to path for config import
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import configuration and auth
from config import GRAPH_API_BASE, DATE_FORMAT
from auth import get_access_token

# Try to import requests
try:
    import requests
except ImportError:
    print("Error: requests package not found.")
    print("Install with: pip install requests")
    sys.exit(1)


def get_headers(token: str) -> Dict[str, str]:
    """Get authorization headers for API requests."""
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }


def parse_datetime(dt_str: str, param_name: str = "datetime", display_timezone: str = None) -> tuple:
    """
    Parse datetime string with timezone.
    
    Args:
        dt_str: Datetime string (plain datetime or "now")
        param_name: Parameter name for error messages
        display_timezone: REQUIRED - timezone for parsing and display
    
    Returns:
        Tuple of (datetime_str, timezone_str, error)
        - datetime_str: ISO format datetime without timezone (e.g., "2026-03-26T12:00:00")
        - timezone_str: IANA timezone for API (e.g., "Asia/Shanghai" or "UTC")
        - error: Error message if validation fails, None otherwise
    """
    from zoneinfo import ZoneInfo
    
    if not dt_str:
        return None, None, None
    
    # display_timezone is always required
    if not display_timezone:
        return None, None, (
            f"TIMEZONE_REQUIRED: --timezone is required.\n"
            f"Example: --{param_name} '2026-03-26T12:00:00' --timezone 'Asia/Shanghai'"
        )
    
    # Handle "now" special value
    if dt_str.lower() == "now":
        tz = ZoneInfo(display_timezone)
        now_dt = datetime.now(tz)
        datetime_str = now_dt.strftime("%Y-%m-%dT%H:%M:%S")
        return datetime_str, display_timezone, None
    
    # Reject embedded timezone (Z or +08:00)
    if dt_str.endswith('Z') or '+' in dt_str or (dt_str.count('-') > 2):
        return None, None, (
            f"INVALID_FORMAT: '{dt_str}' contains embedded timezone.\n"
            f"Use plain datetime with --timezone instead:\n"
            f"  --{param_name} '2026-03-26T12:00:00' --timezone 'Asia/Shanghai'"
        )
    
    # Parse plain datetime (no timezone)
    formats = [
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%d",
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(dt_str, fmt)
            # Add seconds if not present
            if fmt == "%Y-%m-%d":
                dt = dt.replace(hour=0, minute=0, second=0)
            datetime_str = dt.strftime("%Y-%m-%dT%H:%M:%S")
            return datetime_str, display_timezone, None
        except ValueError:
            continue
    
    return None, None, (
        f"INVALID_FORMAT: '{dt_str}' is not a valid datetime.\n"
        f"Expected format: '2026-03-26T12:00:00' or '2026-03-26'"
    )


def parse_datetime_or_raise(dt_str: str, param_name: str = "datetime", display_timezone: str = None) -> tuple:
    """
    Parse datetime string and raise error if validation fails.
    
    Returns:
        Tuple of (datetime_str, timezone_str)
    """
    dt, tz, error = parse_datetime(dt_str, param_name, display_timezone)
    if error:
        raise ValueError(error)
    return dt, tz


def parse_email_list(emails: str) -> List[str]:
    """Parse comma or semicolon separated email list."""
    if not emails:
        return []
    emails = emails.replace(';', ',')
    return [e.strip() for e in emails.split(',') if e.strip()]


# =============================================================================
# LIST EVENTS
# =============================================================================

def list_events(
    calendar_id: str = None,
    display_timezone: str = None,
    start: str = None,
    end: str = None,
    limit: int = 25,
    filter_query: str = None,
    token: str = None
) -> List[Dict[str, Any]]:
    """
    List calendar events.
    
    Args:
        calendar_id: Specific calendar ID (uses default if not provided)
        start: Start datetime (must include timezone, or "now")
        end: End datetime (must include timezone, or "now")
        limit: Maximum number of events to return
        filter_query: OData filter query
        token: Access token
        display_timezone: Timezone for "now" calculation and display
    
    Returns:
        List of event objects
    """
    if token is None:
        token = get_access_token()
    
    if calendar_id:
        url = f"{GRAPH_API_BASE}/me/calendars/{calendar_id}/events"
    else:
        url = f"{GRAPH_API_BASE}/me/calendar/events"
    
    params = {
        "$top": limit,
        "$orderby": "start/dateTime",
        "$select": "id,subject,start,end,organizer,attendees,isAllDay,location,responseStatus"
    }
    
    # Build filter query for date range
    filter_parts = []
    
    if start or end:
        if start:
            start_iso, start_tz = parse_datetime_or_raise(start, 'start', display_timezone)
        else:
            # Default to now
            from zoneinfo import ZoneInfo
            tz = ZoneInfo(display_timezone)
            start_iso = datetime.now(tz).strftime("%Y-%m-%dT%H:%M:%S")
        
        if end:
            end_iso, end_tz = parse_datetime_or_raise(end, 'end', display_timezone)
        else:
            # Default to 30 days from now
            from zoneinfo import ZoneInfo
            tz = ZoneInfo(display_timezone)
            end_iso = (datetime.now(tz) + timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S")
        
        filter_parts.append(f"start/dateTime ge '{start_iso}'")
        filter_parts.append(f"end/dateTime le '{end_iso}'")
    
    if filter_query:
        filter_parts.append(filter_query)
    
    if filter_parts:
        params["$filter"] = " and ".join(filter_parts)
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to list events: {response.status_code} - {response.text}")
    
    data = response.json()
    return data.get("value", [])


# =============================================================================
# GET EVENT
# =============================================================================

def get_event(event_id: str, token: str = None) -> Dict[str, Any]:
    """
    Get a specific event by ID.
    
    Args:
        event_id: Event ID
        token: Access token
    
    Returns:
        Event object
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/events/{event_id}"
    
    response = requests.get(url, headers=get_headers(token))
    
    if response.status_code != 200:
        raise Exception(f"Failed to get event: {response.status_code} - {response.text}")
    
    return response.json()


# =============================================================================
# CREATE EVENT
# =============================================================================

def create_event(
    subject: str,
    start: str,
    end: str,
    display_timezone: str,
    body: str = None,
    body_type: str = "html",
    location: str = None,
    attendees: List[Dict[str, str]] = None,
    is_all_day: bool = False,
    recurrence: Dict = None,
    is_online_meeting: bool = False,
    token: str = None
) -> Dict[str, Any]:
    """
    Create a calendar event.
    
    Args:
        subject: Event subject
        start: Start datetime (must include timezone, or "now")
        end: End datetime (must include timezone, or "now")
        timezone: (DEPRECATED - extracted from start/end) Fallback timezone
        body: Event body/description
        body_type: "html" or "text"
        location: Location string
        attendees: List of attendee dicts with 'email' and optionally 'name', 'type'
        is_all_day: Whether this is an all-day event
        recurrence: Recurrence pattern dict
        is_online_meeting: Whether to create as Teams meeting
        token: Access token
        display_timezone: Timezone for "now" calculation
    
    Returns:
        Created event object
    """
    if token is None:
        token = get_access_token()
    
    # Parse start and end times, extracting timezone
    start_dt, start_tz = parse_datetime_or_raise(start, 'start', display_timezone)
    end_dt, end_tz = parse_datetime_or_raise(end, 'end', display_timezone)
    
    # Use extracted timezone (prefer start timezone)
    event_tz = start_tz or "UTC"
    
    event = {
        "subject": subject,
        "start": {
            "dateTime": start_dt,
            "timeZone": event_tz
        },
        "end": {
            "dateTime": end_dt,
            "timeZone": event_tz
        },
        "isAllDay": is_all_day
    }
    
    if body:
        event["body"] = {
            "contentType": body_type,
            "content": body
        }
    
    if location:
        event["location"] = {"displayName": location}
    
    if attendees:
        event["attendees"] = [
            {
                "emailAddress": {
                    "address": a.get("email"),
                    "name": a.get("name", a.get("email"))
                },
                "type": a.get("type", "required")
            }
            for a in attendees
        ]
    
    if recurrence:
        event["recurrence"] = recurrence
    
    if is_online_meeting:
        event["isOnlineMeeting"] = True
        event["onlineMeetingProvider"] = "teamsForBusiness"
    
    url = f"{GRAPH_API_BASE}/me/events"
    
    response = requests.post(url, headers=get_headers(token), json=event)
    
    if response.status_code != 201:
        raise Exception(f"Failed to create event: {response.status_code} - {response.text}")
    
    return response.json()


# =============================================================================
# UPDATE EVENT
# =============================================================================

def update_event(
    event_id: str,
    display_timezone: str,
    subject: str = None,
    start: str = None,
    end: str = None,
    body: str = None,
    body_type: str = None,
    location: str = None,
    attendees: List[Dict[str, str]] = None,
    token: str = None
) -> Dict[str, Any]:
    """
    Update a calendar event.
    
    Args:
        event_id: Event ID to update
        start: Start datetime (must include timezone, or "now")
        end: End datetime (must include timezone, or "now")
        timezone: (DEPRECATED - extracted from start/end) Fallback timezone
        display_timezone: Timezone for "now" calculation
        Other args: Fields to update
    
    Returns:
        Updated event object
    """
    if token is None:
        token = get_access_token()
    
    event = {}
    
    if subject is not None:
        event["subject"] = subject
    
    if start:
        start_dt, start_tz = parse_datetime_or_raise(start, 'start', display_timezone)
        event["start"] = {
            "dateTime": start_dt,
            "timeZone": start_tz or "UTC"
        }
    
    if end:
        end_dt, end_tz = parse_datetime_or_raise(end, 'end', display_timezone)
        event["end"] = {
            "dateTime": end_dt,
            "timeZone": end_tz or "UTC"
        }
    
    if body is not None:
        event["body"] = {
            "contentType": body_type or "html",
            "content": body
        }
    
    if location is not None:
        event["location"] = {"displayName": location}
    
    if attendees is not None:
        event["attendees"] = [
            {
                "emailAddress": {
                    "address": a.get("email"),
                    "name": a.get("name", a.get("email"))
                },
                "type": a.get("type", "required")
            }
            for a in attendees
        ]
    
    url = f"{GRAPH_API_BASE}/me/events/{event_id}"
    
    response = requests.patch(url, headers=get_headers(token), json=event)
    
    if response.status_code != 200:
        raise Exception(f"Failed to update event: {response.status_code} - {response.text}")
    
    return response.json()


# =============================================================================
# DELETE EVENT
# =============================================================================

def delete_event(event_id: str, permanent: bool = False, token: str = None) -> bool:
    """
    Delete a calendar event.
    
    Default behavior (soft delete): Moves event to "Deleted Items" folder.
    Can be recovered from there. Does NOT notify attendees.
    
    Permanent delete: Permanently removes the event. Cannot be recovered.
    
    Note: To cancel an event and notify attendees (as organizer), use cancel_event() instead.
    
    Args:
        event_id: Event ID to delete
        permanent: If True, permanently delete (default False = soft delete)
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    if permanent:
        # Permanent delete - cannot be recovered
        url = f"{GRAPH_API_BASE}/me/events/{event_id}"
        response = requests.delete(url, headers=get_headers(token))
        if response.status_code != 204:
            raise Exception(f"Failed to permanently delete event: {response.status_code} - {response.text}")
    else:
        # Soft delete - move to Deleted Items folder
        url = f"{GRAPH_API_BASE}/me/events/{event_id}/move"
        payload = {
            "destinationId": "deleteditems"
        }
        response = requests.post(url, headers=get_headers(token), json=payload)
        if response.status_code not in [200, 201]:
            raise Exception(f"Failed to soft delete event: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# ACCEPT EVENT
# =============================================================================

def accept_event(
    event_id: str,
    comment: str = None,
    send_response: bool = True,
    token: str = None
) -> bool:
    """
    Accept a calendar event invitation.
    
    Args:
        event_id: Event ID to accept
        comment: Optional comment to include in response
        send_response: Whether to send response to organizer (default True)
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/events/{event_id}/accept"
    
    payload = {}
    if comment:
        payload["comment"] = comment
    if not send_response:
        payload["sendResponse"] = False
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code not in [200, 202]:
        raise Exception(f"Failed to accept event: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# DECLINE EVENT
# =============================================================================

def decline_event(
    event_id: str,
    comment: str = None,
    send_response: bool = True,
    token: str = None
) -> bool:
    """
    Decline a calendar event invitation.
    
    Args:
        event_id: Event ID to decline
        comment: Optional comment to include in response
        send_response: Whether to send response to organizer (default True)
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/events/{event_id}/decline"
    
    payload = {}
    if comment:
        payload["comment"] = comment
    if not send_response:
        payload["sendResponse"] = False
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code not in [200, 202]:
        raise Exception(f"Failed to decline event: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# TENTATIVELY ACCEPT EVENT
# =============================================================================

def tentatively_accept_event(
    event_id: str,
    comment: str = None,
    send_response: bool = True,
    token: str = None
) -> bool:
    """
    Tentatively accept a calendar event invitation.
    
    Args:
        event_id: Event ID to tentatively accept
        comment: Optional comment to include in response
        send_response: Whether to send response to organizer (default True)
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/events/{event_id}/tentativelyAccept"
    
    payload = {}
    if comment:
        payload["comment"] = comment
    if not send_response:
        payload["sendResponse"] = False
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code not in [200, 202]:
        raise Exception(f"Failed to tentatively accept event: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# CANCEL EVENT (with notification to attendees)
# =============================================================================

def cancel_event(
    event_id: str,
    comment: str = None,
    token: str = None
) -> bool:
    """
    Cancel a calendar event and send cancellation notifications to attendees.
    Only the organizer can cancel an event.
    
    Args:
        event_id: Event ID to cancel
        comment: Optional cancellation message to send to attendees
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/events/{event_id}/cancel"
    
    payload = {}
    if comment:
        payload["comment"] = comment
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code not in [200, 202, 204]:
        raise Exception(f"Failed to cancel event: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# FORWARD EVENT
# =============================================================================

def forward_event(
    event_id: str,
    to_emails: List[str],
    comment: str = None,
    token: str = None
) -> bool:
    """
    Forward a calendar event to new recipients (adds them as optional attendees).
    
    Args:
        event_id: Event ID to forward
        to_emails: List of email addresses to forward to
        comment: Optional message to include
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/events/{event_id}/forward"
    
    payload = {
        "toRecipients": [
            {
                "emailAddress": {
                    "address": email
                }
            }
            for email in to_emails
        ]
    }
    if comment:
        payload["comment"] = comment
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code not in [200, 202, 204]:
        raise Exception(f"Failed to forward event: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# PROPOSE NEW TIME
# =============================================================================

def propose_new_time(
    event_id: str,
    new_start: str,
    new_end: str,
    display_timezone: str,
    comment: str = None,
    send_response: bool = True,
    token: str = None
) -> bool:
    """
    Propose a new meeting time for an event (declines current time and proposes new).
    
    Args:
        event_id: Event ID
        new_start: Proposed new start datetime (must include timezone, or "now")
        new_end: Proposed new end datetime (must include timezone, or "now")
        timezone: (DEPRECATED - extracted from new_start/new_end) Fallback timezone
        comment: Optional message to organizer
        send_response: Whether to send response to organizer (default True)
        token: Access token
        display_timezone: Timezone for "now" calculation
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    # Parse times and extract timezones
    new_start_dt, new_start_tz = parse_datetime_or_raise(new_start, 'start', display_timezone)
    new_end_dt, new_end_tz = parse_datetime_or_raise(new_end, 'end', display_timezone)
    
    url = f"{GRAPH_API_BASE}/me/events/{event_id}/decline"
    
    payload = {
        "proposedNewTime": {
            "start": {
                "dateTime": new_start_dt,
                "timeZone": new_start_tz or "UTC"
            },
            "end": {
                "dateTime": new_end_dt,
                "timeZone": new_end_tz or "UTC"
            }
        }
    }
    if comment:
        payload["comment"] = comment
    if not send_response:
        payload["sendResponse"] = False
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code not in [200, 202]:
        raise Exception(f"Failed to propose new time: {response.status_code} - {response.text}")
    
    return True


# =============================================================================
# GET AVAILABILITY (FREE/BUSY)
# =============================================================================

def get_availability(
    emails: List[str],
    start: str,
    end: str,
    display_timezone: str,
    interval: int = 30,
    token: str = None
) -> Dict[str, Any]:
    """
    Get free/busy availability for specified users.
    
    Args:
        emails: List of email addresses to check
        start: Start datetime (must include timezone, or "now")
        end: End datetime (must include timezone, or "now")
        timezone: (DEPRECATED - extracted from start/end) Fallback timezone
        interval: Meeting time slot interval in minutes
        token: Access token
        display_timezone: Timezone for "now" calculation
    
    Returns:
        Availability information for each user
    """
    if token is None:
        token = get_access_token()
    
    # Parse times and extract timezones
    start_dt, start_tz = parse_datetime_or_raise(start, 'start', display_timezone)
    end_dt, end_tz = parse_datetime_or_raise(end, 'end', display_timezone)
    
    # Use extracted timezone (prefer start timezone)
    event_tz = start_tz or "UTC"
    
    url = f"{GRAPH_API_BASE}/me/calendar/getSchedule"
    
    payload = {
        "schedules": emails,
        "startTime": {
            "dateTime": start_dt,
            "timeZone": event_tz
        },
        "endTime": {
            "dateTime": end_dt,
            "timeZone": event_tz
        },
        "availabilityViewInterval": interval
    }
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code != 200:
        raise Exception(f"Failed to get availability: {response.status_code} - {response.text}")
    
    return response.json()


def get_users_info(
    emails: List[str],
    token: str = None
) -> Dict[str, str]:
    """
    Get display names for a list of email addresses.
    
    Args:
        emails: List of email addresses
        token: Access token
    
    Returns:
        Dictionary mapping email to display name
    """
    if token is None:
        token = get_access_token()
    
    email_to_name = {}
    
    # First, get current user's info
    current_user_email = None
    current_user_name = None
    try:
        url = f"{GRAPH_API_BASE}/me?$select=displayName,mail,userPrincipalName"
        response = requests.get(url, headers=get_headers(token))
        if response.status_code == 200:
            data = response.json()
            current_user_email = data.get('mail') or data.get('userPrincipalName')
            current_user_name = data.get('displayName')
    except Exception as e:
        pass
    
    for email in emails:
        # Check if this is the current user
        if current_user_email and email.lower() == current_user_email.lower() and current_user_name:
            email_to_name[email] = current_user_name
            continue
        
        try:
            # Try to get user info from directory
            url = f"{GRAPH_API_BASE}/users/{email}?$select=displayName,mail"
            response = requests.get(url, headers=get_headers(token))
            
            if response.status_code == 200:
                data = response.json()
                display_name = data.get('displayName', email.split('@')[0])
                email_to_name[email] = display_name
            else:
                # Fallback to short name from email (silently, could be rate limit)
                email_to_name[email] = email.split('@')[0]
        except Exception as e:
            # Fallback to short name from email (silently)
            email_to_name[email] = email.split('@')[0]
    
    return email_to_name


def get_user_working_hours(
    email: str,
    token: str = None
) -> Dict[str, Any]:
    """
    Get user's working hours configuration from mailbox settings.
    
    Args:
        email: User email address
        token: Access token
    
    Returns:
        Working hours information including timezone and time ranges
    """
    if token is None:
        token = get_access_token()
    
    # For other users, we can't access their mailbox settings directly
    # We'll use a default working hours (9:00-18:00) if we can't get their settings
    # Only the current user's settings can be retrieved
    
    url = f"{GRAPH_API_BASE}/me/mailboxSettings"
    
    try:
        response = requests.get(url, headers=get_headers(token))
        
        if response.status_code == 200:
            data = response.json()
            working_hours = data.get('workingHours', {})
            return {
                'email': email,
                'timezone': working_hours.get('timeZone', {}).get('name', 'UTC'),
                'daysOfWeek': working_hours.get('daysOfWeek', ['monday', 'tuesday', 'wednesday', 'thursday', 'friday']),
                'startTime': working_hours.get('startTime', '09:00:00'),
                'endTime': working_hours.get('endTime', '18:00:00')
            }
    except:
        pass
    
    # Default working hours if we can't retrieve
    return {
        'email': email,
        'timezone': 'UTC',
        'daysOfWeek': ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
        'startTime': '09:00:00',
        'endTime': '18:00:00'
    }


def suggest_meeting_times(
    attendees: List[str],
    display_timezone: str,
    duration_minutes: int = 60,
    start: str = None,
    end: str = None,
    top_n: int = 5,
    interval: int = 30,
    token: str = None
) -> Dict[str, Any]:
    """
    Suggest optimal meeting times based on attendee availability.
    
    Intelligently analyzes free/busy data and returns ranked time slots.
    
    Args:
        attendees: List of attendee email addresses
        duration_minutes: Required meeting duration (default 60)
        start: Search start datetime (must include timezone, or "now", default: now)
        end: Search end datetime (must include timezone, or "now", default: 7 days from now)
        timezone: (DEPRECATED - extracted from start/end) Fallback timezone
        top_n: Number of top slots to return (default 5)
        interval: Time slot interval in minutes (default 30)
        token: Access token
        display_timezone: Timezone for "now" calculation and display
    
    Returns:
        Dict with top_time_slots ranked by score, plus detailed availability info
    """
    if token is None:
        token = get_access_token()
    
    from zoneinfo import ZoneInfo
    
    # Handle defaults using "now" semantic
    if start is None:
        start = "now"
    if end is None:
        # Default to 7 days from now
        tz = ZoneInfo(display_timezone)
        end_dt = datetime.now(tz) + timedelta(days=7)
        offset = end_dt.strftime('%z')
        end = end_dt.strftime("%Y-%m-%dT%H:%M:%S") + f"{offset[:3]}:{offset[3:]}"
    
    # Parse times and extract timezones
    start_str, start_tz = parse_datetime_or_raise(start, 'start', display_timezone)
    end_str, end_tz = parse_datetime_or_raise(end, 'end', display_timezone)
    
    # Use extracted timezone (prefer start timezone)
    event_tz = start_tz or "UTC"
    
    # Create datetime objects for time calculations
    tz_obj = ZoneInfo(event_tz) if event_tz != "UTC" else ZoneInfo("UTC")
    start_dt = datetime.fromisoformat(start_str)
    start_dt = start_dt.replace(tzinfo=tz_obj)
    end_dt = datetime.fromisoformat(end_str)
    end_dt = end_dt.replace(tzinfo=tz_obj)
    
    url = f"{GRAPH_API_BASE}/me/calendar/getschedule"
    
    payload = {
        "schedules": attendees,
        "startTime": {
            "dateTime": start_str,
            "timeZone": event_tz
        },
        "endTime": {
            "dateTime": end_str,
            "timeZone": event_tz
        },
        "availabilityViewInterval": interval
    }
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code != 200:
        raise Exception(f"Failed to get suggestions: {response.status_code} - {response.text}")
    
    data = response.json()
    availability_data = data.get("value", [])
    
    # Track free status for each slot
    # Key: slot_index, Value: set of free attendee emails
    slot_free_status = {}  # {slot_index: set(free_emails)}
    slot_busy_status = {}  # {slot_index: [(email, status), ...]}
    
    for attendee_info in availability_data:
        email = attendee_info.get("scheduleId", "unknown")
        availability_view = attendee_info.get("availabilityView", "")
        schedule_items = attendee_info.get("scheduleItems", [])
        
        # Parse availability view string
        for i, status_code in enumerate(availability_view):
            if i not in slot_free_status:
                slot_free_status[i] = set()
                slot_busy_status[i] = []
            
            if status_code == "0":  # Free
                slot_free_status[i].add(email)
            else:
                status_map = {
                    "1": "Tentative",
                    "2": "Busy",
                    "3": "Out of Office",
                    "4": "Working Elsewhere",
                    "?": "Unknown"
                }
                slot_busy_status[i].append({
                    "email": email,
                    "status": status_map.get(status_code, "Unknown")
                })
    
    if not slot_free_status:
        return {
            "success": True,
            "search_range": {"start": start, "end": end},
            "timezone": display_timezone,
            "duration_minutes": duration_minutes,
            "total_attendees": len(attendees),
            "top_time_slots": [],
            "message": "No availability data found"
        }
    
    # Find continuous free slots that fit the meeting duration
    required_slots = max(1, duration_minutes // interval)
    total_attendees = len(attendees)
    
    def score_slot(free_count, total, has_tentative=False):
        """Score a time slot based on availability."""
        # Base score: percentage of free attendees (0-100)
        if total > 0:
            base_score = (free_count / total) * 100
        else:
            base_score = 100
        
        # Bonus for all attendees free (+30)
        all_free_bonus = 30 if free_count == total else 0
        
        # Penalty for tentative events (-10)
        tentative_penalty = -10 if has_tentative else 0
        
        return base_score + all_free_bonus + tentative_penalty
    
    meeting_slots = []
    max_slot_index = max(slot_free_status.keys())
    
    # Slide through slots to find continuous windows
    for i in range(max_slot_index + 1):
        if i + required_slots > max_slot_index + 1:
            break
        
        # Check if all required slots exist and are continuous
        slots_valid = True
        free_intersection = set(attendees)  # Start with all attendees
        has_tentative = False
        all_busy_info = []
        
        for j in range(required_slots):
            slot_idx = i + j
            if slot_idx not in slot_free_status:
                slots_valid = False
                break
            
            # Intersect free attendees
            free_intersection = free_intersection & slot_free_status[slot_idx]
            
            # Check for tentative status
            for busy_info in slot_busy_status.get(slot_idx, []):
                all_busy_info.append(busy_info)
                if busy_info["status"] == "Tentative":
                    has_tentative = True
        
        if not slots_valid:
            continue
        
        # Calculate score
        free_count = len(free_intersection)
        score = score_slot(free_count, total_attendees, has_tentative)
        
        # Calculate actual time
        slot_start = start_dt + timedelta(minutes=i * interval)
        slot_end = slot_start + timedelta(minutes=duration_minutes)
        
        # Build unavailable attendees list
        unavailable = []
        for email in attendees:
            if email not in free_intersection:
                # Find their status
                status = "Busy"
                for busy_info in all_busy_info:
                    if busy_info["email"] == email:
                        status = busy_info["status"]
                        break
                unavailable.append({"email": email, "status": status})
        
        meeting_slots.append({
            "start": slot_start.strftime("%Y-%m-%d %H:%M"),
            "end": slot_end.strftime("%Y-%m-%d %H:%M"),
            "score": round(score, 1),
            "free_attendees": list(free_intersection),
            "free_count": free_count,
            "total_attendees": total_attendees,
            "all_free": free_count == total_attendees,
            "unavailable_attendees": unavailable
        })
    
    # Sort by score (descending), then by start time
    meeting_slots.sort(key=lambda x: (-x["score"], x["start"]))
    
    # Return top N slots
    top_slots = meeting_slots[:top_n]
    
    # Add rank
    for i, slot in enumerate(top_slots):
        slot["rank"] = i + 1
    
    return {
        "success": True,
        "search_range": {"start": start, "end": end},
        "timezone": display_timezone,
        "duration_minutes": duration_minutes,
        "total_attendees": total_attendees,
        "top_time_slots": top_slots,
        "raw_availability": availability_data  # Include raw data for debugging
    }


# =============================================================================
# SEND EMAIL TO ATTENDEES
# =============================================================================

def send_meeting_email(
    event_id: str,
    subject: str,
    body: str,
    token: str = None
) -> bool:
    """
    Send an email to all attendees of a meeting.
    
    Args:
        event_id: Event ID
        subject: Email subject
        body: Email body
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    # Get event details to retrieve attendees
    event = get_event(event_id, token)
    
    attendees = event.get("attendees", [])
    recipient_emails = [
        a.get("emailAddress", {}).get("address")
        for a in attendees
        if a.get("emailAddress", {}).get("address")
    ]
    
    if not recipient_emails:
        raise Exception("No attendees found for this event")
    
    # Import email module
    from email_operations import send_email
    
    return send_email(
        to=recipient_emails,
        subject=subject,
        body=body,
        token=token
    )


# =============================================================================
# LIST CALENDARS
# =============================================================================

def list_calendars(token: str = None) -> List[Dict[str, Any]]:
    """List all calendars for the user."""
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/calendars"
    
    response = requests.get(url, headers=get_headers(token))
    
    if response.status_code != 200:
        raise Exception(f"Failed to list calendars: {response.status_code} - {response.text}")
    
    data = response.json()
    return data.get("value", [])


# =============================================================================
# DISPLAY HELPERS
# =============================================================================

def display_event_list(events: List[Dict], timezone: str):
    """Display a list of events in a readable format with timezone conversion."""
    from zoneinfo import ZoneInfo
    
    print(f"\n{'='*120}")
    
    # Parse timezone
    if timezone.startswith('+') or timezone.startswith('-'):
        tz = None
        tz_offset = timezone
    else:
        tz = ZoneInfo(timezone)
        tz_offset = None
    
    for i, event in enumerate(events, 1):
        start_dt_str = event.get('start', {}).get('dateTime', '')
        end_dt_str = event.get('end', {}).get('dateTime', '')
        
        if start_dt_str:
            # Handle Microsoft Graph date format with 7 decimal places
            if '.' in start_dt_str:
                start_dt_str = start_dt_str.split('.')[0]
            try:
                start_dt = datetime.fromisoformat(start_dt_str).replace(tzinfo=ZoneInfo('UTC'))
                
                # Convert to user timezone
                if tz_offset:
                    # Manual offset conversion
                    sign = 1 if tz_offset[0] == '+' else -1
                    hours = int(tz_offset[1:3])
                    minutes = int(tz_offset[4:6])
                    from datetime import timedelta
                    offset = timedelta(hours=sign*hours, minutes=sign*minutes)
                    start_local = start_dt + offset
                    start_str = f"{start_local.strftime('%Y-%m-%d %H:%M')} (UTC{tz_offset})"
                else:
                    start_local = start_dt.astimezone(tz)
                    start_str = f"{start_local.strftime('%Y-%m-%d %H:%M')} ({timezone})"
            except ValueError:
                start_str = start_dt_str[:16]
        else:
            start_str = ''
        
        if end_dt_str:
            if '.' in end_dt_str:
                end_dt_str = end_dt_str.split('.')[0]
            try:
                end_dt = datetime.fromisoformat(end_dt_str).replace(tzinfo=ZoneInfo('UTC'))
                
                # Convert to user timezone
                if tz_offset:
                    sign = 1 if tz_offset[0] == '+' else -1
                    hours = int(tz_offset[1:3])
                    minutes = int(tz_offset[4:6])
                    from datetime import timedelta
                    offset = timedelta(hours=sign*hours, minutes=sign*minutes)
                    end_local = end_dt + offset
                    end_str = end_local.strftime('%H:%M')
                else:
                    end_local = end_dt.astimezone(tz)
                    end_str = end_local.strftime('%H:%M')
            except ValueError:
                end_str = end_dt_str[11:16]
        else:
            end_str = ''
        
        subject = event.get('subject', '(No Subject)')
        location = event.get('location', {}).get('displayName', '')
        
        # Get organizer
        organizer = event.get('organizer', {}).get('emailAddress', {})
        organizer_name = organizer.get('name', organizer.get('address', ''))
        
        # Get response status
        response_status = event.get('responseStatus', {})
        response = response_status.get('response', 'none')
        status_map = {
            'accepted': '✅ Accepted',
            'tentativelyAccepted': '❓ Tentative',
            'declined': '❌ Declined',
            'notResponded': '⏳ Not Responded',
            'organizer': '👤 Organizer',
            'none': '⏳ Not Responded'
        }
        status_str = status_map.get(response, response)
        
        # Get attendees
        attendees = event.get('attendees', [])
        attendee_names = []
        for att in attendees[:5]:  # Show first 5 attendees
            email_addr = att.get('emailAddress', {})
            name = email_addr.get('name', email_addr.get('address', ''))
            attendee_names.append(name)
        
        attendee_str = ', '.join(attendee_names)
        if len(attendees) > 5:
            attendee_str += f' (+{len(attendees)-5} more)'
        
        print(f"\n{i}. {subject}")
        
        # For not responded meetings, check availability and show recommendation
        if response in ['notResponded', 'none'] and start_dt_str and end_dt_str:
            try:
                # Check availability for this time slot
                from pathlib import Path
                import sys
                sys.path.insert(0, str(Path(__file__).parent))
                
                # Get current user's email
                token = get_access_token()
                headers = {
                    'Authorization': f'Bearer {token}',
                    'Content-Type': 'application/json'
                }
                response_data = requests.get(
                    'https://graph.microsoft.com/v1.0/me',
                    headers=headers
                )
                my_email = response_data.json().get('mail', response_data.json().get('userPrincipalName', ''))
                
                # Check availability
                avail_result = get_availability(
                    emails=[my_email],
                    start=start_dt.isoformat(),
                    end=end_dt.isoformat(),
                    display_timezone=timezone
                )
                
                # Parse availability
                schedules = avail_result.get('value', [])
                availability_status = "❓ Unknown"
                recommendation = ""
                
                if schedules:
                    schedule = schedules[0]
                    schedule_items = schedule.get('scheduleItems', [])
                    
                    # Get this event's subject
                    current_subject = subject.strip()
                    
                    # Check if there are any OTHER events (excluding this one by subject)
                    other_events = []
                    for item in schedule_items:
                        item_subject = item.get('subject', '').strip()
                        
                        # Skip if this is the same event (match by subject)
                        if item_subject.lower() == current_subject.lower():
                            continue
                        
                        other_events.append(item)
                    
                    if not other_events:
                        availability_status = "✅ Free"
                        recommendation = " → 建议接受"
                    else:
                        availability_status = f"❌ Busy ({len(other_events)} conflict{'s' if len(other_events) > 1 else ''})"
                        recommendation = " → 不建议接受"
                
                print(f"   Time: {start_str} - {end_str} {availability_status}{recommendation}")
            except Exception as e:
                # If availability check fails, just show time without recommendation
                print(f"   Time: {start_str} - {end_str}")
        else:
            # For accepted/declined/tentative meetings, just show time
            print(f"   Time: {start_str} - {end_str}")
        
        print(f"   Status: {status_str}")
        if organizer_name:
            print(f"   Organizer: {organizer_name}")
        if location:
            print(f"   Location: {location}")
        if attendee_str:
            print(f"   Attendees: {attendee_str}")
    
    print(f"\n{'='*120}")
    print(f"Total: {len(events)} events")


def display_event(event: Dict):
    """Display a single event in detail."""
    print(f"\n{'='*80}")
    print(f"Subject: {event.get('subject', '(No Subject)')}")
    print(f"Organizer: {event.get('organizer', {}).get('emailAddress', {})}")
    print(f"Start: {event.get('start', {})}")
    print(f"End: {event.get('end', {})}")
    print(f"Location: {event.get('location', {}).get('displayName', '')}")
    print(f"All Day: {event.get('isAllDay', False)}")
    print(f"Online Meeting: {event.get('isOnlineMeeting', False)}")
    
    attendees = event.get('attendees', [])
    if attendees:
        print(f"\nAttendees:")
        for a in attendees:
            email_info = a.get('emailAddress', {})
            print(f"  - {email_info.get('name', '')} <{email_info.get('address', '')}> ({a.get('type', 'required')})")
    
    body = event.get('body', {}).get('content', '')
    if body:
        print(f"\nDescription:\n{body[:500]}{'...' if len(body) > 500 else ''}")
    
    print(f"\n{'='*80}")


def display_availability(data: Dict, timezone: str, query_start: str = None, query_end: str = None):
    """Display availability information with timezone conversion and multi-person comparison."""
    from zoneinfo import ZoneInfo
    from datetime import datetime, timedelta
    
    print(f"\n{'='*120}")
    print("Availability Information")
    print(f"{'='*120}")
    
    # Parse timezone
    if timezone.startswith('+') or timezone.startswith('-'):
        # Offset format like +08:00
        tz = ZoneInfo('UTC')
        tz_offset = timezone
    else:
        # Named timezone like Asia/Shanghai
        tz = ZoneInfo(timezone)
        tz_offset = None
    
    # Calculate query start time slot index (for mapping availabilityView indices to actual time slots)
    query_start_slot = 0
    if query_start:
        try:
            # Parse query start time
            query_start_dt = datetime.fromisoformat(query_start.replace('Z', '+00:00'))
            if query_start_dt.tzinfo is None:
                query_start_dt = query_start_dt.replace(tzinfo=tz)
            else:
                query_start_dt = query_start_dt.astimezone(tz)
            # Calculate slot index from midnight (each slot is 30 minutes)
            query_start_slot = query_start_dt.hour * 2 + (1 if query_start_dt.minute >= 30 else 0)
        except:
            query_start_slot = 0
    
    schedules = data.get("value", [])
    
    # Extract display names from schedule data (more efficient than separate API calls)
    email_to_name = {}
    for schedule in schedules:
        email = schedule.get("scheduleId", "Unknown")
        # Try to get display name from availability response first
        availability_view = schedule.get("availabilityView", "")
        # Check if there's a display name in the response
        # Microsoft Graph may include it in some responses, but typically scheduleId is just the email
        # So we'll use get_users_info as fallback, but with better error handling
        email_to_name[email] = email.split('@')[0]  # Default to short name
    
    # Try to get better names from get_users_info (if API quota allows)
    try:
        all_emails = [schedule.get("scheduleId", "Unknown") for schedule in schedules]
        better_names = get_users_info(all_emails)
        # Only update if we got actual names (not just fallbacks)
        for email, name in better_names.items():
            if name != email.split('@')[0]:  # If it's not just the email username
                email_to_name[email] = name
    except:
        pass  # Use the default names we already set
    
    # Extract time range from availability data to fetch full event details
    # We'll fetch events for the current user to get complete information
    time_range_start = None
    time_range_end = None
    
    for schedule in schedules:
        schedule_items = schedule.get("scheduleItems", [])
        for item in schedule_items:
            start_info = item.get('start', {})
            end_info = item.get('end', {})
            start_dt_str = start_info.get('dateTime', '')
            end_dt_str = end_info.get('dateTime', '')
            
            if start_dt_str:
                try:
                    if '.' in start_dt_str:
                        start_dt_str = start_dt_str.split('.')[0]
                    start_dt = datetime.fromisoformat(start_dt_str)
                    if time_range_start is None or start_dt < time_range_start:
                        time_range_start = start_dt
                except:
                    pass
            
            if end_dt_str:
                try:
                    if '.' in end_dt_str:
                        end_dt_str = end_dt_str.split('.')[0]
                    end_dt = datetime.fromisoformat(end_dt_str)
                    if time_range_end is None or end_dt > time_range_end:
                        time_range_end = end_dt
                except:
                    pass
    
    # Fetch full event details for the time range
    full_events = {}
    if time_range_start and time_range_end:
        try:
            # Add buffer to time range
            time_range_start = time_range_start - timedelta(hours=1)
            time_range_end = time_range_end + timedelta(hours=1)
            
            events = list_events(
                start=time_range_start.strftime("%Y-%m-%dT%H:%M:%S"),
                end=time_range_end.strftime("%Y-%m-%dT%H:%M:%S"),
                display_timezone=timezone
            )
            
            # Index events by subject and start time for quick lookup
            for event in events:
                subject = event.get('subject', '').strip()
                start_info = event.get('start', {})
                start_dt_str = start_info.get('dateTime', '')
                if start_dt_str and subject:
                    try:
                        if '.' in start_dt_str:
                            start_dt_str = start_dt_str.split('.')[0]
                        start_dt = datetime.fromisoformat(start_dt_str.replace('Z', ''))
                        # Use subject + start time as key
                        key = f"{subject}|{start_dt.strftime('%Y-%m-%d %H:%M')}"
                        full_events[key] = event
                    except:
                        pass
        except Exception as e:
            # If fetching full events fails, continue with limited info
            pass
    
    # If multiple people, show comparison table first
    if len(schedules) > 1:
        print(f"\n📊 Multi-Person Availability Comparison")
        print(f"{'='*120}")
        
        # Get working hours for each person from the schedule data
        working_hours_info = {}
        for schedule in schedules:
            email = schedule.get("scheduleId", "Unknown")
            working_hours = schedule.get("workingHours", {})
            
            # Extract working hours info
            if working_hours:
                tz_info = working_hours.get('timeZone', {})
                timezone_name = tz_info.get('name', 'UTC') if isinstance(tz_info, dict) else str(tz_info)
                start_time = working_hours.get('startTime', '09:00:00')
                end_time = working_hours.get('endTime', '18:00:00')
            else:
                # Default working hours
                timezone_name = 'UTC'
                start_time = '09:00:00'
                end_time = '18:00:00'
            
            working_hours_info[email] = {
                'timezone': timezone_name,
                'start': start_time,
                'end': end_time
            }
        
        # Display working hours for each person (converted to display timezone)
        print(f"\n⏰ Working Hours (in {timezone}):")
        for email, wh in working_hours_info.items():
            display_name = email_to_name.get(email, email.split('@')[0])
            wh_tz_name = wh['timezone']
            wh_start_str = wh['start']
            wh_end_str = wh['end']
            
            try:
                from zoneinfo import ZoneInfo
                from datetime import datetime
                
                # Map Windows timezone names to IANA timezone names
                tz_map = {
                    'China Standard Time': 'Asia/Shanghai',
                    'India Standard Time': 'Asia/Kolkata',
                    'Singapore Standard Time': 'Asia/Singapore',
                    'Pacific Standard Time': 'America/Los_Angeles',
                    'Eastern Standard Time': 'America/New_York',
                    'GMT Standard Time': 'Europe/London',
                    'UTC': 'UTC'
                }
                
                iana_tz_name = tz_map.get(wh_tz_name, wh_tz_name)
                wh_start_time = datetime.strptime(wh_start_str[:5], '%H:%M').time()
                wh_end_time = datetime.strptime(wh_end_str[:5], '%H:%M').time()
                
                today = datetime.now(ZoneInfo(iana_tz_name)).date()
                wh_start_dt = datetime.combine(today, wh_start_time, tzinfo=ZoneInfo(iana_tz_name))
                wh_end_dt = datetime.combine(today, wh_end_time, tzinfo=ZoneInfo(iana_tz_name))
                
                display_tz = ZoneInfo(timezone)
                wh_start_display = wh_start_dt.astimezone(display_tz)
                wh_end_display = wh_end_dt.astimezone(display_tz)
                
                print(f"  {display_name}: {wh_start_display.strftime('%H:%M')} - {wh_end_display.strftime('%H:%M')} (originally {wh_start_str[:5]}-{wh_end_str[:5]} {wh_tz_name})")
            except Exception as e:
                # Fallback to original display
                print(f"  {display_name}: {wh_start_str[:5]} - {wh_end_str[:5]} ({wh_tz_name})")
        
        # Get all availability views
        all_views = []
        all_display_names = []
        for schedule in schedules:
            email = schedule.get("scheduleId", "Unknown")
            availability = schedule.get("availabilityView", "")
            display_name = email_to_name.get(email, email.split('@')[0])
            all_display_names.append(display_name)
            all_views.append(availability)
        
        # Find common free slots
        if all_views and all_views[0]:
            view_length = len(all_views[0])
            common_free_slots = []
            
            for i in range(view_length):
                all_free = all(view[i] == '0' if i < len(view) else False for view in all_views)
                if all_free:
                    common_free_slots.append(i)
            
            # Group consecutive slots and filter by working hours (9:00-18:00)
            if common_free_slots:
                # Filter slots within working hours (9:00-18:00 = slots 18-36)
                # Each slot is 30 minutes, so 9:00 = slot 18, 18:00 = slot 36
                working_hours_slots = [slot for slot in common_free_slots if 18 <= slot < 36]
                
                if working_hours_slots:
                    print(f"\n✅ Common Free Time Slots (Working Hours 9:00-18:00):")
                    slot_groups = []
                    current_group = [working_hours_slots[0]]
                    
                    for slot in working_hours_slots[1:]:
                        if slot == current_group[-1] + 1:
                            current_group.append(slot)
                        else:
                            slot_groups.append(current_group)
                            current_group = [slot]
                    slot_groups.append(current_group)
                    
                    # Display time ranges (assuming each slot is 30 minutes)
                    time_slots_info = []
                    for group in slot_groups:
                        start_mins = group[0] * 30
                        end_mins = (group[-1] + 1) * 30
                        start_time = f"{start_mins // 60:02d}:{start_mins % 60:02d}"
                        end_time = f"{end_mins // 60:02d}:{end_mins % 60:02d}"
                        duration = len(group) * 30
                        time_slots_info.append((start_time, end_time, duration, start_mins // 60))
                        print(f"  • {start_time} - {end_time} ({duration} min)")
                    
                else:
                    print(f"\n❌ No common free time slots found during working hours (9:00-18:00)")
            else:
                print(f"\n❌ No common free time slots found")
        
        # Show visual comparison with working hours overlay
        print(f"\n📅 Time Slot Comparison (each character = 30 min, displayed in {timezone}):")
        print(f"   Legend: 0=Free ✅  1=Tentative ❓  2=Busy 🔴  3=OOF 🏖️  4=WorkElsewhere 💼  ⬛=Out of Office Hours")
        print()
        
        max_name_len = max(len(name) for name in all_display_names)
        
        for display_name, view in zip(all_display_names, all_views):
            # Get email for this display name (reverse lookup)
            email = None
            for e, n in email_to_name.items():
                if n == display_name:
                    email = e
                    break
            if email is None:
                # Fallback: try to find by matching short name
                for e in email_to_name.keys():
                    if e.split('@')[0] == display_name:
                        email = e
                        break
            # Get working hours for this person (only if we found the email)
            wh = working_hours_info.get(email, {}) if email else {}
            wh_start_str = wh.get('start', '09:00:00')
            wh_end_str = wh.get('end', '18:00:00')
            wh_tz_name = wh.get('timezone', 'UTC')
            
            # Convert working hours to display timezone
            try:
                from zoneinfo import ZoneInfo
                from datetime import datetime, time
                
                # Map Windows timezone names to IANA timezone names
                tz_map = {
                    'China Standard Time': 'Asia/Shanghai',
                    'India Standard Time': 'Asia/Kolkata',
                    'Singapore Standard Time': 'Asia/Singapore',
                    'Pacific Standard Time': 'America/Los_Angeles',
                    'Eastern Standard Time': 'America/New_York',
                    'GMT Standard Time': 'Europe/London',
                    'UTC': 'UTC'
                }
                
                # Get IANA timezone name
                iana_tz_name = tz_map.get(wh_tz_name, wh_tz_name)
                
                # Parse working hours
                wh_start_time = datetime.strptime(wh_start_str[:5], '%H:%M').time()
                wh_end_time = datetime.strptime(wh_end_str[:5], '%H:%M').time()
                
                # Create datetime objects for today in person's timezone
                today = datetime.now(ZoneInfo(iana_tz_name)).date()
                wh_start_dt = datetime.combine(today, wh_start_time, tzinfo=ZoneInfo(iana_tz_name))
                wh_end_dt = datetime.combine(today, wh_end_time, tzinfo=ZoneInfo(iana_tz_name))
                
                # Convert to display timezone
                display_tz = ZoneInfo(timezone)
                wh_start_display = wh_start_dt.astimezone(display_tz)
                wh_end_display = wh_end_dt.astimezone(display_tz)
                
                # Calculate slot indices (each slot is 30 minutes, starting from 00:00)
                start_slot = wh_start_display.hour * 2 + (1 if wh_start_display.minute >= 30 else 0)
                end_slot = wh_end_display.hour * 2 + (1 if wh_end_display.minute >= 30 else 0)
                
                # Build the colored view with out-of-office markers
                colored_chars = []
                for i, char in enumerate(view):
                    # Calculate actual slot index (i is relative to query start time)
                    actual_slot = query_start_slot + i
                    
                    # Check if this slot is within working hours
                    if actual_slot < start_slot or actual_slot >= end_slot:
                        # Out of office hours
                        colored_chars.append('⬛')
                    else:
                        # Within working hours, use normal coloring
                        if char == '0':
                            colored_chars.append('✅')
                        elif char == '1':
                            colored_chars.append('❓')
                        elif char == '2':
                            colored_chars.append('🔴')
                        elif char == '3':
                            colored_chars.append('🏖️')
                        elif char == '4':
                            colored_chars.append('💼')
                        else:
                            colored_chars.append(char)
                
                colored_view = ''.join(colored_chars)
                
                # Format working hours in display timezone
                wh_display_str = f"{wh_start_display.strftime('%H:%M')}-{wh_end_display.strftime('%H:%M')} {timezone}"
                
            except Exception as e:
                # Fallback to simple coloring if timezone conversion fails
                colored_view = view.replace('0', '✅').replace('1', '❓').replace('2', '🔴').replace('3', '🏖️').replace('4', '💼')
                wh_display_str = f"{wh_start_str[:5]}-{wh_end_str[:5]} {wh_tz_name}"
            
            # Extract short name from email (e.g., "luomn" from "luomn@cn.ibm.com")
            print(f"  {display_name:<{max_name_len}} | {colored_view}  [{wh_display_str}]")
        
        # Add smart meeting suggestions using suggest_meeting_times function
        if len(schedules) > 1 and all_views:
            try:
                # Use suggest_meeting_times for intelligent recommendations
                suggestions = suggest_meeting_times(
                    attendees=all_emails,
                    display_timezone=timezone,
                    duration_minutes=30,  # Default 30 minute meeting
                    start=query_start,
                    end=query_end,
                    top_n=5,
                    interval=30
                )
                
                if suggestions.get("top_time_slots"):
                    print(f"\n💡 Smart Meeting Suggestions (30-minute meetings):")
                    for slot in suggestions["top_time_slots"]:
                        all_free_indicator = "✓ All free" if slot["all_free"] else f"{slot['free_count']}/{slot['total_attendees']} available"
                        print(f"  #{slot['rank']}. {slot['start']} - {slot['end']} | Score: {slot['score']} | {all_free_indicator}")
                        
                        # Show unavailable attendees if any
                        if slot["unavailable_attendees"] and not slot["all_free"]:
                            unavailable_names = []
                            for u in slot["unavailable_attendees"][:3]:  # Show first 3
                                name = email_to_name.get(u['email'], u['email'].split('@')[0])
                                unavailable_names.append(f"{name} ({u['status']})")
                            unavailable_str = ", ".join(unavailable_names)
                            if len(slot["unavailable_attendees"]) > 3:
                                unavailable_str += f" +{len(slot['unavailable_attendees'])-3} more"
                            print(f"       Unavailable: {unavailable_str}")
                else:
                    print(f"\n💡 No suitable meeting times found in the given range")
            except Exception as e:
                # If suggestion fails, continue without it (don't break availability display)
                print(f"\n💡 Smart Meeting Suggestions: Unable to generate (error: {str(e)})")
        
        print(f"\n{'='*120}")
    
    # For multi-person availability, we only show the comparison view
    # Individual event details are not shown to keep output concise
    if len(schedules) == 1:
        # For single person, show detailed information
        schedule = schedules[0]
        email = schedule.get("scheduleId", "Unknown")
        availability = schedule.get("availabilityView", "No data")
        
        # Determine overall recommendation based on availability view
        overall_recommendation = ""
        if availability and availability != "No data":
            if '2' in availability or '3' in availability:
                overall_recommendation = " → ❌ 不建议接受 (有时间冲突)"
            elif '1' in availability:
                overall_recommendation = " → ❓ 待定 (有待定会议)"
            elif '4' in availability:
                overall_recommendation = " → 💼 谨慎考虑 (Working Elsewhere)"
            elif availability == '0' * len(availability):
                overall_recommendation = " → ✅ 建议接受 (完全空闲)"
        
        print(f"\n{email}:")
        print(f"  Availability View: {availability}{overall_recommendation}")
        
        schedule_items = schedule.get("scheduleItems", [])
        if not schedule_items:
            print(f"  No scheduled items")
        else:
            for i, item in enumerate(schedule_items, 1):
                subject = item.get('subject', 'Busy')
                status = item.get('status', 'busy')
                
                # Parse start and end times
                start_info = item.get('start', {})
                end_info = item.get('end', {})
                
                start_dt_str = start_info.get('dateTime', '')
                end_dt_str = end_info.get('dateTime', '')
                
                if start_dt_str and end_dt_str:
                    try:
                        # Parse datetime
                        if '.' in start_dt_str:
                            start_dt_str = start_dt_str.split('.')[0]
                        if '.' in end_dt_str:
                            end_dt_str = end_dt_str.split('.')[0]
                        
                        start_dt = datetime.fromisoformat(start_dt_str).replace(tzinfo=ZoneInfo('UTC'))
                        end_dt = datetime.fromisoformat(end_dt_str).replace(tzinfo=ZoneInfo('UTC'))
                        
                        # Convert to user timezone
                        if tz_offset:
                            start_local = start_dt
                            end_local = end_dt
                            time_str = f"{start_local.strftime('%Y-%m-%d %H:%M')} - {end_local.strftime('%H:%M')} UTC"
                            if tz_offset != '+00:00':
                                sign = 1 if tz_offset[0] == '+' else -1
                                hours = int(tz_offset[1:3])
                                minutes = int(tz_offset[4:6])
                                offset = timedelta(hours=sign*hours, minutes=sign*minutes)
                                start_local = start_dt + offset
                                end_local = end_dt + offset
                                time_str = f"{start_local.strftime('%Y-%m-%d %H:%M')} - {end_local.strftime('%H:%M')} (UTC{tz_offset})"
                        else:
                            start_local = start_dt.astimezone(tz)
                            end_local = end_dt.astimezone(tz)
                            time_str = f"{start_local.strftime('%Y-%m-%d %H:%M')} - {end_local.strftime('%H:%M')} ({timezone})"
                        
                        # Duration
                        duration = end_dt - start_dt
                        duration_mins = int(duration.total_seconds() / 60)
                        
                        # Status emoji
                        status_map = {
                            'free': '✅',
                            'tentative': '❓',
                            'busy': '🔴',
                            'oof': '🏖️',
                            'workingElsewhere': '💼'
                        }
                        status_emoji = status_map.get(status, '⏰')
                        
                        # Try to find full event details
                        event_key = f"{subject}|{start_dt.strftime('%Y-%m-%d %H:%M')}"
                        full_event = full_events.get(event_key)
                        
                        # Recommendation based on status
                        if status == 'free':
                            recommendation = "✅ 建议接受 (Free - 可以接受)"
                        elif status == 'tentative':
                            recommendation = "❓ 待定 (Tentative - 已有待定会议)"
                        elif status == 'busy':
                            recommendation = "❌ 不建议接受 (Busy - 时间冲突)"
                        elif status == 'oof':
                            recommendation = "🏖️ 不建议接受 (Out of Office)"
                        elif status == 'workingElsewhere':
                            recommendation = "💼 谨慎考虑 (Working Elsewhere)"
                        else:
                            recommendation = ""
                        
                        print(f"\n  {i}. {status_emoji} {subject}")
                        print(f"     Time: {time_str}")
                        print(f"     Status: {status}")
                        
                        # Display full event details if available
                        if full_event:
                            organizer = full_event.get('organizer', {}).get('emailAddress', {})
                            organizer_name = organizer.get('name', organizer.get('address', 'Unknown'))
                            print(f"     Organizer: {organizer_name}")
                            
                            location = full_event.get('location', {})
                            if location:
                                location_name = location.get('displayName', '')
                                if location_name:
                                    print(f"     Location: {location_name}")
                            
                            attendees = full_event.get('attendees', [])
                            if attendees:
                                attendee_names = []
                                for att in attendees[:5]:
                                    email_addr = att.get('emailAddress', {})
                                    name = email_addr.get('name', email_addr.get('address', ''))
                                    if name:
                                        attendee_names.append(name)
                                if attendee_names:
                                    attendee_str = ', '.join(attendee_names)
                                    if len(attendees) > 5:
                                        attendee_str += f" (+{len(attendees) - 5} more)"
                                    print(f"     Attendees: {attendee_str}")
                        
                        if recommendation:
                            print(f"     {recommendation}")
                        
                    except Exception as e:
                        print(f"\n  {i}. {subject}")
                        print(f"     Start: {start_dt_str}")
                        print(f"     End: {end_dt_str}")
                        print(f"     (Error parsing time: {e})")
                else:
                    print(f"\n  {i}. {subject}")
                    print(f"     Status: {status}")
    
    print(f"\n{'='*120}")


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Microsoft Graph Calendar Operations")
    subparsers = parser.add_subparsers(dest="command", required=True)
    
    # Global --json flag
    parser.add_argument("--json", action="store_true", help="Output in JSON format")
    
    # List command
    list_parser = subparsers.add_parser("list", help="List calendar events")
    list_parser.add_argument("--calendar", help="Calendar ID")
    list_parser.add_argument("--start", help="Start datetime (format: '2026-03-26T12:00:00' or 'now')")
    list_parser.add_argument("--end", help="End datetime (format: '2026-03-26T12:00:00' or 'now')")
    list_parser.add_argument("--limit", type=int, default=25, help="Max events")
    list_parser.add_argument("--timezone", required=True, help="Display timezone (e.g., 'Asia/Shanghai', 'UTC')")
    
    # Get command
    get_parser = subparsers.add_parser("get", help="Get an event")
    get_parser.add_argument("event_id", help="Event ID")
    
    # Create command
    create_parser = subparsers.add_parser("create", help="Create an event")
    create_parser.add_argument("--subject", required=True, help="Event subject")
    create_parser.add_argument("--start", required=True, help="Start datetime (format: '2026-03-26T12:00:00' or 'now')")
    create_parser.add_argument("--end", required=True, help="End datetime (format: '2026-03-26T12:00:00' or 'now')")
    create_parser.add_argument("--timezone", required=True, help="Display timezone (e.g., 'Asia/Shanghai', 'UTC')")
    create_parser.add_argument("--body", help="Event description")
    create_parser.add_argument("--location", help="Location")
    create_parser.add_argument("--required", help="Required attendee emails (comma-separated)")
    create_parser.add_argument("--optional", help="Optional attendee emails (comma-separated)")
    create_parser.add_argument("--all-day", action="store_true", help="All day event")
    create_parser.add_argument("--no-teams", action="store_false", dest="teams", default=True, help="Disable Teams meeting (Teams is enabled by default)")
    
    # Update command
    update_parser = subparsers.add_parser("update", help="Update an event")
    update_parser.add_argument("event_id", help="Event ID")
    update_parser.add_argument("--subject", help="New subject")
    update_parser.add_argument("--start", help="New start datetime (format: '2026-03-26T12:00:00' or 'now')")
    update_parser.add_argument("--end", help="New end datetime (format: '2026-03-26T12:00:00' or 'now')")
    update_parser.add_argument("--body", help="New description")
    update_parser.add_argument("--location", help="New location")
    update_parser.add_argument("--required", help="Required attendee emails (comma-separated)")
    update_parser.add_argument("--optional", help="Optional attendee emails (comma-separated)")
    update_parser.add_argument("--timezone", help="Display timezone (e.g., 'Asia/Shanghai', 'UTC')")
    
    # Delete command
    delete_parser = subparsers.add_parser("delete", help="Soft delete an event (move to Deleted Items)")
    delete_parser.add_argument("event_id", help="Event ID")
    delete_parser.add_argument("--permanent", action="store_true", help="Permanently delete (cannot recover)")
    
    # Availability command
    avail_parser = subparsers.add_parser("availability", help="Get availability")
    avail_parser.add_argument("--emails", required=True, help="Email addresses (comma-separated)")
    avail_parser.add_argument("--start", required=True, help="Start datetime (format: '2026-03-26T12:00:00')")
    avail_parser.add_argument("--end", required=True, help="End datetime (format: '2026-03-26T12:00:00' or 'now')")
    avail_parser.add_argument("--timezone", required=True, help="Display timezone (e.g., 'Asia/Shanghai', 'UTC')")
    
    # Accept command
    accept_parser = subparsers.add_parser("accept", help="Accept an event invitation")
    accept_parser.add_argument("event_id", help="Event ID")
    accept_parser.add_argument("--comment", help="Optional comment")
    accept_parser.add_argument("--no-send", action="store_true", help="Don't send response to organizer")
    
    # Decline command
    decline_parser = subparsers.add_parser("decline", help="Decline an event invitation")
    decline_parser.add_argument("event_id", help="Event ID")
    decline_parser.add_argument("--comment", help="Optional comment")
    decline_parser.add_argument("--no-send", action="store_true", help="Don't send response to organizer")
    
    # Tentatively accept command
    tentative_parser = subparsers.add_parser("tentative", help="Tentatively accept an event invitation")
    tentative_parser.add_argument("event_id", help="Event ID")
    tentative_parser.add_argument("--comment", help="Optional comment")
    tentative_parser.add_argument("--no-send", action="store_true", help="Don't send response to organizer")
    
    # Cancel command (organizer only)
    cancel_parser = subparsers.add_parser("cancel", help="Cancel an event (organizer only, sends notifications)")
    cancel_parser.add_argument("event_id", help="Event ID")
    cancel_parser.add_argument("--comment", help="Cancellation message to attendees")
    
    # Forward command
    forward_parser = subparsers.add_parser("forward", help="Forward an event to new recipients")
    forward_parser.add_argument("event_id", help="Event ID")
    forward_parser.add_argument("--to", required=True, dest="to_emails", help="Email addresses (comma-separated)")
    forward_parser.add_argument("--comment", help="Optional message")
    
    # Propose new time command
    propose_parser = subparsers.add_parser("propose", help="Propose a new meeting time")
    propose_parser.add_argument("event_id", help="Event ID")
    propose_parser.add_argument("--start", required=True, help="Proposed new start datetime (format: '2026-03-26T12:00:00' or 'now')")
    propose_parser.add_argument("--end", required=True, help="Proposed new end datetime (format: '2026-03-26T12:00:00' or 'now')")
    propose_parser.add_argument("--timezone", required=True, help="Display timezone (e.g., 'Asia/Shanghai', 'UTC')")
    propose_parser.add_argument("--comment", help="Optional message to organizer")
    
    # List calendars command
    subparsers.add_parser("calendars", help="List all calendars")
    
    args = parser.parse_args()
    
    try:
        if args.command == "list":
            events = list_events(
                calendar_id=args.calendar,
                start=args.start,
                end=args.end,
                limit=args.limit,
                display_timezone=args.timezone
            )
            if args.json:
                print(json.dumps({"success": True, "events": events, "total": len(events)}, indent=2, default=str))
            else:
                display_event_list(events, timezone=args.timezone)
        
        elif args.command == "get":
            event = get_event(args.event_id)
            if args.json:
                print(json.dumps({"success": True, "event": event}, indent=2, default=str))
            else:
                display_event(event)
        
        elif args.command == "create":
            attendees = []
            
            # Process --required
            if args.required:
                attendees.extend([{"email": e, "type": "required"} for e in parse_email_list(args.required)])
            
            # Process --optional
            if args.optional:
                attendees.extend([{"email": e, "type": "optional"} for e in parse_email_list(args.optional)])
            
            # If no attendees specified, set to None
            if not attendees:
                attendees = None
            
            event = create_event(
                subject=args.subject,
                start=args.start,
                end=args.end,
                display_timezone=args.timezone,
                body=args.body,
                location=args.location,
                attendees=attendees,
                is_all_day=args.all_day,
                is_online_meeting=args.teams
            )
            if args.json:
                result = {"success": True, "event": event, "eventId": event.get('id')}
                if args.teams and event.get('onlineMeeting'):
                    result["teamsLink"] = event['onlineMeeting'].get('joinUrl')
                print(json.dumps(result, indent=2, default=str))
            else:
                print(f"✓ Event created: {event.get('id')}")
                if args.teams and event.get('onlineMeeting'):
                    print(f"  Teams Link: {event['onlineMeeting'].get('joinUrl')}")
        
        elif args.command == "update":
            attendees = None
            
            # Process attendees only if any are specified
            if args.required or args.optional:
                attendees = []
                
                # Process --required
                if args.required:
                    attendees.extend([{"email": e, "type": "required"} for e in parse_email_list(args.required)])
                
                # Process --optional
                if args.optional:
                    attendees.extend([{"email": e, "type": "optional"} for e in parse_email_list(args.optional)])
            
            event = update_event(
                event_id=args.event_id,
                display_timezone=args.timezone,
                subject=args.subject,
                start=args.start,
                end=args.end,
                body=args.body,
                location=args.location,
                attendees=attendees
            )
            if args.json:
                print(json.dumps({"success": True, "event": event}, indent=2, default=str))
            else:
                print(f"✓ Event updated")
        
        elif args.command == "delete":
            delete_event(args.event_id, permanent=args.permanent)
            if args.json:
                msg = "Event permanently deleted" if args.permanent else "Event moved to Deleted Items"
                print(json.dumps({"success": True, "message": msg}))
            else:
                msg = "✓ Event permanently deleted" if args.permanent else "✓ Event moved to Deleted Items (soft delete)"
                print(msg)
        
        elif args.command == "availability":
            # Validate that start is not "now"
            if args.start.lower() == "now":
                raise ValueError("--start cannot be 'now' for availability checks. Please provide a specific datetime like '2026-03-26T12:00:00'")
            
            data = get_availability(
                emails=parse_email_list(args.emails),
                start=args.start,
                end=args.end,
                display_timezone=args.timezone
            )
            if args.json:
                print(json.dumps({"success": True, "availability": data}, indent=2, default=str))
            else:
                display_availability(data, timezone=args.timezone, query_start=args.start, query_end=args.end)
        
        elif args.command == "calendars":
            calendars = list_calendars()
            if args.json:
                print(json.dumps({"success": True, "calendars": calendars, "total": len(calendars)}, indent=2, default=str))
            else:
                print(f"\nCalendars ({len(calendars)}):")
                for cal in calendars:
                    print(f"  - {cal.get('name', 'Unknown')} (ID: {cal.get('id')})")
        
        elif args.command == "accept":
            accept_event(
                event_id=args.event_id,
                comment=args.comment,
                send_response=not args.no_send
            )
            if args.json:
                print(json.dumps({"success": True, "message": "Event accepted"}))
            else:
                print("✓ Event accepted")
        
        elif args.command == "decline":
            decline_event(
                event_id=args.event_id,
                comment=args.comment,
                send_response=not args.no_send
            )
            if args.json:
                print(json.dumps({"success": True, "message": "Event declined"}))
            else:
                print("✓ Event declined")
        
        elif args.command == "tentative":
            tentatively_accept_event(
                event_id=args.event_id,
                comment=args.comment,
                send_response=not args.no_send
            )
            if args.json:
                print(json.dumps({"success": True, "message": "Event tentatively accepted"}))
            else:
                print("✓ Event tentatively accepted")
        
        elif args.command == "cancel":
            cancel_event(
                event_id=args.event_id,
                comment=args.comment
            )
            if args.json:
                print(json.dumps({"success": True, "message": "Event cancelled, notifications sent to attendees"}))
            else:
                print("✓ Event cancelled, notifications sent to attendees")
        
        elif args.command == "forward":
            forward_event(
                event_id=args.event_id,
                to_emails=parse_email_list(args.to_emails),
                comment=args.comment
            )
            if args.json:
                print(json.dumps({"success": True, "message": f"Event forwarded to {len(parse_email_list(args.to_emails))} recipient(s)"}))
            else:
                print(f"✓ Event forwarded to {len(parse_email_list(args.to_emails))} recipient(s)")
        
        elif args.command == "propose":
            propose_new_time(
                event_id=args.event_id,
                new_start=args.start,
                new_end=args.end,
                display_timezone=args.timezone,
                comment=args.comment
            )
            if args.json:
                print(json.dumps({"success": True, "message": f"New time proposed: {args.start} - {args.end}"}))
            else:
                print(f"✓ New time proposed: {args.start} - {args.end}")
        
    except Exception as e:
        if args.json:
            print(json.dumps({"success": False, "error": str(e)}))
        else:
            print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
