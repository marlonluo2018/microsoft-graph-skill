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


def parse_datetime(dt_str: str) -> str:
    """Parse datetime string to ISO 8601 format."""
    # Try various formats
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
            return dt.strftime("%Y-%m-%dT%H:%M:%S")
        except ValueError:
            continue
    
    # If already in ISO format, return as is
    return dt_str


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
        start: Start datetime (ISO 8601)
        end: End datetime (ISO 8601)
        limit: Maximum number of events to return
        filter_query: OData filter query
        token: Access token
    
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
        "$select": "id,subject,start,end,organizer,attendees,isAllDay,location"
    }
    
    # Set default time range if not provided
    if start or end:
        params["startDateTime"] = parse_datetime(start) if start else datetime.now().isoformat()
        params["endDateTime"] = parse_datetime(end) if end else (datetime.now() + timedelta(days=30)).isoformat()
    
    if filter_query:
        params["$filter"] = filter_query
    
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
    timezone: str = "UTC",
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
        start: Start datetime
        end: End datetime
        timezone: Timezone (default UTC)
        body: Event body/description
        body_type: "html" or "text"
        location: Location string
        attendees: List of attendee dicts with 'email' and optionally 'name', 'type'
        is_all_day: Whether this is an all-day event
        recurrence: Recurrence pattern dict
        is_online_meeting: Whether to create as Teams meeting
        token: Access token
    
    Returns:
        Created event object
    """
    if token is None:
        token = get_access_token()
    
    event = {
        "subject": subject,
        "start": {
            "dateTime": parse_datetime(start),
            "timeZone": timezone
        },
        "end": {
            "dateTime": parse_datetime(end),
            "timeZone": timezone
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
    subject: str = None,
    start: str = None,
    end: str = None,
    timezone: str = None,
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
        event["start"] = {
            "dateTime": parse_datetime(start),
            "timeZone": timezone or "UTC"
        }
    
    if end:
        event["end"] = {
            "dateTime": parse_datetime(end),
            "timeZone": timezone or "UTC"
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
    timezone: str = "UTC",
    comment: str = None,
    send_response: bool = True,
    token: str = None
) -> bool:
    """
    Propose a new meeting time for an event (declines current time and proposes new).
    
    Args:
        event_id: Event ID
        new_start: Proposed new start datetime
        new_end: Proposed new end datetime
        timezone: Timezone (default UTC)
        comment: Optional message to organizer
        send_response: Whether to send response to organizer (default True)
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/events/{event_id}/decline"
    
    payload = {
        "proposedNewTime": {
            "start": {
                "dateTime": parse_datetime(new_start),
                "timeZone": timezone
            },
            "end": {
                "dateTime": parse_datetime(new_end),
                "timeZone": timezone
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
    timezone: str = "UTC",
    interval: int = 30,
    token: str = None
) -> Dict[str, Any]:
    """
    Get free/busy availability for specified users.
    
    Args:
        emails: List of email addresses to check
        start: Start datetime (ISO 8601)
        end: End datetime (ISO 8601)
        timezone: Timezone
        interval: Meeting time slot interval in minutes
        token: Access token
    
    Returns:
        Availability information for each user
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/calendar/getSchedule"
    
    payload = {
        "schedules": emails,
        "startTime": {
            "dateTime": parse_datetime(start),
            "timeZone": timezone
        },
        "endTime": {
            "dateTime": parse_datetime(end),
            "timeZone": timezone
        },
        "availabilityViewInterval": interval
    }
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code != 200:
        raise Exception(f"Failed to get availability: {response.status_code} - {response.text}")
    
    return response.json()


def suggest_meeting_times(
    attendees: List[str],
    duration_minutes: int = 60,
    start: str = None,
    end: str = None,
    timezone: str = "UTC",
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
        start: Search start datetime (default: now)
        end: Search end datetime (default: 7 days from now)
        timezone: Timezone for results (default UTC)
        top_n: Number of top slots to return (default 5)
        interval: Time slot interval in minutes (default 30)
        token: Access token
    
    Returns:
        Dict with top_time_slots ranked by score, plus detailed availability info
    """
    if token is None:
        token = get_access_token()
    
    # Default to next 7 days if not specified
    if start is None:
        start = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    if end is None:
        end = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%S")
    
    # Parse start datetime to get reference point
    start_dt = datetime.fromisoformat(start.replace("Z", "+00:00").replace("+00:00", ""))
    
    url = f"{GRAPH_API_BASE}/me/calendar/getschedule"
    
    payload = {
        "schedules": attendees,
        "startTime": {
            "dateTime": parse_datetime(start),
            "timeZone": timezone
        },
        "endTime": {
            "dateTime": parse_datetime(end),
            "timeZone": timezone
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
            "timezone": timezone,
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
        "timezone": timezone,
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

def display_event_list(events: List[Dict]):
    """Display a list of events in a readable format."""
    print(f"\n{'='*80}")
    print(f"{'Start':<25} {'Subject':<40} {'Location':<20}")
    print(f"{'='*80}")
    
    for event in events:
        start_dt = event.get('start', {}).get('dateTime', '')
        if start_dt:
            # Handle Microsoft Graph date format with 7 decimal places
            # Truncate to 6 decimal places for Python compatibility
            if '.' in start_dt:
                parts = start_dt.split('.')
                fractional = parts[1].rstrip('Z')[:6]
                start_dt = f"{parts[0]}.{fractional}"
            try:
                dt = datetime.fromisoformat(start_dt.replace('Z', '+00:00'))
                start_str = dt.strftime('%Y-%m-%d %H:%M')
            except ValueError:
                start_str = start_dt[:16]
        else:
            start_str = ''
        
        subject = event.get('subject', '(No Subject)')[:40]
        location = event.get('location', {}).get('displayName', '')[:20]
        
        print(f"{start_str:<25} {subject:<40} {location:<20}")
    
    print(f"{'='*80}")
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


def display_availability(data: Dict):
    """Display availability information."""
    print(f"\n{'='*80}")
    print("Availability Information")
    print(f"{'='*80}")
    
    for schedule in data.get("value", []):
        email = schedule.get("scheduleId", "Unknown")
        availability = schedule.get("availabilityView", "No data")
        print(f"\n{email}:")
        print(f"  Availability View: {availability[:50]}{'...' if len(availability) > 50 else ''}")
        
        for item in schedule.get("scheduleItems", []):
            print(f"  - {item.get('subject', 'Busy')}: {item.get('start')} - {item.get('end')}")
    
    print(f"\n{'='*80}")


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
    list_parser.add_argument("--start", help="Start datetime")
    list_parser.add_argument("--end", help="End datetime")
    list_parser.add_argument("--limit", type=int, default=25, help="Max events")
    
    # Get command
    get_parser = subparsers.add_parser("get", help="Get an event")
    get_parser.add_argument("event_id", help="Event ID")
    
    # Create command
    create_parser = subparsers.add_parser("create", help="Create an event")
    create_parser.add_argument("--subject", required=True, help="Event subject")
    create_parser.add_argument("--start", required=True, help="Start datetime")
    create_parser.add_argument("--end", required=True, help="End datetime")
    create_parser.add_argument("--timezone", default="UTC", help="Timezone")
    create_parser.add_argument("--body", help="Event description")
    create_parser.add_argument("--location", help="Location")
    create_parser.add_argument("--attendees", help="Attendee emails (comma-separated)")
    create_parser.add_argument("--all-day", action="store_true", help="All day event")
    create_parser.add_argument("--teams", action="store_true", help="Create Teams meeting")
    
    # Update command
    update_parser = subparsers.add_parser("update", help="Update an event")
    update_parser.add_argument("event_id", help="Event ID")
    update_parser.add_argument("--subject", help="New subject")
    update_parser.add_argument("--start", help="New start datetime")
    update_parser.add_argument("--end", help="New end datetime")
    update_parser.add_argument("--body", help="New description")
    update_parser.add_argument("--location", help="New location")
    
    # Delete command
    delete_parser = subparsers.add_parser("delete", help="Soft delete an event (move to Deleted Items)")
    delete_parser.add_argument("event_id", help="Event ID")
    delete_parser.add_argument("--permanent", action="store_true", help="Permanently delete (cannot recover)")
    
    # Availability command
    avail_parser = subparsers.add_parser("availability", help="Get availability")
    avail_parser.add_argument("--emails", required=True, help="Email addresses (comma-separated)")
    avail_parser.add_argument("--start", required=True, help="Start datetime")
    avail_parser.add_argument("--end", required=True, help="End datetime")
    avail_parser.add_argument("--timezone", default="UTC", help="Timezone")
    
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
    
    # Suggest meeting times command
    suggest_parser = subparsers.add_parser("suggest", help="Suggest optimal meeting times based on attendee availability")
    suggest_parser.add_argument("--attendees", required=True, help="Attendee emails (comma-separated)")
    suggest_parser.add_argument("--duration", type=int, default=60, help="Meeting duration in minutes (default 60)")
    suggest_parser.add_argument("--start", help="Search start datetime (default: now)")
    suggest_parser.add_argument("--end", help="Search end datetime (default: 7 days)")
    suggest_parser.add_argument("--timezone", default="UTC", help="Timezone (default UTC)")
    suggest_parser.add_argument("--top", type=int, default=5, help="Number of top slots to show (default 5)")
    
    # Propose new time command
    propose_parser = subparsers.add_parser("propose", help="Propose a new meeting time")
    propose_parser.add_argument("event_id", help="Event ID")
    propose_parser.add_argument("--start", required=True, help="Proposed new start datetime")
    propose_parser.add_argument("--end", required=True, help="Proposed new end datetime")
    propose_parser.add_argument("--timezone", default="UTC", help="Timezone")
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
                limit=args.limit
            )
            if args.json:
                print(json.dumps({"success": True, "events": events, "total": len(events)}, indent=2, default=str))
            else:
                display_event_list(events)
        
        elif args.command == "get":
            event = get_event(args.event_id)
            if args.json:
                print(json.dumps({"success": True, "event": event}, indent=2, default=str))
            else:
                display_event(event)
        
        elif args.command == "create":
            attendees = None
            if args.attendees:
                attendees = [{"email": e} for e in parse_email_list(args.attendees)]
            
            event = create_event(
                subject=args.subject,
                start=args.start,
                end=args.end,
                timezone=args.timezone,
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
            event = update_event(
                event_id=args.event_id,
                subject=args.subject,
                start=args.start,
                end=args.end,
                body=args.body,
                location=args.location
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
            data = get_availability(
                emails=parse_email_list(args.emails),
                start=args.start,
                end=args.end,
                timezone=args.timezone
            )
            if args.json:
                print(json.dumps({"success": True, "availability": data}, indent=2, default=str))
            else:
                display_availability(data)
        
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
                timezone=args.timezone,
                comment=args.comment
            )
            if args.json:
                print(json.dumps({"success": True, "message": f"New time proposed: {args.start} - {args.end}"}))
            else:
                print(f"✓ New time proposed: {args.start} - {args.end}")
        
        elif args.command == "suggest":
            result = suggest_meeting_times(
                attendees=parse_email_list(args.attendees),
                duration_minutes=args.duration,
                start=args.start,
                end=args.end,
                timezone=args.timezone,
                top_n=args.top
            )
            if args.json:
                # Remove raw_availability for cleaner JSON output
                output = {k: v for k, v in result.items() if k != "raw_availability"}
                print(json.dumps(output, indent=2, default=str))
            else:
                print(f"\n📅 Suggested Meeting Times (duration: {args.duration}min)")
                print(f"   Attendees: {', '.join(parse_email_list(args.attendees))}")
                print(f"   Timezone: {args.timezone}")
                print()
                if result["top_time_slots"]:
                    for slot in result["top_time_slots"]:
                        all_free = "✓ All free" if slot["all_free"] else f"{slot['free_count']}/{slot['total_attendees']} free"
                        print(f"   #{slot['rank']} {slot['start']} - {slot['end']} (Score: {slot['score']}) - {all_free}")
                        if slot["unavailable_attendees"]:
                            for u in slot["unavailable_attendees"]:
                                print(f"       - {u['email']}: {u['status']}")
                else:
                    print("   No suitable time slots found")
    
    except Exception as e:
        if args.json:
            print(json.dumps({"success": False, "error": str(e)}))
        else:
            print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
