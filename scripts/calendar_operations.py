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

def delete_event(event_id: str, token: str = None) -> bool:
    """
    Delete a calendar event.
    
    Args:
        event_id: Event ID to delete
        token: Access token
    
    Returns:
        bool: True if successful
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/events/{event_id}"
    
    response = requests.delete(url, headers=get_headers(token))
    
    if response.status_code != 204:
        raise Exception(f"Failed to delete event: {response.status_code} - {response.text}")
    
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
    token: str = None
) -> List[Dict[str, Any]]:
    """
    Suggest available meeting times based on attendee availability.
    
    Args:
        attendees: List of attendee email addresses
        duration_minutes: Required meeting duration
        start: Search start datetime
        end: Search end datetime
        timezone: Timezone
        token: Access token
    
    Returns:
        List of suggested meeting time slots
    """
    if token is None:
        token = get_access_token()
    
    # Default to next 7 days if not specified
    if start is None:
        start = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    if end is None:
        end = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%S")
    
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
        "availabilityViewInterval": 30
    }
    
    response = requests.post(url, headers=get_headers(token), json=payload)
    
    if response.status_code != 200:
        raise Exception(f"Failed to get suggestions: {response.status_code} - {response.text}")
    
    # Parse and analyze availability
    data = response.json()
    # TODO: Implement intelligent time slot suggestion based on availability data
    
    return data


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
    delete_parser = subparsers.add_parser("delete", help="Delete an event")
    delete_parser.add_argument("event_id", help="Event ID")
    
    # Availability command
    avail_parser = subparsers.add_parser("availability", help="Get availability")
    avail_parser.add_argument("--emails", required=True, help="Email addresses (comma-separated)")
    avail_parser.add_argument("--start", required=True, help="Start datetime")
    avail_parser.add_argument("--end", required=True, help="End datetime")
    avail_parser.add_argument("--timezone", default="UTC", help="Timezone")
    
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
            delete_event(args.event_id)
            if args.json:
                print(json.dumps({"success": True, "message": "Event deleted"}))
            else:
                print("✓ Event deleted")
        
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
    
    except Exception as e:
        if args.json:
            print(json.dumps({"success": False, "error": str(e)}))
        else:
            print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
