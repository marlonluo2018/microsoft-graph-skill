#!/usr/bin/env python3
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from scripts.email_operations import get_event_from_message
from datetime import datetime
from zoneinfo import ZoneInfo

message_id = sys.argv[1] if len(sys.argv) > 1 else None
if not message_id:
    print("Usage: python get_event_details.py <message_id>")
    sys.exit(1)

try:
    event = get_event_from_message(message_id)
    
    print(f"Subject: {event['subject']}")
    print(f"Start: {event['start']['dateTime']} ({event['start']['timeZone']})")
    print(f"End: {event['end']['dateTime']} ({event['end']['timeZone']})")
    print(f"Location: {event.get('location', {}).get('displayName', 'Microsoft Teams')}")
    
    # Convert to user's timezone (UTC+8)
    # Handle Microsoft's datetime format with extra precision
    start_str = event['start']['dateTime'].split('.')[0]  # Remove microseconds
    end_str = event['end']['dateTime'].split('.')[0]
    
    start_dt = datetime.fromisoformat(start_str).replace(tzinfo=ZoneInfo('UTC'))
    end_dt = datetime.fromisoformat(end_str).replace(tzinfo=ZoneInfo('UTC'))
    
    user_tz = ZoneInfo('Asia/Shanghai')
    start_local = start_dt.astimezone(user_tz)
    end_local = end_dt.astimezone(user_tz)
    
    print(f"\nLocal Time (UTC+8):")
    print(f"Start: {start_local.strftime('%Y-%m-%d %H:%M')}")
    print(f"End: {end_local.strftime('%Y-%m-%d %H:%M')}")
    print(f"Duration: {(end_dt - start_dt).total_seconds() / 60:.0f} minutes")
    
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)

# Made with Bob
