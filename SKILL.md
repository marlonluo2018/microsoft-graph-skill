---
name: microsoft-graph-skill
description: This skill should be used when the user needs to interact with Microsoft Graph API for email operations (read, send, reply, forward), calendar management (create events, query availability), or user/contact queries. Use this skill for Outlook email automation, Teams calendar operations, and Microsoft 365 user management. Triggers include requests like "read my emails", "send an email to...", "check calendar availability", "create a meeting", or "find user information".
---

# Microsoft Graph Skill

## Script Execution

**Always use absolute paths to execute scripts.**

## Overview

This skill enables interaction with Microsoft Graph API to manage emails, calendars, and user contacts. It provides authentication via device code flow and supports comprehensive email and calendar operations with proper handling of company restrictions (e.g., 500 recipient limit per email).

## Core Capabilities

### 1. Authentication (Device Code Flow)

Use the device code flow for initial login to Microsoft Graph. This is required before any API operations.

**Commands:**
```bash
# Start login flow - displays URL and device code
python scripts/auth.py --start

# Complete authentication after user enters code in browser
python scripts/auth.py --complete

# Check authentication status
python scripts/auth.py --status

# Refresh expired token
python scripts/auth.py --refresh

# Logout and clear cached tokens
python scripts/auth.py --logout

# Extend current session (refresh token proactively)
python scripts/auth.py --extend
```

**Script:** `scripts/auth.py`

### 2. Email Operations

Comprehensive email management including read, send, reply, and forward operations.

**All commands support `--json` flag for machine-readable output (place before subcommand):**
```bash
python scripts/email_operations.py --json list --limit 10
python scripts/email_operations.py --json search --from "John"
```

**Commands:**

#### List Emails
```bash
# List recent emails (default 25)
python scripts/email_operations.py list

# Limit number of results
python scripts/email_operations.py list --limit 10

# JSON output for AI agent
python scripts/email_operations.py --json list --limit 10
```

#### Search Emails
```bash
# Search by sender name/email
python scripts/email_operations.py search --from "John"

# Search by recipient
python scripts/email_operations.py search --to "recipient@example.com"

# Search by subject
python scripts/email_operations.py search --subject "meeting"

# Search in email body/content
python scripts/email_operations.py search --content "important"

# Combine search criteria
python scripts/email_operations.py search --from "John" --subject "report"

# JSON output
python scripts/email_operations.py --json search --from "John"
```

#### Get Email Details
```bash
# Get full email content by message ID
python scripts/email_operations.py get <message_id>

# JSON output
python scripts/email_operations.py --json get <message_id>
```

#### View Email Thread
```bash
# View entire conversation thread
python scripts/email_operations.py thread <message_id>

# JSON output
python scripts/email_operations.py --json thread <message_id>
```

#### Send Email
```bash
# Send new email
python scripts/email_operations.py send \
  --to "recipient@example.com" \
  --subject "Test Subject" \
  --body "Email body content"

# With CC and BCC
python scripts/email_operations.py send \
  --to "recipient@example.com" \
  --cc "cc1@example.com,cc2@example.com" \
  --bcc "bcc@example.com" \
  --subject "Subject" \
  --body "Body content"

# JSON output (returns message ID on success)
python scripts/email_operations.py --json send \
  --to "recipient@example.com" \
  --subject "Subject" \
  --body "Body"
```

#### Reply to Email
```bash
# Reply to sender only
python scripts/email_operations.py reply <message_id> --body "Reply content"

# Reply to all recipients
python scripts/email_operations.py reply <message_id> --body "Reply content" --reply-all

# JSON output
python scripts/email_operations.py --json reply <message_id> --body "Reply content"
```

#### Forward Email
```bash
# Forward email to new recipient
python scripts/email_operations.py forward <message_id> \
  --to "recipient@example.com" \
  --comment "FYI"

# JSON output
python scripts/email_operations.py --json forward <message_id> \
  --to "recipient@example.com"
```

**Important Constraints:**
- Maximum 500 recipients per email (To + CC + BCC combined)
- Company policy restriction

**Script:** `scripts/email_operations.py`

### 3. Calendar Management

Create and manage calendar events with attendee management.

**All commands support `--json` flag for machine-readable output (place before subcommand):**
```bash
python scripts/calendar_operations.py --json list --limit 10
python scripts/calendar_operations.py --json get <event_id>
```

**Commands:**

#### List Events
```bash
# List upcoming events (default 25)
python scripts/calendar_operations.py list

# Limit number of results
python scripts/calendar_operations.py list --limit 10

# JSON output
python scripts/calendar_operations.py --json list --limit 10
```

#### Get Event Details
```bash
# Get full event details by event ID
python scripts/calendar_operations.py get <event_id>

# JSON output
python scripts/calendar_operations.py --json get <event_id>
```

#### Create Event
```bash
# Create a meeting
python scripts/calendar_operations.py create \
  --subject "Meeting Subject" \
  --start "2026-03-22T10:00:00" \
  --end "2026-03-22T11:00:00"

# With attendees
python scripts/calendar_operations.py create \
  --subject "Team Meeting" \
  --start "2026-03-22T10:00:00" \
  --end "2026-03-22T11:00:00" \
  --attendees "user1@example.com,user2@example.com"

# With location and body
python scripts/calendar_operations.py create \
  --subject "Team Meeting" \
  --start "2026-03-22T10:00:00" \
  --end "2026-03-22T11:00:00" \
  --location "Conference Room A" \
  --body "Meeting agenda:\n1. Review\n2. Discussion"

# JSON output (returns event ID on success)
python scripts/calendar_operations.py --json create \
  --subject "Meeting" \
  --start "2026-03-22T10:00:00" \
  --end "2026-03-22T11:00:00"
```

#### Update Event
```bash
# Update event subject
python scripts/calendar_operations.py update <event_id> --subject "New Subject"

# Update event time
python scripts/calendar_operations.py update <event_id> \
  --start "2026-03-22T14:00:00" \
  --end "2026-03-22T15:00:00"

# JSON output
python scripts/calendar_operations.py --json update <event_id> --subject "New Subject"
```

#### Delete Event
```bash
# Delete an event
python scripts/calendar_operations.py delete <event_id>

# JSON output
python scripts/calendar_operations.py --json delete <event_id>
```

**Script:** `scripts/calendar_operations.py`

### 4. Availability Queries (Free/Busy)

Query attendees' availability to find suitable meeting times.

**Commands:**

#### Query Free/Busy Status
```bash
# Check availability for one or more users
python scripts/calendar_operations.py freebusy \
  --emails "user1@example.com,user2@example.com" \
  --start "2026-03-22T09:00:00" \
  --end "2026-03-22T18:00:00"

# JSON output
python scripts/calendar_operations.py --json freebusy \
  --emails "user1@example.com,user2@example.com" \
  --start "2026-03-22T09:00:00" \
  --end "2026-03-22T18:00:00"
```

**Returns:**
- Free/busy status for each time slot
- Detailed availability information (free, busy, tentative, out of office)

**Script:** `scripts/calendar_operations.py`

### 5. User/Contact Queries

Search and retrieve user and contact information.

**All commands support `--json` flag for machine-readable output (place before subcommand):**
```bash
python scripts/user_operations.py --json search "john"
python scripts/user_operations.py --json get
```

**Commands:**

#### Get Current User Info
```bash
# Get current authenticated user's profile
python scripts/user_operations.py get

# JSON output
python scripts/user_operations.py --json get
```

#### Search Users
```bash
# Search users by name or email
python scripts/user_operations.py search "john"

# Limit number of results
python scripts/user_operations.py search "john" --limit 10

# Search by given name only (first name)
python scripts/user_operations.py search "john" --name-only

# Filter by office location or email domain (client-side filter)
python scripts/user_operations.py search "john" --office "philippines"
python scripts/user_operations.py search "john" --office "ph.ibm.com"

# Show detailed user info
python scripts/user_operations.py search "john" --detail

# JSON output
python scripts/user_operations.py --json search "john"
python scripts/user_operations.py --json search "john" --name-only --office "philippines"
```

**Search Options:**
- `--name-only`: Search only by given name (first name)
- `--office <text>`: Filter by office location or email domain (client-side filtering)
- `--detail`: Show detailed user information (job title, department, phone, office)
- `--limit <n>`: Limit number of results

#### Get User's Manager
```bash
# Get manager of a user
python scripts/user_operations.py manager <user_id_or_email>

# JSON output
python scripts/user_operations.py --json manager <user_id>
```

#### Get Direct Reports
```bash
# Get direct reports of a user
python scripts/user_operations.py reports <user_id_or_email>

# JSON output
python scripts/user_operations.py --json reports <user_id>
```

#### Find Meeting Times
```bash
# Find available meeting times for attendees
python scripts/user_operations.py find-times \
  --attendees "user1@example.com,user2@example.com" \
  --duration 60

# JSON output
python scripts/user_operations.py --json find-times \
  --attendees "user1@example.com,user2@example.com" \
  --duration 60
```

**Script:** `scripts/user_operations.py`

## Workflow Decision Tree

```
User Request
    │
    ├─► Authentication needed?
    │       ├─► Start login → python scripts/auth.py --start
    │       ├─► Complete → python scripts/auth.py --complete
    │       ├─► Check status → python scripts/auth.py --status
    │       └─► Refresh → python scripts/auth.py --refresh
    │
    ├─► Email operations?
    │       ├─► List emails → python scripts/email_operations.py list
    │       ├─► Search → python scripts/email_operations.py search --from "..."
    │       ├─► Get details → python scripts/email_operations.py get <id>
    │       ├─► View thread → python scripts/email_operations.py thread <id>
    │       ├─► Send → python scripts/email_operations.py send --to "..." --subject "..." --body "..."
    │       ├─► Reply → python scripts/email_operations.py reply <id> --body "..."
    │       └─► Forward → python scripts/email_operations.py forward <id> --to "..."
    │
    ├─► Calendar operations?
    │       ├─► List events → python scripts/calendar_operations.py list
    │       ├─► Get event → python scripts/calendar_operations.py get <id>
    │       ├─► Create → python scripts/calendar_operations.py create --subject "..." --start "..." --end "..."
    │       ├─► Update → python scripts/calendar_operations.py update <id> --subject "..."
    │       ├─► Delete → python scripts/calendar_operations.py delete <id>
    │       └─► Availability → python scripts/calendar_operations.py freebusy --emails "..."
    │
    └─► User operations?
            ├─► Get current user → python scripts/user_operations.py get
            ├─► Search users → python scripts/user_operations.py search "..."
            ├─► Get manager → python scripts/user_operations.py manager <id>
            ├─► Get reports → python scripts/user_operations.py reports <id>
            └─► Find times → python scripts/user_operations.py find-times --attendees "..."
```

## JSON Output for AI Agents

All scripts support `--json` flag for machine-readable output. Place `--json` BEFORE the subcommand:

```bash
# Correct placement
python scripts/email_operations.py --json list
python scripts/calendar_operations.py --json get <id>
python scripts/user_operations.py --json search "john"

# Incorrect (won't work)
python scripts/email_operations.py list --json
```

JSON output returns structured data without formatting, colors, or progress messages - ideal for AI agent consumption.

## Important Constraints

1. **Recipient Limit**: Maximum 500 recipients per email (To + CC + BCC combined)
2. **Authentication**: Device code flow requires user interaction
3. **Rate Limits**: Microsoft Graph has rate limits; implement appropriate delays for bulk operations
4. **Permissions**: Requires appropriate delegated permissions (see `references/permissions.md`)

## Resources

### scripts/
- `auth.py` - OAuth2 device code authentication
- `email_operations.py` - Email read, send, reply, forward operations
- `calendar_operations.py` - Calendar event management and availability queries
- `user_operations.py` - User and contact search operations

### references/
- `api_endpoints.md` - Microsoft Graph API endpoint reference
- `permissions.md` - Required permissions for different operations
