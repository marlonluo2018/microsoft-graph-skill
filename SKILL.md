---
name: microsoft-graph
description: This skill should be used when the user needs to interact with Microsoft Graph API for email operations (read, send, reply, forward), calendar management (create events, query availability), or user/contact queries. Use this skill for Outlook email automation, Teams calendar operations, and Microsoft 365 user management. Triggers include requests like "read my emails", "send an email to...", "check calendar availability", "create a meeting", or "find user information".
---

# Microsoft Graph

## Overview

This skill enables interaction with Microsoft Graph API to manage emails, calendars, and user contacts. It provides authentication via device code flow and supports comprehensive email and calendar operations with proper handling of company restrictions (e.g., 500 recipient limit per email).

## Core Capabilities

### 1. Authentication (Device Code Flow)

Use the device code flow for initial login to Microsoft Graph. This is required before any API operations.

**Workflow:**
1. Call `scripts/auth.py` to initiate authentication
2. Display the login URL and device code to the user
3. User visits the URL and enters the code
4. Wait for authentication completion
5. Tokens are cached for subsequent operations

**Script:** `scripts/auth.py`

### 2. Email Operations

Comprehensive email management including read, send, reply, and forward operations.

#### Read Emails
- List messages from inbox or specific folders
- Filter by date, sender, subject, or read status
- Retrieve message content and attachments

#### Send Emails
- Send new emails with body content
- Support for CC and BCC recipients
- **Important:** Company restricts each email to maximum 500 recipients (To + CC + BCC)
- Validate recipient count before sending

#### Reply to Emails
- Reply to sender only (Reply)
- Reply to all recipients (Reply All)
- Include original message in reply

#### Forward Emails
- Forward emails to new recipients
- Add optional comments
- Support CC/BCC for forwarded messages

**Script:** `scripts/email_operations.py`

### 3. Calendar Management

Create and manage calendar events with attendee management.

#### Create Events
- Create meetings with subject, body, and time
- Add required and optional attendees
- Set location (physical or online/Teams)
- Configure reminders and recurrence

#### Update/Delete Events
- Modify existing events
- Cancel meetings
- Send updates to attendees

#### Send Emails to Attendees
- Send meeting-related communications
- Notify attendees of changes

**Script:** `scripts/calendar_operations.py`

### 4. Availability Queries (Free/Busy)

Query attendees' availability to find suitable meeting times.

**Capabilities:**
- Check free/busy status for specified time periods
- Query multiple attendees simultaneously
- Get detailed availability information (free, busy, tentative, out of office)
- Suggest optimal meeting times based on availability

**Script:** `scripts/calendar_operations.py` (function: `get_availability`)

### 5. User/Contact Queries

Search and retrieve user and contact information.

#### User Queries
- Search users by name, email, or other attributes
- Get user profile information
- List direct reports and manager

#### Contact Queries
- Search personal contacts
- Get contact details
- List contact folders

**Script:** `scripts/user_operations.py`

## Workflow Decision Tree

```
User Request
    │
    ├─► Is this about email?
    │       ├─► Read → Use email_operations.py:list_messages()
    │       ├─► Send → Use email_operations.py:send_email()
    │       │           └─► Validate recipient count ≤ 500
    │       ├─► Reply → Use email_operations.py:reply_email()
    │       └─► Forward → Use email_operations.py:forward_email()
    │
    ├─► Is this about calendar?
    │       ├─► Create event → Use calendar_operations.py:create_event()
    │       ├─► Query availability → Use calendar_operations.py:get_availability()
    │       ├─► Update event → Use calendar_operations.py:update_event()
    │       └─► Delete event → Use calendar_operations.py:delete_event()
    │
    └─► Is this about users/contacts?
            ├─► Search users → Use user_operations.py:search_users()
            └─► Search contacts → Use user_operations.py:search_contacts()
```

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
