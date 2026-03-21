# Microsoft Graph API Endpoints Reference

This document provides a quick reference for commonly used Microsoft Graph API endpoints.

## Base URL

```
https://graph.microsoft.com/v1.0
```

For beta features:
```
https://graph.microsoft.com/beta
```

---

## Authentication Endpoints

### OAuth2 Device Code Flow

**Get Device Code:**
```
POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode
Content-Type: application/x-www-form-urlencoded

client_id={client_id}&scope={scopes}
```

**Poll for Token:**
```
POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

grant_type=urn:ietf:params:oauth:grant-type:device_code
&client_id={client_id}
&device_code={device_code}
```

---

## Mail Endpoints

### List Messages
```
GET /me/mailFolders/{folder}/messages
GET /me/mailFolders/{folder}/messages?$top=50&$orderby=receivedDateTime desc
GET /me/mailFolders/{folder}/messages?$filter=isRead eq false
```

**Common Folders:**
- `inbox` - Inbox
- `sentitems` - Sent Items
- `drafts` - Drafts
- `deleteditems` - Deleted Items
- `junkemail` - Junk Email
- `outbox` - Outbox

### Get Message
```
GET /me/messages/{message-id}
GET /me/messages/{message-id}?$select=subject,body,from,toRecipients
```

### Send Email
```
POST /me/sendMail
Content-Type: application/json

{
  "message": {
    "subject": "Subject",
    "body": {
      "contentType": "HTML",
      "content": "Body content"
    },
    "toRecipients": [
      {
        "emailAddress": {
          "address": "user@example.com"
        }
      }
    ],
    "ccRecipients": [],
    "bccRecipients": []
  },
  "saveToSentItems": true
}
```

### Reply to Message
```
POST /me/messages/{message-id}/reply
POST /me/messages/{message-id}/replyAll

{
  "message": {
    "body": {
      "contentType": "HTML",
      "content": "Reply content"
    }
  }
}
```

### Forward Message
```
POST /me/messages/{message-id}/forward

{
  "toRecipients": [
    {
      "emailAddress": {
        "address": "forward@example.com"
      }
    }
  ],
  "comment": "Forward note"
}
```

### Mark as Read/Unread
```
PATCH /me/messages/{message-id}

{
  "isRead": true
}
```

### Delete Message
```
DELETE /me/messages/{message-id}
```

---

## Calendar Endpoints

### List Events
```
GET /me/calendar/events
GET /me/calendars/{calendar-id}/events
GET /me/calendar/events?startDateTime={start}&endDateTime={end}
```

### Get Event
```
GET /me/events/{event-id}
```

### Create Event
```
POST /me/events
Content-Type: application/json

{
  "subject": "Meeting Subject",
  "body": {
    "contentType": "HTML",
    "content": "Meeting description"
  },
  "start": {
    "dateTime": "2024-01-15T10:00:00",
    "timeZone": "Pacific Standard Time"
  },
  "end": {
    "dateTime": "2024-01-15T11:00:00",
    "timeZone": "Pacific Standard Time"
  },
  "location": {
    "displayName": "Conference Room"
  },
  "attendees": [
    {
      "emailAddress": {
        "address": "attendee@example.com",
        "name": "Attendee Name"
      },
      "type": "required"
    }
  ],
  "isOnlineMeeting": true,
  "onlineMeetingProvider": "teamsForBusiness"
}
```

### Update Event
```
PATCH /me/events/{event-id}
```

### Delete Event
```
DELETE /me/events/{event-id}
```

### Get Availability (Free/Busy)
```
POST /me/calendar/getSchedule
Content-Type: application/json

{
  "schedules": ["user1@example.com", "user2@example.com"],
  "startTime": {
    "dateTime": "2024-01-15T09:00:00",
    "timeZone": "UTC"
  },
  "endTime": {
    "dateTime": "2024-01-15T17:00:00",
    "timeZone": "UTC"
  },
  "availabilityViewInterval": 30
}
```

### Find Meeting Times
```
POST /me/findMeetingTimes
Content-Type: application/json

{
  "attendees": [
    {
      "emailAddress": {
        "address": "attendee@example.com"
      },
      "type": "required"
    }
  ],
  "timeConstraint": {
    "timeslots": [
      {
        "start": {
          "dateTime": "2024-01-15T09:00:00",
          "timeZone": "UTC"
        },
        "end": {
          "dateTime": "2024-01-15T17:00:00",
          "timeZone": "UTC"
        }
      }
    ]
  },
  "meetingDuration": "PT1H"
}
```

### List Calendars
```
GET /me/calendars
```

---

## User Endpoints

### Get Current User
```
GET /me
GET /me?$select=displayName,mail,jobTitle,department
```

### Get User by ID/UPN
```
GET /users/{user-id}
GET /users/{user-principal-name}
```

### List/Search Users
```
GET /users
GET /users?$filter=startsWith(displayName,'John')
GET /users?$search="displayName:John"
```

### Get Manager
```
GET /me/manager
GET /users/{user-id}/manager
```

### Get Direct Reports
```
GET /me/directReports
GET /users/{user-id}/directReports
```

---

## Contact Endpoints

### List Contacts
```
GET /me/contacts
GET /me/contactFolders/{folder-id}/contacts
```

### Get Contact
```
GET /me/contacts/{contact-id}
```

### Search Contacts
```
GET /me/contacts?$filter=startsWith(displayName,'John')
```

### List Contact Folders
```
GET /me/contactFolders
```

---

## People Endpoints (Suggested)

### Get People
```
GET /me/people
GET /me/people?$search="John"
```

---

## Common Query Parameters

### $select - Select specific properties
```
GET /me/messages?$select=subject,from,receivedDateTime
```

### $filter - Filter results
```
GET /me/messages?$filter=isRead eq false
GET /me/messages?$filter=receivedDateTime ge 2024-01-01T00:00:00Z
```

### $orderby - Sort results
```
GET /me/messages?$orderby=receivedDateTime desc
```

### $top - Limit results
```
GET /me/messages?$top=50
```

### $skip - Pagination
```
GET /me/messages?$skip=50&$top=50
```

### $expand - Expand related entities
```
GET /me/messages/{id}?$expand=attachments
```

---

## Common Time Zones

| Windows Time Zone | IANA Time Zone |
|-------------------|----------------|
| UTC | UTC |
| Pacific Standard Time | America/Los_Angeles |
| Mountain Standard Time | America/Denver |
| Central Standard Time | America/Chicago |
| Eastern Standard Time | America/New_York |
| China Standard Time | Asia/Shanghai |
| Singapore Standard Time | Asia/Singapore |
| Tokyo Standard Time | Asia/Tokyo |
| GMT Standard Time | Europe/London |
| Central European Standard Time | Europe/Paris |

---

## Error Responses

### Common HTTP Status Codes

| Code | Description |
|------|-------------|
| 200 | Success |
| 201 | Created |
| 202 | Accepted (async operation) |
| 204 | No Content (success, no body) |
| 400 | Bad Request |
| 401 | Unauthorized |
| 403 | Forbidden |
| 404 | Not Found |
| 409 | Conflict |
| 429 | Too Many Requests (rate limited) |
| 500 | Internal Server Error |

### Error Response Format
```json
{
  "error": {
    "code": "Error code",
    "message": "Error description",
    "innerError": {
      "request-id": "UUID",
      "date": "Timestamp"
    }
  }
}
```

---

## Rate Limits

Microsoft Graph API has service-specific limits. Key considerations:

- **Mail**: 10,000 requests per 10 minutes per mailbox
- **Calendar**: Similar limits apply
- **Users**: Standard throttling applies

When rate limited (429), check the `Retry-After` header for wait time.
