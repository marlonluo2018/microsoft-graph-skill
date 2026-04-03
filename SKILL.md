---
name: microsoft-graph-skill
description: |
  Microsoft Graph API for email, calendar, and user operations.
  **Triggers:** "read my emails", "send an email", "check calendar", "create meeting".
  **Time parameters:** `--end` supports "now" keyword (e.g., `--end "now"`).
---

# Microsoft Graph Skill
**All commands require email addresses only** (e.g., `user@example.com`). Names are NOT supported.

## Authentication (`scripts/auth.py`)

| Action | Command |
|--------|---------|
| Check Status | `--status` |
| Start Login | `--start` |
| Complete Login | `--complete` |
| Logout | `--logout` |

**Note:** Commands auto-refresh expired tokens. Manual auth only needed for initial setup or after logout.

## Email Operations (`scripts/email_operations.py`)

| Action | Command |
|--------|---------|
| List | `list --timezone TZ` |
| Search | `search --timezone TZ` |
| Find | `find --timezone TZ` |
| Get Full | `get <message_id>` |
| Send | `send --to EMAILS --subject TXT --body TXT` |
| Reply | `reply <id> --body TXT` |
| Forward | `forward <id> --to EMAILS` |
| Folders | `folders` |
| Attachments | `attachments <message_id>` |
| Accept Invite | `accept-invite <message_id>` |
| Decline Invite | `decline-invite <message_id>` |

**Required for list/search/find:** `--timezone` ⚠️ (e.g., "Asia/Shanghai", "UTC")

**Filtering Options (for list/search/find):**
- `--from EMAIL` - Filter by sender email address
- `--to EMAIL` - Filter by recipient email address (useful for searching sent items)
- `--subject TXT` - Filter by subject text
- `--folder FOLDER` - Specify mail folder (e.g., "inbox", "sentitems", "deleteditems", "drafts")
- `--start TS` - Start date/time for filtering (use with --timezone)
- `--end TS` - End date/time for filtering (use with --timezone, supports "now")
- `--top N` - Limit number of results (default: 10)
- `--detail` - Show detailed email content
- `--emails-only` - Show only emails (exclude calendar events)
- `--events-only` - Show only calendar events (exclude emails)

**Send/Reply/Forward Options:**
- `--cc EMAILS` - CC recipients (comma-separated)
- `--bcc EMAILS` - BCC recipients (comma-separated)
- `--attachments PATH` - Attach files (comma-separated paths)
- `--csv PATH` - Load BCC recipients from CSV file (for send/reply)
- `--download` - Download attachments (for attachments command)

**BCC CSV File Support:**
- `--csv PATH` - Load BCC recipients from CSV file (auto-detects email column)
- `--email-column NAME` - Specify CSV column name (optional)
- Auto-batching when total recipients exceed 500
- Example CSV format: `email` column with one email per row

## Calendar Operations (`scripts/calendar_operations.py`)

| Action | Command |
|--------|---------|
| List | `list --timezone TZ` |
| Get | `get <event_id>` |
| Create | `create --subject TXT --start TS --end TS --timezone TZ` |
| Update | `update <event_id> --timezone TZ` |
| Delete | `delete <event_id>` |
| Cancel | `cancel <event_id>` |
| Forward | `forward <event_id> --to EMAILS` |
| Availability | `availability --emails EMAILS --start TS --end TS --timezone TZ` |
| Propose | `propose <event_id> --start TS --end TS --timezone TZ` |
| Accept/Decline/Tentative | `accept|decline|tentative <event_id>` |

**Required for most commands:** `--timezone` ⚠️ (e.g., "Asia/Shanghai", "UTC")

**Calendar Options:**

**Time & Date Filtering:**
- `--start TS` - Start date/time (use with --timezone)
- `--end TS` - End date/time (use with --timezone, supports "now")

**Attendee Management:**
- `--required EMAILS` - Required attendees (must attend) - comma-separated emails
- `--optional EMAILS` - Optional attendees (nice to have) - comma-separated emails
- `--emails EMAILS` - Email addresses for availability checks (comma-separated)
- Note: Can use both --required and --optional together to specify different attendee types

**Meeting Options:**
- `--subject TXT` - Meeting subject/title
- `--no-teams` - Disable Teams meeting (Teams enabled by default for create command)
- `--comment TXT` - Add comment when canceling or forwarding
- `--permanent` - Permanently delete event (for delete command)
- `--to EMAILS` - Forward recipients (comma-separated)

**Notes:**
- Cancel command is organizer only and notifies attendees
- Availability command includes smart meeting suggestions

## Time Format

`--start` and `--end` must be **plain datetime** (no timezone) + `--timezone`:

| Format | Example |
|--------|---------|
| Plain datetime | `--start "2026-03-26T12:00:00" --timezone "Asia/Shanghai"` |
| Special | `--end "now" --timezone "Asia/Shanghai"` |

**Note:** Embedded timezone (`Z` or `+08:00`) is NOT supported. Always use `--timezone`.

## User Operations (`scripts/user_operations.py`)

| Action | Command |
|--------|---------|
| Get Me | `get` |
| Search | `search QUERY [--detail]` |
| Manager | `manager <user_id>` |
| Reports | `reports <user_id>` |

## Notes
- All commands support `--json` flag
- Max 500 recipients per email
- Rate limits handled automatically
