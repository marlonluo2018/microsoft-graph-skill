---
name: microsoft-graph-skill
description: |
  Microsoft Graph API for email, calendar, and user operations.  
  **Triggers:** "read my emails", "send an email", "check calendar", "create meeting".
version: 1.0.0
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
| List | `list --timezone TZ [--from EMAIL] [--subject TXT] [--start TS] [--end TS]` |
| Search | `search --timezone TZ [--from EMAIL] [--subject TXT]` |
| Find | `find --timezone TZ [--from EMAIL] [--subject TXT]` |
| Get Full | `get <message_id>` |
| Send | `send --to EMAILS --subject TXT --body TXT [--cc EMAILS] [--bcc EMAILS] [--csv PATH] [--attachments PATH]` |
| Reply | `reply <id> --body TXT [--bcc EMAILS] [--csv PATH] [--attachments PATH]` |
| Forward | `forward <id> --to EMAILS [--bcc EMAILS] [--csv PATH]` |
| Folders | `folders` |
| Attachments | `attachments <message_id> [--download]` |
| Accept Invite | `accept-invite <message_id>` |
| Decline Invite | `decline-invite <message_id>` |

**Required:** `--timezone` ⚠️ (e.g., "Asia/Shanghai", "UTC")

**Options:** `--folder`, `--top N`, `--detail`, `--emails-only`, `--events-only`

**BCC CSV File Support:**
- `--csv PATH` - Load BCC recipients from CSV file (auto-detects email column)
- `--email-column NAME` - Specify CSV column name (optional)
- Auto-batching when total recipients exceed 500
- Example CSV format: `email` column with one email per row

## Calendar Operations (`scripts/calendar_operations.py`)

| Action | Command |
|--------|---------|
| List | `list --timezone TZ [--start TS] [--end TS]` |
| Get | `get <event_id>` |
| Create | `create --subject TXT --start TS --end TS --timezone TZ [--required EMAILS] [--optional EMAILS] [--no-teams]` *(Teams enabled by default)* |
| Update | `update <event_id> --timezone TZ [--subject TXT] [--required EMAILS] [--optional EMAILS]` |
| Delete | `delete <event_id> [--permanent]` |
| Cancel | `cancel <event_id> [--comment TXT]` *(organizer only, notifies attendees)* |
| Forward | `forward <event_id> --to EMAILS [--comment TXT]` |
| Availability | `availability --emails EMAILS --start TS --end TS --timezone TZ` *(includes smart meeting suggestions)* |
| Propose | `propose <event_id> --start TS --end TS --timezone TZ` |
| Accept/Decline/Tentative | `accept|decline|tentative <event_id>` |

**Required:** `--timezone` ⚠️ (e.g., "Asia/Shanghai", "UTC")

**Attendee Types:**
- `--required EMAILS` - Required attendees (must attend)
- `--optional EMAILS` - Optional attendees (nice to have)
- Can use both parameters together to specify different types

**Other Options:** `--emails` (comma-separated for availability checks)

## Time Format

`--start` and `--end` must be **plain datetime** (no timezone) + `--timezone`:

| Format | Example |
|--------|---------|
| Plain datetime | `--start "2026-03-26T12:00:00" --timezone "Asia/Shanghai"` |
| Special | `--start "now" --timezone "Asia/Shanghai"` |

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
