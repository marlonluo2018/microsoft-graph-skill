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
| Send | `send --to EMAILS --subject TXT --body TXT [--attachments PATH]` |
| Reply | `reply <id> --body TXT [--attachments PATH]` |
| Forward | `forward <id> --to EMAILS` |
| Folders | `folders` |
| Attachments | `attachments <message_id> [--download]` |
| Accept Invite | `accept-invite <message_id>` |
| Decline Invite | `decline-invite <message_id>` |

**Required:** `--timezone` ⚠️ (e.g., "Asia/Shanghai", "UTC")

**Options:** `--folder`, `--top N`, `--detail`, `--emails-only`, `--events-only`

## Calendar Operations (`scripts/calendar_operations.py`)

| Action | Command |
|--------|---------|
| List | `list --timezone TZ [--start TS] [--end TS]` |
| Get | `get <event_id>` |
| Create | `create --subject TXT --start TS --end TS --timezone TZ [--attendees EMAILS]` |
| Update | `update <event_id> --timezone TZ [--subject TXT]` |
| Delete | `delete <event_id> [--permanent]` |
| Availability | `availability --emails EMAILS --start TS --end TS --timezone TZ` |
| Suggest | `suggest --attendees EMAILS --timezone TZ [--duration MIN]` |
| Propose | `propose <event_id> --start TS --end TS --timezone TZ` |
| Accept/Decline/Tentative | `accept|decline|tentative <event_id>` |

**Required:** `--timezone` ⚠️ (e.g., "Asia/Shanghai", "UTC")

**Options:** `--attendees`, `--emails` (comma-separated)

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
