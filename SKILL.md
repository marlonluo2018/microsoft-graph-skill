---
name: microsoft-graph-skill
description: Microsoft Graph API for email, calendar, and user operations. **ON STARTUP: Immediately check auth status (`python scripts/auth.py --status`). If not logged in, prompt user to login.** Triggers: "read my emails", "send an email", "check calendar", "create meeting", "find user".
---

# Microsoft Graph Skill

## Overview

Microsoft Graph API for email, calendar, and user operations with OAuth2 device code authentication.

## Core Capabilities

### 1. Authentication (REQUIRED first)

**Script:** `scripts/auth.py`

|| Action | Command |
||--------|---------|
|| Check Status | `python scripts/auth.py --status` |
|| Start Login | `python scripts/auth.py --start` |
|| Complete Login | `python scripts/auth.py --complete` |
|| Logout | `python scripts/auth.py --logout` |

**Workflow:** Run `--status` first. If `authenticated: false`, run `--start` → user enters code → run `--complete`. Token auto-refresh is handled automatically.

### 2. Email Operations

**Script:** `scripts/email_operations.py`

| Action | Command |
|--------|---------|
| List | `python scripts/email_operations.py list --limit 10` |
| Search | `python scripts/email_operations.py search --from "John" --subject "report"` |
| Get | `python scripts/email_operations.py get <message_id>` |
| Thread | `python scripts/email_operations.py thread <message_id>` |
| Send | `python scripts/email_operations.py send --to "email" [--cc "..."] [--bcc "..."] --subject "..." --body "..."` |
| Reply | `python scripts/email_operations.py reply <id> --body "..." [--reply-all]` |
| Forward | `python scripts/email_operations.py forward <id> --to "email" --comment "..."` |

**Search options:** `--from`, `--to`, `--subject`, `--content`
**Send options:** `--to`, `--cc`, `--bcc`, `--subject`, `--body`

### 3. Calendar Operations

**Script:** `scripts/calendar_operations.py`

| Action | Command |
|--------|---------|
| List | `python scripts/calendar_operations.py list --limit 10` |
| Get | `python scripts/calendar_operations.py get <event_id>` |
| Create | `python scripts/calendar_operations.py create --subject "..." --start "..." --end "..." [--attendees "..."] [--location "..."] [--body "..."]` |
| Update | `python scripts/calendar_operations.py update <event_id> --subject "..."` |
| Delete | `python scripts/calendar_operations.py delete <event_id>` |
| Free/Busy | `python scripts/calendar_operations.py freebusy --emails "..." --start "..." --end "..."` |

**Create options:** `--subject`, `--start`, `--end`, `--attendees`, `--location`, `--body`

### 4. User Operations

**Script:** `scripts/user_operations.py`

| Action | Command |
|--------|---------|
| Get Me | `python scripts/user_operations.py get` |
| Search | `python scripts/user_operations.py search "john" [--limit 10] [--name-only] [--office "..."] [--detail]` |
| Manager | `python scripts/user_operations.py manager <user_id>` |
| Reports | `python scripts/user_operations.py reports <user_id>` |
| Find Times | `python scripts/user_operations.py find-times --attendees "..." --duration 60` |

**Search options:** `--limit`, `--name-only` (first name only), `--office`, `--detail`

## JSON Output

All commands support `--json` flag (place BEFORE subcommand):

```bash
python scripts/email_operations.py --json list --limit 10
python scripts/calendar_operations.py --json get <event_id>
```

## Constraints

- **Max 500 recipients** per email (To + CC + BCC combined)
- **Rate limits** apply - add delays for bulk operations
- Requires delegated permissions (see `references/permissions.md`)

## Resources

- `scripts/auth.py` - Authentication
- `scripts/email_operations.py` - Email CRUD
- `scripts/calendar_operations.py` - Calendar + Free/Busy
- `scripts/user_operations.py` - User search
- `references/api_endpoints.md` - API reference
- `references/permissions.md` - Required permissions
