---
name: microsoft-graph-skill
description: |
  Microsoft Graph API for email, calendar, and user operations.
  **ON USE:** Check auth status first. **CRITICAL:** `--since`/`--before` require timezone.
  **Triggers:** "read my emails", "send an email", "check calendar", "create meeting"
version: 1.0.0
---

# Microsoft Graph Skill

## Authentication (REQUIRED first)

**Script:** `scripts/auth.py`

| Action | Command |
|--------|---------|
| Check Status | `python scripts/auth.py --status` |
| Start Login | `python scripts/auth.py --start` |
| Complete Login | `python scripts/auth.py --complete` |
| Logout | `python scripts/auth.py --logout` |

**Workflow:** Run `--status` first. If `authenticated: false`, run `--start` → user enters code → run `--complete`.

## Email Operations

**Script:** `scripts/email_operations.py`

| Action | Command |
|--------|---------|
| List/Search | `list --from "..." --subject "..." --top 25` |
| Get Full | `get <message_id>` |
| Send | `send --to "..." --subject "..." --body "..."` |
| Reply | `reply <id> --body "..." [--sender-only]` |
| Forward | `forward <id> --to "..."` |
| Folders | `folders` |
| Attachments | `attachments <message_id> [--download] [--id <att_id>]` |

**Key Options:**
- `--from` / `--to` / `--subject` / `--body` - Search fields
- `--folder <name>` - inbox/sent/drafts/deleted/all (default: inbox)
- `--top N` - Max results (default 25)
- `--detail` - Show full body instead of preview
- `--since <timestamp>` / `--before <timestamp>` - Time filter (**MUST include timezone**)

**⚠️ Time Format (CRITICAL):**
- ✅ `"2026-03-26T12:00:00+08:00"` or `"2026-03-26T04:00:00Z"`
- ❌ `"2026-03-26"` or `"2026-03-26T12:00:00"` (no timezone)

## Calendar Operations

**Script:** `scripts/calendar_operations.py`

| Action | Command |
|--------|---------|
| List | `list --limit 10` |
| Get | `get <event_id>` |
| Create | `create --subject "..." --start "..." --end "..." [--attendees] [--teams]` |
| Update | `update <event_id> --subject "..."` |
| Delete | `delete <event_id> [--permanent]` |
| Availability | `availability --emails "..." --start "..." --end "..."` |
| Suggest | `suggest --attendees "..." --duration 60` |

**Attendee Actions:** `accept`, `decline`, `tentative`, `propose`

## User Operations

**Script:** `scripts/user_operations.py`

| Action | Command |
|--------|---------|
| Get Me | `get` |
| Search | `search "john" [--detail]` |
| Manager | `manager <user_id>` |
| Reports | `reports <user_id>` |

## Notes
- All commands support `--json` flag
- Max 500 recipients per email
- Rate limits handled automatically
