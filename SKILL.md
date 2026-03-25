---
name: microsoft-graph-skill
description: Microsoft Graph API for email, calendar, and user operations. **ON STARTUP: Immediately check auth status (`python scripts/auth.py --status`). If not logged in, prompt user to login.** Triggers: "read my emails", "send an email", "check calendar", "create meeting", "find user".
version: 1.0.0
metadata:
  openclaw:
    requires:
      bins:
        - python
    emoji: "\U0001F4E7"
    homepage: https://github.com/openclaw/skills/tree/main/skills/microsoft-graph-skill
---

# Microsoft Graph Skill

## Overview

Microsoft Graph API for email, calendar, and user operations with OAuth2 device code authentication.

## Core Capabilities

### 1. Authentication (REQUIRED first)

**Script:** `scripts/auth.py`

| Action | Command |
|--------|---------|
| Start Login | `python scripts/auth.py --start` |
| Complete Login | `python scripts/auth.py --complete` |
| Check Status | `python scripts/auth.py --status` |
| Logout | `python scripts/auth.py --logout` |

**Workflow:** Run `--status` first. If `authenticated: false`, run `--start` → user enters code → run `--complete`.

### 2. Email Operations

**Script:** `scripts/email_operations.py`

| Action | Command | Description |
|--------|---------|-------------|
| List | `list [options]` | List/search messages |
| Search | `search [options]` | Alias for `list` (identical) |
| Find | `find [options]` | Alias for `list` (identical) |
| Get | `get <message_id>` | View full email content |
| Send | `send --to "..." --subject "..." --body "..."` | Send new email |
| Reply | `reply <id> --body "..." [--all]` | Reply to email |
| Forward | `forward <id> --to "..."` | Forward email |
| Folders | `folders` | List all mail folders |
| Attachments | `attachments <message_id>` | List attachments |
| Download All | `attachments <message_id> --download` | Download all attachments |
| Download One | `attachments <message_id> --id <att_id> -d` | Download specific attachment |

**Folder Options:**
- `--folder <name>` - Search in specific folder (default: inbox)
- `--folder all` - Search across ALL folders
- Available folder names: `inbox`, `sent`, `drafts`, `deleted`, `junk`, `outbox`
- Or use folder ID directly (get IDs with `folders` command)

**Key Options (common to list/search/find):**
- `--from <email_or_name>` - Search by sender
- `--to <email_or_name>` - Search by recipient
- `--subject <text>` - Search in subject
- `--body <text>` - Search in body
- `--top N` - Max results (default 25)
- `--folder <name>` - Specific folder (inbox/sent/drafts/deleted/all)
- `--filter <odata>` - Advanced OData filter for date queries

**Important - Search by Sender:**
- ✅ `list --from "email@domain.com"`
- ❌ `search --query "from:email"` (no --query parameter)

**Attachment Options:**
- `--download`, `-d` - Download attachments
- `--save-dir <path>` - Save directory (default: Desktop)
- `--id <att_id>` - Download specific attachment

**Auto-Features:** Outlook syntax detection, rate limit retry, CSV batching.

### 3. Calendar Operations

**Script:** `scripts/calendar_operations.py`

| Action | Command |
|--------|---------|
| List | `list --limit 10` |
| Get | `get <event_id>` |
| Create | `create --subject "..." --start "..." --end "..." [--attendees] [--teams]` |
| Update | `update <event_id> --subject "..."` |
| Delete | `delete <event_id> [--permanent]` |
| Availability | `availability --emails "..." --start "..." --end "..."` |
| Suggest | `suggest --attendees "..." --duration 60 [--top 5]` |

**As Organizer:** `create`, `update`, `cancel`, `forward`
**As Attendee:** `accept`, `decline`, `tentative`, `propose`

### 4. User Operations

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
