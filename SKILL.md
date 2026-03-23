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

| Action | Command |
|--------|---------|
| Start Login | `python scripts/auth.py --start` |
| Complete Login | `python scripts/auth.py --complete` |
| Check Status | `python scripts/auth.py --status` |
| Logout | `python scripts/auth.py --logout` |

**Workflow:** Run `--status` first. If `authenticated: false`, run `--start` → user enters code → run `--complete`. Token auto-refresh is handled automatically.

**Robustness Improvements (v2.0):**
- Automatic cleanup: Expired device flows are automatically cleaned on startup
- Comprehensive logging: All operations logged to `cache/auth.log`
- Retry mechanism: Network operations retry (3 attempts) with exponential backoff
- Thread-safe operations: Token operations protected with locks
- Configuration validation: Validates required config before operations
- Enhanced error messages: Detailed errors with suggested actions
- Graceful error handling: Catches and logs exceptions
- UTF-8 encoding support: Proper file encoding for international characters
- Verbose mode: Add `--verbose` flag for detailed debug logs

### 2. Email Operations

**Script:** `scripts/email_operations.py`

**CRITICAL: Command Format Rules (READ FIRST)**
1. Subcommand FIRST, then options: `py -3 email_operations.py <subcommand> [options]`
2. NO --json after subcommand: Use `py -3 email_operations.py --json <subcommand>` if needed
3. Quote multi-word arguments: `--subject "Annual Enrollment"` NOT `--subject Annual Enrollment`

| Action | Command |
|--------|---------|
| List/Search | `list [options]` or `search [options]` (identical) |
| Find | `find --from/--to/--subject [--folder]` |
| Get | `get <message_id>` |
| Send | `send --to "..." --subject "..." --body "..." [--csv]` |
| Reply | `reply <id> --body "..." [--all] [--csv]` |
| Forward | `forward <id> [--to "..."] [--csv] [--comment "..."]` |

**Common Parameters:** `--folder` (inbox/sent/drafts), `--limit N`, `--unread`, `--from/--to/--subject/--body`

**Note:** `list` and `search` commands automatically show preview by default.

**Correct Command Examples:**
```bash
# List today's emails
py -3 email_operations.py list --filter "receivedDateTime ge 2026-03-23T00:00:00Z"

# Find specific email (ONE-STEP: search + display full content)
py -3 email_operations.py find --from "sender@example.com" --subject "keyword"

# Search with multiple criteria
py -3 email_operations.py list --from "sender@example.com" --subject "keyword"

# Get email by ID
py -3 email_operations.py get <message_id>
```

**Common Errors to AVOID:**
```bash
# WRONG: --json after subcommand
py -3 email_operations.py list --json --filter "..."  # Error

# WRONG: Unquoted multi-word arguments
py -3 email_operations.py find --subject Annual Enrollment  # Error

# WRONG: Using list + get when find works
py -3 email_operations.py list --json | extract ID | get <id>  # Inefficient
```

**LLM Command Selection Strategy:**
1. View specific email → `find --from/--to/--subject` (ONE-STEP)
2. Browse multiple emails → `list --from/--to/--subject` (with previews + IDs)
3. Known message ID → `get <message_id>` (direct access)
4. NEVER use `list --json` → extract ID → `get <id>` when `find` works in one step

**Auto-Features:**
- `--top` alias: `--top N` works as alias for `--limit N`
- Smart pattern detection: Auto-detects natural language patterns
  - English: "sent to X", "received from X"
  - Chinese: "发给 X", "收到/来自 X"
- Self-sent email warning: `reply` warns when replying to your own sent email
- Auto-fallback: `find` falls back to `list --preview` if search API fails
- Attachment preservation: `forward` uses Graph API's native endpoint
- Line break handling: `forward --comment` converts newlines to HTML `<br>` tags

**Manual Notes:**
- `reply --all` preserves all original To + CC recipients (excludes self)
- Both `reply` and `forward` auto-include conversation history in HTML body

**CSV Support (Mass Mailing):**
- `send --csv`, `reply --csv`, `forward --csv` support CSV for BCC recipients
- Auto-splits into batches of 500 if needed
- CSV column auto-detected (email, Email, etc.)

```bash
# Send with CSV (recipients go to BCC)
python scripts/email_operations.py send --csv recipients.csv --subject "..." --body "..."

# Forward with CSV
python scripts/email_operations.py forward <id> --csv recipients.csv
```

### 3. Calendar Operations

**Script:** `scripts/calendar_operations.py`

| Action | Command |
|--------|---------|
| List | `list --limit 10` |
| Get | `get <event_id>` |
| Create | `create --subject "..." --start "..." --end "..." [--attendees] [--location] [--body]` |
| Update | `update <event_id> --subject "..."` |
| Delete | `delete <event_id>` |
| Free/Busy | `freebusy --emails "..." --start "..." --end "..."` |

### 4. User Operations

**Script:** `scripts/user_operations.py`

| Action | Command |
|--------|---------|
| Get Me | `get` |
| Search | `search "john" [--limit] [--name-only] [--office] [--detail]` |
| Manager | `manager <user_id>` |
| Reports | `reports <user_id>` |
| Find Times | `find-times --attendees "..." --duration 60` |

**Search Implementation:**
- Email address search: Detects '@' in query, uses exact match (`eq`)
- Name search: Uses `startsWith` for partial matching
- Example: `search "user@ibm.com"` uses exact match, `search "John"` uses partial match

## Notes
- All commands support `--json` flag (place BEFORE subcommand)
- Max 500 recipients per email (To + CC + BCC)
- Rate limits apply - add delays for bulk operations
