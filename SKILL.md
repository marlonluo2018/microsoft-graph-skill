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

**Robustness Improvements (v2.0):**
- ✅ **Automatic cleanup**: Expired device flows are automatically cleaned on startup to prevent authentication errors
- ✅ **Comprehensive logging**: All operations are logged to `cache/auth.log` for debugging
- ✅ **Retry mechanism**: Network operations automatically retry (3 attempts) with exponential backoff
- ✅ **Thread-safe operations**: Token operations are protected with locks to prevent race conditions
- ✅ **Configuration validation**: Validates required config (TENANT_ID, CLIENT_ID) before operations
- ✅ **Enhanced error messages**: More detailed error messages with suggested actions (e.g., "请重新登录: python auth.py --start")
- ✅ **Graceful error handling**: Catches and logs exceptions instead of crashing
- ✅ **UTF-8 encoding support**: Proper file encoding for international characters
- ✅ **Verbose mode**: Add `--verbose` flag for detailed debug logs

### 2. Email Operations

**Script:** `scripts/email_operations.py`

**⚠️ CRITICAL: Command Format Rules (READ FIRST)**
1. **Subcommand FIRST, then options**: `py -3 email_operations.py <subcommand> [options]`
2. **NO --json after subcommand**: Use `py -3 email_operations.py --json <subcommand>` if needed
3. **Quote multi-word arguments**: `--subject "Annual Enrollment"` NOT `--subject Annual Enrollment`

| Action | Command |
|--------|---------|
| List/Search | `list [options]` or `search [options]` (identical) |
| Find | `find --from/--to/--subject [--folder]` |
| Get | `get <message_id>` |
| Send | `send --to "..." --subject "..." --body "..."` |
| Reply | `reply <id> --body "..." [--all]` |
| Forward | `forward <id> --to "..." --comment "..."` |

**Common Parameters:** `--folder` (inbox/sent/drafts), `--limit N`, `--unread`, `--from/--to/--subject/--body`
**⚠️ Note:** `list` and `search` commands **automatically show preview** by default - no need to add `--preview` flag.

**✅ Correct Command Examples:**
```bash
# List today's emails (preview is automatic)
py -3 email_operations.py list --filter "receivedDateTime ge 2026-03-23T00:00:00Z"

# Find specific email (ONE-STEP: search + display full content)
py -3 email_operations.py find --from "sender@example.com" --subject "keyword"

# Search with multiple criteria (preview is automatic)
py -3 email_operations.py list --from "sender@example.com" --subject "keyword"

# Get email by ID (when ID is known)
py -3 email_operations.py get <message_id>
```

**❌ Common Errors to AVOID:**
```bash
# WRONG: --json after subcommand
py -3 email_operations.py list --json --filter "..."  # ❌ Error: unrecognized arguments

# WRONG: Unquoted multi-word arguments
py -3 email_operations.py find --subject Annual Enrollment  # ❌ Error: unrecognized arguments

# WRONG: Using list + get when find works
py -3 email_operations.py list --json | extract ID | get <id>  # ❌ Inefficient, use find instead
```

**LLM Command Selection Strategy:**
1. **View specific email** → `find --from/--to/--subject` (ONE-STEP: search + display full content)
2. **Browse multiple emails** → `list --from/--to/--subject` (list with **automatic previews** + **message IDs**)
3. **Known message ID** → `get <message_id>` (direct access)
4. **NEVER** use `list --json` → extract ID → `get <id>` when `find` works in one step

**Best Practice - Email Workflow Efficiency:**
- Use `list` for **detailed view with automatic previews** + **message IDs** (default behavior)
- Only use `get <message_id>` when full email body is needed
- **All list/find commands now display message IDs** - no need to extract from JSON
- Reason: Message IDs are displayed directly, saving query steps for quick access

**Auto-Features (Implemented in Code):**
- **`--top` alias**: `--top N` works as alias for `--limit N`
- **Smart pattern detection**: Automatically detects natural language patterns and adjusts folder/parameters
  - English: "sent to X", "send to X", "received from X", "from X"
  - Chinese: "发给/发送给/发送到 X", "收到/来自/从 X"
  - Example: `--from "sent to john@example.com"` → searches sent folder for emails to john@example.com
- **Self-sent email warning**: `reply` warns when replying to your own sent email, suggests using `forward`
- **Auto-fallback**: `find` automatically falls back to `list --preview` if search API fails
- **Attachment preservation**: `forward` uses Graph API's native endpoint to preserve attachments
- **Line break handling**: `forward --comment` automatically converts newlines to HTML `<br>` tags

**Manual Notes:**
- `reply --all` preserves all original To + CC recipients (excludes self)
- Both `reply` and `forward` auto-include conversation history in HTML body

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

**Search Implementation:**
- Email address search: Detects '@' in query, uses exact match (`eq`) for email fields
- Name search: Uses `startsWith` for partial matching
- Example: `search "user@ibm.com"` uses exact match, `search "John"` uses partial match
| Manager | `manager <user_id>` |
| Reports | `reports <user_id>` |
| Find Times | `find-times --attendees "..." --duration 60` |

## Notes
- All commands support `--json` flag (place BEFORE subcommand)
- Max 500 recipients per email (To + CC + BCC)
- Rate limits apply - add delays for bulk operations
