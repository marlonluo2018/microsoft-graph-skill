---
name: microsoft-graph-skill
description: |
  Microsoft Graph API for email, calendar, and user operations.  
  **Triggers:** "read my emails", "send an email", "check calendar", "create meeting".
  **⚠️ MUST read full SKILL.md before using this skill.**
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
| List/Search | `list --from "..." --subject "..." --top 25 --timezone "Asia/Shanghai"` ⚠️ **--timezone REQUIRED** |
| Get Full | `get <message_id>` |
| Send | `send --to "..." --subject "..." --body "..." [--attachments "file.pdf"]` |
| Reply | `reply <id> --body "..." [--sender-only] [--attachments "file.pdf"]` |
| Forward | `forward <id> --to "..."` |
| Folders | `folders` |
| Attachments | `attachments <message_id> [--download] [--id <att_id>]` |
| Accept Invite | `accept-invite <message_id> [--comment "..."] [--no-send]` |
| Decline Invite | `decline-invite <message_id> [--comment "..."] [--no-send]` |

**Key Options:**
- `--from` / `--to` / `--subject` / `--body` - Search fields
- `--folder <name>` - inbox/sent/drafts/deleted/all (default: inbox)
- `--top N` - Max results (default 25)
- `--detail` - Show full body instead of preview
- `--timezone` - Display timezone ⚠️ **MANDATORY for list/search/find commands** (e.g., "Asia/Shanghai", "UTC", "+08:00")
- `--since <timestamp>` / `--before <timestamp>` - Time filter (**MUST include timezone**)
- `--emails-only` - Show only regular emails (exclude meeting invites)
- `--events-only` - Show only meeting invites (exclude regular emails)
- `--attachments "path"` - Attach files to send/reply (can use multiple times, supports up to 150 MB per file)

**📎 Attachment Support:**
- **Small files (<3 MB)**: Sent inline with email
- **Large files (3-150 MB)**: Uploaded via Microsoft Graph upload sessions with progress tracking
- **Multiple files**: Use `--attachments` multiple times or separate paths with spaces
- **Example**: `send --to "user@example.com" --subject "Report" --body "See attached" --attachments "report.pdf" --attachments "data.xlsx"`

**⚠️ Time Format (CRITICAL):**
- ✅ `"2026-03-26T12:00:00+08:00"` or `"2026-03-26T04:00:00Z"`
- ❌ `"2026-03-26"` or `"2026-03-26T12:00:00"` (no timezone)

## Calendar Operations

**Script:** `scripts/calendar_operations.py`

| Action | Command |
|--------|---------|
| List | `list --limit 10 [--start "..."] [--end "..."] [--timezone "..."]` |
| Get | `get <event_id>` |
| Create | `create --subject "..." --start "..." --end "..." [--attendees] [--teams]` |
| Update | `update <event_id> --subject "..."` |
| Delete | `delete <event_id> [--permanent]` |
| Availability | `availability --emails "..." --start "..." --end "..." [--timezone "..."]` |
| Suggest | `suggest --attendees "..." --duration 60` |

**Attendee Actions:** `accept`, `decline`, `tentative`, `propose`

**Key Features:**

### Calendar List Enhancements
- **Timezone Conversion**: Displays event times in specified timezone (default: Asia/Shanghai)
- **Organizer Info**: Shows who organized each meeting
- **Smart Recommendations**: For unaccepted meetings, shows:
  - ✅ Recommend Accept: All attendees are free
  - ⚠️ Partial Conflict: Some attendees are busy
  - ❌ Not Recommend: You or key attendees are busy
- **Date Filtering**: Use `--start` and `--end` to filter events by date range

### Availability Command Features
- **Multi-Person Comparison**: Visual timeline showing availability for multiple people
- **Working Hours Integration**:
  - Fetches each person's configured working hours from Outlook
  - Converts working hours to display timezone
  - Marks out-of-office hours with ⬛ in timeline
- **Smart Meeting Suggestions**:
  - Suggests best meeting times based on:
    - Common free slots
    - Within everyone's working hours
    - Sorted by availability percentage
- **Visual Timeline**:
  - Each character = 30 minutes
  - ✅ Free, ❓ Tentative, 🔴 Busy, 🏖️ Out of Office, 💼 Working Elsewhere
- **User Display Names**: Shows full names (e.g., "Meng Ning Luo") instead of email addresses

**Example:**
```bash
# Check availability for multiple people
python scripts/calendar_operations.py availability \
  --emails "user1@company.com,user2@company.com" \
  --start "2026-03-27T09:00:00" \
  --end "2026-03-27T18:00:00" \
  --timezone "Asia/Shanghai"
```

**Output includes:**
- Working hours for each person (converted to your timezone)
- Visual timeline comparison
- Smart meeting time suggestions with availability percentages

## User Operations

**Script:** `scripts/user_operations.py`

| Action | Command |
|--------|---------|
| Get Me | `get` |
| Search | `search "john" [--detail]` |
| Manager | `manager <user_id>` |
| Reports | `reports <user_id>` |

## Meeting Invite Features

**Email list commands automatically:**
- Categorize messages into 📧 EMAILS and 📅 MEETING INVITES
- Display response status for meeting invites:
  - ✅ Accepted
  - ❌ Declined
  - ❓ Tentative
  - ⏳ Not Responded
  - 👤 Organizer

**Accept/Decline meeting invites:**
- Use `accept-invite` or `decline-invite` with message ID from email list
- Optional `--comment` to add a message to organizer
- Use `--no-send` to accept/decline without notifying organizer

## Notes
- All commands support `--json` flag
- Max 500 recipients per email
- Rate limits handled automatically
- Meeting invites are identified by `@odata.type` containing "eventMessage"
