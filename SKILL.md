---
name: microsoft-graph-skill
description: |
  This skill enables Microsoft Graph API operations for email, calendar, and user management.
  Use when the user requests: reading/sending emails, checking/creating calendar events,
  or searching user information. Supports timezone-aware operations with natural time
  parameters like "--end now".
---

# Microsoft Graph Skill

**All commands require email addresses only** (e.g., `user@example.com`). Names are NOT supported.

## Quick Reference

### Authentication (`scripts/auth.py`)
`--status` `--start` `--complete` `--logout`

### Email (`scripts/email_operations.py`)
`list` `search` `find` `get` `thread` `send` `reply` `forward` `read` `delete` `folders` `attachments` `accept-invite` `decline-invite`

### Calendar (`scripts/calendar_operations.py`)
`list` `get` `create` `update` `delete` `cancel` `forward` `availability` `propose` `accept` `decline` `tentative` `calendars`

### User (`scripts/user_operations.py`)
`get` `search` `manager` `directreports` `contacts` `people` `folders`

## Time Format ⚠️

**CRITICAL:** `--start` and `--end` require plain datetime + `--timezone`:

```
--start "2026-03-26T12:00:00" --timezone "Asia/Shanghai"
--end "now" --timezone "UTC"
```

Embedded timezone (`Z` or `+08:00`) is NOT supported.

## Common Patterns

### List/Search/Find Emails
```bash
python scripts/email_operations.py list --timezone "Asia/Shanghai"
python scripts/email_operations.py search --from "boss@company.com" --timezone "UTC"
python scripts/email_operations.py find --subject "meeting" --start "2026-03-01T00:00:00" --end "now" --timezone "Asia/Shanghai"
```

### Send Email
```bash
python scripts/email_operations.py send --to "user@example.com" --subject "Hello" --body "Message content"
```

## Email Body Formatting

**Use `\n` for line breaks**

Example:
```bash
--body "Line 1\nLine 2\nLine 3"
--body "Hello,\n\nParagraph 1.\n\nParagraph 2."
```

### Create Calendar Event
```bash
python scripts/calendar_operations.py create --subject "Team Meeting" --start "2026-03-26T10:00:00" --end "2026-03-26T11:00:00" --timezone "Asia/Shanghai"
```

### Check Availability
```bash
python scripts/calendar_operations.py availability --emails "user1@example.com,user2@example.com" --start "2026-03-26T09:00:00" --end "2026-03-26T17:00:00" --timezone "Asia/Shanghai"
```

## Key Options

### Email Filtering
- `--from EMAIL` - Sender filter
- `--to EMAIL` - Recipient filter
- `--subject TXT` - Subject filter
- `--folder FOLDER` - Folder name (inbox, sentitems, all, etc.)
- `--unread` - Unread only
- `--detail` - Full content

### Calendar Attendees
- `--required EMAILS` - Required attendees
- `--optional EMAILS` - Optional attendees
- `--no-teams` - Disable Teams meeting (enabled by default)

### Email Recipients
- `--to EMAILS` - To recipients
- `--cc EMAILS` - CC recipients
- `--bcc EMAILS` - BCC recipients
- `--csv PATH` - Load BCC from CSV (auto-batch if >500 recipients)

## Notes
- All commands support `--json` flag
- Max 500 recipients per email (auto-batched from CSV)
- Rate limits handled automatically
- **For detailed parameters, see `references/commands.md`**
