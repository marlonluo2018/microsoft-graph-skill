# Microsoft Graph API Commands Reference

Complete parameter documentation for all commands.

## Authentication (`scripts/auth.py`)

| Argument | Type | Description |
|----------|------|-------------|
| `--status` | flag | Check authentication status |
| `--start` | flag | Start auth flow (output URL and code) |
| `--complete` | flag | Complete auth flow |
| `--logout` | flag | Clear cached tokens |
| `--client-id ID` | optional | Azure AD application client ID |
| `--verbose`, `-v` | flag | Enable verbose logging |

## Email Operations (`scripts/email_operations.py`)

### Commands

| Command | Description | Required Args |
|---------|-------------|---------------|
| `list` | List emails | `--timezone TZ` |
| `search` | Search emails | `--timezone TZ` |
| `find` | Find emails | `--timezone TZ` |
| `get <message_id>` | Get full email | message_id |
| `thread <message_id>` | View email thread | message_id |
| `send` | Send email | `--to`, `--subject`, `--body` |
| `reply <message_id>` | Reply to email | message_id, `--body` |
| `forward <message_id>` | Forward email | message_id, `--to` |
| `read <message_id>` | Mark as read/unread | message_id |
| `delete <message_id>` | Delete email | message_id |
| `folders` | List mail folders | - |
| `attachments <message_id>` | List attachments | message_id |
| `accept-invite <message_id>` | Accept meeting invite | message_id |
| `decline-invite <message_id>` | Decline meeting invite | message_id |

### Filtering Options (list/search/find)

| Argument | Description |
|----------|-------------|
| `--from EMAIL` | Filter by sender email |
| `--to EMAIL` | Filter by recipient email |
| `--subject TXT` | Filter by subject text |
| `--body TXT` | Filter by body text |
| `--folder FOLDER` | Mail folder (inbox, sentitems, drafts, all, etc.) |
| `--start TS` | Start datetime (with --timezone) |
| `--end TS` | End datetime (with --timezone, supports "now") |
| `--limit N` | Max results (default: 25) |
| `--filter QUERY` | OData filter query |
| `--unread` | Show unread only |
| `--preview` | Show body preview |
| `--detail` | Show full body content |
| `--focused` | Focused inbox only |
| `--other` | Other inbox only |
| `--emails-only` | Exclude calendar events |
| `--events-only` | Exclude regular emails |

### Send Options

| Argument | Description |
|----------|-------------|
| `--to EMAILS` | To recipients (comma-separated or multiple) |
| `--cc EMAILS` | CC recipients |
| `--bcc EMAILS` | BCC recipients |
| `--subject TXT` | Subject (required) |
| `--body TXT` | Body (required) |
| `--body-type TYPE` | "html" or "text" (default: html) |
| `--attachments PATH` | Attach files (multiple allowed) |
| `--csv PATH` | Load BCC recipients from CSV |
| `--email-column NAME` | CSV email column name |

### Reply Options

| Argument | Description |
|----------|-------------|
| `--body TXT` | Reply body (required) |
| `--sender-only` | Reply to sender only (default: reply all) |
| `--to EMAILS` | Replace To recipients |
| `--cc EMAILS` | Replace CC recipients |
| `--add-to EMAILS` | Add To recipients |
| `--add-cc EMAILS` | Add CC recipients |
| `--bcc EMAILS` | BCC recipients |
| `--attachments PATH` | Attach files |
| `--csv PATH` | Load BCC from CSV |
| `--email-column NAME` | CSV column name |
| `--importance LEVEL` | "low", "normal", or "high" |

### Forward Options

| Argument | Description |
|----------|-------------|
| `--to EMAILS` | Recipients (required) |
| `--cc EMAILS` | CC recipients |
| `--bcc EMAILS` | BCC recipients |
| `--comment TXT` | Comment to add |
| `--csv PATH` | Load BCC from CSV |
| `--email-column NAME` | CSV column name |

### Other Email Options

**read:** `--unread` (mark as unread)

**folders:** `--all` (include hidden)

**attachments:** `--download`, `--save-dir DIR`, `--id ID`

**accept-invite / decline-invite:** `--comment TXT`, `--no-send`

## Calendar Operations (`scripts/calendar_operations.py`)

### Commands

| Command | Description | Required Args |
|---------|-------------|---------------|
| `list` | List events | `--timezone TZ` |
| `get <event_id>` | Get event details | event_id |
| `create` | Create event | `--subject`, `--start`, `--end`, `--timezone` |
| `update <event_id>` | Update event | event_id, `--timezone` |
| `delete <event_id>` | Delete event | event_id |
| `cancel <event_id>` | Cancel event (organizer only) | event_id |
| `forward <event_id>` | Forward event | event_id, `--to` |
| `availability` | Check availability | `--emails`, `--start`, `--end`, `--timezone` |
| `propose <event_id>` | Propose new time | event_id, `--start`, `--end`, `--timezone` |
| `accept <event_id>` | Accept event | event_id |
| `decline <event_id>` | Decline event | event_id |
| `tentative <event_id>` | Tentatively accept | event_id |
| `calendars` | List all calendars | - |

### List Options

| Argument | Description |
|----------|-------------|
| `--calendar ID` | Calendar ID |
| `--start TS` | Start datetime |
| `--end TS` | End datetime (supports "now") |
| `--limit N` | Max events (default: 25) |

### Create Options

| Argument | Description |
|----------|-------------|
| `--subject TXT` | Subject (required) |
| `--start TS` | Start datetime (required) |
| `--end TS` | End datetime (required) |
| `--timezone TZ` | Timezone (required) |
| `--body TXT` | Description |
| `--location TXT` | Location |
| `--required EMAILS` | Required attendees |
| `--optional EMAILS` | Optional attendees |
| `--all-day` | All day event |
| `--no-teams` | Disable Teams meeting |

### Update Options

| Argument | Description |
|----------|-------------|
| `--subject TXT` | New subject |
| `--start TS` | New start |
| `--end TS` | New end |
| `--timezone TZ` | Timezone |
| `--body TXT` | New description |
| `--location TXT` | New location |
| `--required EMAILS` | Required attendees |
| `--optional EMAILS` | Optional attendees |

### Other Calendar Options

**delete:** `--permanent` (cannot recover)

**cancel:** `--comment TXT` (message to attendees)

**forward:** `--to EMAILS` (required), `--comment TXT`

**availability:** `--emails EMAILS` (required), `--start TS`, `--end TS`, `--timezone TZ`

**propose:** `--start TS`, `--end TS`, `--timezone TZ`, `--comment TXT`

**accept/decline/tentative:** `--comment TXT`, `--no-send`

## User Operations (`scripts/user_operations.py`)

### Commands

| Command | Description | Required Args |
|---------|-------------|---------------|
| `get [user_id]` | Get user info | - (defaults to 'me') |
| `search QUERY` | Search users | query |
| `manager [user_id]` | Get manager | - |
| `directreports [user_id]` | Get direct reports | - |
| `contacts` | Search contacts | - |
| `people` | Search people | - |
| `folders` | List contact folders | - |

### Search Options

| Argument | Description |
|----------|-------------|
| `--limit N` | Max results (default: 25) |
| `--name-only` | Search by first name only |
| `--office LOCATION` | Filter by office |
| `--detail` | Show detailed info |

### Contacts Options

| Argument | Description |
|----------|-------------|
| `--search QUERY` | Search query |
| `--folder ID` | Folder ID |
| `--limit N` | Max results (default: 25) |

### People Options

| Argument | Description |
|----------|-------------|
| `--search QUERY` | Search query |
| `--limit N` | Max results (default: 25) |

## Global Options

- All commands support `--json` for JSON output
- Max 500 recipients per email
- Rate limits handled automatically with retry

## CSV File Support

- Load BCC recipients from CSV files
- Auto-detects email column
- Auto-batching when total recipients exceed 500
- Example format: single column with header `email`
