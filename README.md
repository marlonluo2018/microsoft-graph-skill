# Microsoft Graph Skill

A comprehensive Python skill for interacting with Microsoft Graph API, providing email, calendar, and user management capabilities.

## Features

- **Authentication**: OAuth2 device code flow with automatic token refresh and robust error handling
- **Email Operations**: Read, send, reply, forward emails with CC/BCC support
- **Email Search**: Search emails by sender, recipient, subject, or content
- **Email Find**: One-step find and display specific email (combines search + get)
- **Email Threads**: View complete conversation threads
- **Smart Pattern Detection**: Auto-detect natural language patterns (e.g., "sent to X", "发给 X")
- **Calendar Management**: Create, update, delete events; query free/busy time
- **User Operations**: Search users, get manager/direct reports, contacts, and suggested people

## Installation

```bash
# Install dependencies
pip install -r requirements.txt
```

## Quick Start

### 1. Authentication

```bash
# Check authentication status (auto-refreshes token if needed)
python scripts/auth.py --status

# If not authenticated, start login flow
python scripts/auth.py --start
# Visit the URL and enter the code, then:
python scripts/auth.py --complete

# Logout
python scripts/auth.py --logout

# Verbose mode for debugging
python scripts/auth.py --status --verbose
```

**Note:** Tokens auto-refresh via `--status`. Manual `--refresh` is optional.

**Robustness Features:**
- Automatic cleanup of expired device flows
- Network operations retry (3 attempts) with backoff
- Thread-safe token operations
- Comprehensive logging to `~/.ms_graph_skill/auth.log`

### 2. Email Operations

```bash
# List recent emails (preview is automatic)
python scripts/email_operations.py list --limit 10

# Search emails from a sender
python scripts/email_operations.py search --from "John"

# Find and display specific email (ONE-STEP: search + display full content)
python scripts/email_operations.py find --from "sender@example.com" --subject "keyword"

# Get email details
python scripts/email_operations.py get <message_id>

# View email thread
python scripts/email_operations.py thread <message_id>

# Send email
python scripts/email_operations.py send \
  --to "recipient@example.com" \
  --subject "Test Subject" \
  --body "Email body content"

# Reply to email
python scripts/email_operations.py reply <message_id> --body "Reply content"

# Forward email (preserves attachments)
python scripts/email_operations.py forward <message_id> \
  --to "recipient@example.com" \
  --comment "FYI"
```

**Smart Pattern Detection:**
The email operations automatically detect natural language patterns:
- English: "sent to X", "received from X"
- Chinese: "发给 X", "收到/来自 X"
- Example: `--from "sent to john@example.com"` → searches sent folder for emails to john@example.com

### 3. Calendar Operations

```bash
# List upcoming events
python scripts/calendar_operations.py list --limit 10

# Get event details
python scripts/calendar_operations.py get <event_id>

# Create event
python scripts/calendar_operations.py create \
  --subject "Meeting" \
  --start "2026-03-22T10:00:00" \
  --end "2026-03-22T11:00:00" \
  --attendees "colleague@example.com"

# Query free/busy time
python scripts/calendar_operations.py freebusy \
  --emails "user1@example.com,user2@example.com" \
  --start "2026-03-22T09:00:00" \
  --end "2026-03-22T18:00:00"
```

### 4. User Operations

```bash
# Get current user info
python scripts/user_operations.py get

# Get specific user info
python scripts/user_operations.py get <user_id_or_email>

# Search users
python scripts/user_operations.py search "john"

# Search with detailed output
python scripts/user_operations.py search "john" --detail

# Search by first name only
python scripts/user_operations.py search "john" --name-only

# Filter by office location
python scripts/user_operations.py search "john" --office "Philippines"

# Get user's manager
python scripts/user_operations.py manager <user_id>

# Get direct reports
python scripts/user_operations.py directreports <user_id>

# Get suggested people
python scripts/user_operations.py people --search "john"

# List contacts
python scripts/user_operations.py contacts --limit 10

# Search contacts
python scripts/user_operations.py contacts --search "john"

# List contact folders
python scripts/user_operations.py folders
```

## Configuration

Edit `config.py` to customize settings:

```python
# Use your own Azure AD application
CLIENT_ID = "your-client-id"

# Or set via environment variable
export MS_GRAPH_CLIENT_ID="your-client-id"

# Change tenant
TENANT_ID = "organizations"  # or "common" or specific tenant ID
```

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `MS_GRAPH_CLIENT_ID` | Azure AD application client ID | Microsoft Office public client |
| `MS_GRAPH_TENANT_ID` | Azure AD tenant ID | `organizations` |

## Project Structure

```
microsoft-graph-skill/
├── config.py              # Configuration settings
├── requirements.txt       # Python dependencies
├── README.md              # This file
├── SKILL.md               # Skill definition for AI assistant
├── scripts/
│   ├── auth.py            # Authentication module
│   ├── email_operations.py    # Email operations
│   ├── calendar_operations.py # Calendar operations
│   └── user_operations.py     # User/contact operations
└── references/
    ├── api_endpoints.md   # Microsoft Graph API endpoints
    └── permissions.md     # Required API permissions
```

## Authentication Details

This skill uses OAuth2 device code flow, which is ideal for:
- Command-line applications
- Environments without a browser
- Headless/automated scenarios

The default configuration uses Microsoft Office's public client ID, which is pre-authorized for Microsoft Graph API access. This works with organizational accounts without needing to register a new Azure AD application.

### Token Caching

Tokens are cached locally in `~/.ms_graph_skill/`:
- `tokens.json` - Access and refresh tokens
- `device_flow.json` - Pending device flow data

### Token Refresh

Access tokens expire after ~1 hour. The skill automatically refreshes tokens when running `--status` or any API operation. Refresh tokens are valid for 14-90 days depending on Azure AD configuration.

## API Permissions

When using the default client ID, the following delegated permissions are available:

| Permission | Description |
|------------|-------------|
| `User.Read` | Read user profile |
| `Mail.Read` | Read user mail |
| `Mail.ReadWrite` | Read and write user mail |
| `Mail.Send` | Send mail as user |
| `Calendars.Read` | Read user calendars |
| `Calendars.ReadWrite` | Read and write user calendars |
| `Calendars.Read.Shared` | Read shared calendars |
| `Contacts.Read` | Read user contacts |
| `People.Read` | Read people relevant to user |

## JSON Output Format

All scripts support `--json` flag for machine-readable output, ideal for AI agents and automation:

```bash
# JSON output for email operations
python scripts/email_operations.py --json list --limit 10
python scripts/email_operations.py --json search --from "John"

# JSON output for calendar operations
python scripts/calendar_operations.py --json list --limit 10
python scripts/calendar_operations.py --json get <event_id>

# JSON output for user operations
python scripts/user_operations.py --json search "john"
python scripts/user_operations.py --json manager <user_id>
```

JSON output provides structured data without formatting, colors, or progress messages.

## Limitations

- Maximum 500 recipients per email (company policy)
- Microsoft Graph API has rate limits (throttling)
- Some operations require additional permissions

## Troubleshooting

### "Not authenticated" error

Run authentication first:
```bash
python scripts/auth.py --start
# Follow the instructions, then:
python scripts/auth.py --complete
```

### Token expired

Tokens auto-refresh when running `--status`. If refresh fails, re-login:
```bash
python scripts/auth.py --start
python scripts/auth.py --complete
```

### "InefficientFilter" error

Microsoft Graph API has limitations on `$filter` queries. The search functionality uses a combination of `$search` and client-side filtering to work around this.

## Dependencies

- `msal` - Microsoft Authentication Library
- `requests` - HTTP library for API calls

## License

Internal use only.
