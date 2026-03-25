# Microsoft Graph Skill - Detailed Examples

## Email Operations Examples

### Search by Sender
```bash
py -3 email_operations.py list --from "sswarupa@in.ibm.com" --top 5
py -3 email_operations.py list --from "John Smith" --top 10
```

### Search by Subject
```bash
py -3 email_operations.py list --subject "RHCSA" --top 5
```

### Search by Date (use --filter for date queries)
```bash
py -3 email_operations.py list --filter "receivedDateTime ge 2026-03-24T00:00:00Z" --top 20
```

### Search in Specific Folders
```bash
py -3 email_operations.py list --folder sent --top 10
py -3 email_operations.py find --folder drafts --subject "report"
py -3 email_operations.py search --folder deleted --from "john"
```

### Search Across ALL Folders
```bash
py -3 email_operations.py list --folder all --from "beng"
```

### List All Available Folders
```bash
py -3 email_operations.py folders
```

### Outlook Syntax Auto-Conversion
```bash
py -3 email_operations.py find --from "from:beng"  # Auto-converts to --from "beng"
```

## Attachment Examples

```bash
# List attachments
py -3 email_operations.py attachments <message_id>

# Download all attachments (default to Desktop)
py -3 email_operations.py attachments <message_id> --download

# Download to specific directory
py -3 email_operations.py attachments <message_id> --download --save-dir ~/Downloads

# Download specific attachment
py -3 email_operations.py attachments <message_id> --id <attachment_id> -d
```

## Common Mistakes to Avoid

| Wrong | Correct | Reason |
|-------|---------|--------|
| `search --query "from:email"` | `list --from "email"` | No --query parameter |
| `find "email"` | `list --from "email"` | Positional args not supported |
| `list --filter "from/..."` | `list --from "email"` | Too complex, use --from |
