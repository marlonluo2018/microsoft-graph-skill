# Microsoft Graph API Permissions Reference

This document outlines the required permissions for different Microsoft Graph operations.

## Permission Types

### Delegated Permissions
Used when an app is signed in by a user. The app acts on behalf of the signed-in user.

### Application Permissions
Used when the app runs as a background service without a signed-in user.

---

## Mail Permissions

| Operation | Minimum Permission | Description |
|-----------|-------------------|-------------|
| Read user's mail | `Mail.Read` | Read user's emails |
| Read user's mail (shared) | `Mail.Read.Shared` | Read emails in shared folders |
| Send mail as user | `Mail.Send` | Send emails as the user |
| Send mail (shared) | `Mail.Send.Shared` | Send from shared mailboxes |
| Read and write mail | `Mail.ReadWrite` | Read, update, delete emails |

### Required Scopes for This Skill

```
Mail.Read           - List and read emails
Mail.ReadWrite      - Mark as read, move, delete emails
Mail.Send           - Send, reply, forward emails
```

---

## Calendar Permissions

| Operation | Minimum Permission | Description |
|-----------|-------------------|-------------|
| Read user's calendars | `Calendars.Read` | Read calendar events |
| Read shared calendars | `Calendars.Read.Shared` | Read shared calendar events |
| Read and write calendars | `Calendars.ReadWrite` | Create, update, delete events |
| Write shared calendars | `Calendars.ReadWrite.Shared` | Write to shared calendars |

### Required Scopes for This Skill

```
Calendars.Read           - List and read events
Calendars.ReadWrite      - Create, update, delete events
Calendars.Read.Shared    - Read shared calendars
Calendars.ReadWrite.Shared - Write to shared calendars
```

---

## User Permissions

| Operation | Minimum Permission | Description |
|-----------|-------------------|-------------|
| Read signed-in user's profile | `User.Read` | Basic profile information |
| Read all users' basic info | `User.Read.All` | Read all users in organization |
| Read all users' full info | `User.ReadBasic.All` | Basic info for all users |

### Required Scopes for This Skill

```
User.Read        - Read current user's profile
User.Read.All    - Search and query other users (delegated)
```

---

## Contact Permissions

| Operation | Minimum Permission | Description |
|-----------|-------------------|-------------|
| Read user's contacts | `Contacts.Read` | Read personal contacts |
| Read and write contacts | `Contacts.ReadWrite` | Create, update, delete contacts |

### Required Scopes for This Skill

```
Contacts.Read      - List and search contacts
Contacts.ReadWrite - Create, update, delete contacts
```

---

## People Permissions

| Operation | Minimum Permission | Description |
|-----------|-------------------|-------------|
| Read people | `People.Read` | Access suggested people |
| Read all people | `People.Read.All` | Access all relevant people |

### Required Scopes for This Skill

```
People.Read    - Get suggested people
```

---

## Complete Permission List for This Skill

### Recommended Scopes

```
# User
User.Read

# Mail
Mail.Read
Mail.ReadWrite
Mail.Send

# Calendar
Calendars.Read
Calendars.ReadWrite
Calendars.Read.Shared
Calendars.ReadWrite.Shared

# Contacts
Contacts.Read
Contacts.ReadWrite

# People
People.Read
```

### Minimal Scopes (Read-Only Operations)

```
User.Read
Mail.Read
Calendars.Read
Contacts.Read
People.Read
```

---

## Admin Consent Requirements

Some permissions require **admin consent** in Azure AD:

| Permission | Admin Consent Required |
|------------|----------------------|
| Mail.Read | No |
| Mail.ReadWrite | No |
| Mail.Send | No |
| Calendars.Read | No |
| Calendars.ReadWrite | No |
| User.Read.All | **Yes** |
| Contacts.Read | No |
| People.Read | No |

If you need to search users (`User.Read.All`), an administrator must grant consent in the Azure portal.

---

## Configuring Permissions in Azure AD

### Step 1: Register Application

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Enter app name and select supported account types
5. Note the **Application (client) ID**

### Step 2: Add Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Delegated permissions**
5. Add the required permissions listed above
6. Click **Grant admin consent** if required

### Step 3: Enable Device Code Flow

1. Go to **Authentication**
2. Add platform: **Mobile and desktop applications**
3. Select the redirect URI for device code flow (usually no specific URI needed)
4. The client will use `https://login.microsoftonline.com/common/oauth2/nativeclient`

---

## Testing Permissions

### Check Current User's Permissions

```bash
# Get current user info (requires User.Read)
GET /me

# List inbox (requires Mail.Read)
GET /me/mailFolders/inbox/messages

# Create event (requires Calendars.ReadWrite)
POST /me/events

# Search users (requires User.Read.All)
GET /users?$filter=startsWith(displayName,'test')
```

---

## Permission Error Handling

### Insufficient Permissions Error

```json
{
  "error": {
    "code": "ErrorAccessDenied",
    "message": "Access is denied. Check credentials and try again."
  }
}
```

**Solution:** Add the required permission scope and re-authenticate.

### Admin Consent Required

```json
{
  "error": {
    "code": "ErrorInsufficientPrivileges",
    "message": "Insufficient privileges to complete the operation."
  }
}
```

**Solution:** Request admin consent in Azure AD portal.

---

## Least Privilege Principle

Always request the minimum permissions needed for your operations:

| Use Case | Minimal Permissions |
|----------|---------------------|
| Read emails only | `Mail.Read` |
| Send emails only | `Mail.Send` |
| Read calendar only | `Calendars.Read` |
| Manage calendar | `Calendars.ReadWrite` |
| Search users | `User.Read.All` (requires admin) |
| Read own profile | `User.Read` |

---

## Scopes String Format

When requesting tokens, combine scopes with spaces:

```
User.Read Mail.Read Mail.Send Calendars.ReadWrite
```

Example authorization URL:
```
https://login.microsoftonline.com/common/oauth2/v2.0/authorize
  ?client_id={client_id}
  &response_type=code
  &scope=User.Read%20Mail.Read%20Mail.Send%20Calendars.ReadWrite
  &redirect_uri={redirect_uri}
```
