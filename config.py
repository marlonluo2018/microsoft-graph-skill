#!/usr/bin/env python3
"""
Microsoft Graph Skill Configuration

This file contains all configurable settings for the Microsoft Graph skill.
You can modify these values or override them with environment variables.
"""

import os
from pathlib import Path

# =============================================================================
# Azure AD / Microsoft Graph Configuration
# =============================================================================

# Tenant ID
# - "organizations" - for company/organizational accounts only
# - "common" - for both personal and organizational accounts
# - Specific tenant ID - for single-tenant applications
TENANT_ID = os.environ.get("MS_GRAPH_TENANT_ID", "organizations")

# Client ID (Application ID)
# Default: Microsoft Office public client ID (pre-authorized for Graph API)
# You can use your own Azure AD application client ID if needed
DEFAULT_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
CLIENT_ID = os.environ.get("MS_GRAPH_CLIENT_ID", DEFAULT_CLIENT_ID)

# Scopes for Microsoft Graph API
# Using .default scope to request all permissions configured for the app
DEFAULT_SCOPES = ["https://graph.microsoft.com/.default"]

# Alternative: Specify individual scopes
# DEFAULT_SCOPES = [
#     "User.Read",
#     "Mail.Read",
#     "Mail.ReadWrite",
#     "Mail.Send",
#     "Calendars.Read",
#     "Calendars.ReadWrite",
#     "Calendars.Read.Shared",
#     "Calendars.ReadWrite.Shared",
#     "Contacts.Read",
#     "People.Read",
# ]

# =============================================================================
# API Configuration
# =============================================================================

# Microsoft Graph API base URL
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

# API request timeout (seconds)
API_TIMEOUT = 30

# Maximum retry attempts for API calls
MAX_RETRIES = 3

# =============================================================================
# Email Configuration
# =============================================================================

# Maximum recipients per email (company policy)
MAX_RECIPIENTS_PER_EMAIL = 500

# Default email body type
DEFAULT_BODY_TYPE = "html"  # "html" or "text"

# =============================================================================
# Cache Configuration
# =============================================================================

# Directory for storing cached tokens and device flow data
CACHE_DIR = Path.home() / ".ms_graph_skill"

# Token cache file
TOKEN_CACHE_FILE = CACHE_DIR / "tokens.json"

# Device flow cache file
DEVICE_FLOW_FILE = CACHE_DIR / "device_flow.json"

# =============================================================================
# Display Configuration
# =============================================================================

# Maximum characters to display for email body in thread view
MAX_BODY_DISPLAY_LENGTH = 1000

# Maximum characters to display for email body in single message view
MAX_MESSAGE_DISPLAY_LENGTH = 2000

# Date format for displaying dates
DATE_FORMAT = "%Y-%m-%d %H:%M"

# =============================================================================
# Helper Functions
# =============================================================================

def ensure_cache_dir():
    """Ensure the cache directory exists."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)


def get_client_id() -> str:
    """Get the client ID, checking environment variable first."""
    return os.environ.get("MS_GRAPH_CLIENT_ID", DEFAULT_CLIENT_ID)


def get_tenant_id() -> str:
    """Get the tenant ID, checking environment variable first."""
    return os.environ.get("MS_GRAPH_TENANT_ID", TENANT_ID)


def get_scopes() -> list:
    """Get the scopes for Microsoft Graph API."""
    return DEFAULT_SCOPES
