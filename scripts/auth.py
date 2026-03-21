#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Microsoft Graph Authentication Module

Provides OAuth2 device code flow authentication for Microsoft Graph API.
Tokens are cached locally for subsequent use.

Usage:
    python auth.py --start              # Start auth flow, output URL and code as JSON
    python auth.py --complete           # Complete auth flow (after user enters code)
    python auth.py --status             # Check authentication status
    python auth.py --logout             # Clear cached tokens
"""

import os
import sys
import json
import time
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Add parent directory to path for config import
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import configuration
from config import (
    TENANT_ID, CLIENT_ID, DEFAULT_SCOPES,
    CACHE_DIR, TOKEN_CACHE_FILE, DEVICE_FLOW_FILE,
    ensure_cache_dir, get_client_id
)

# Try to import msal, provide guidance if not available
try:
    from msal import PublicClientApplication
except ImportError:
    print(json.dumps({"error": "msal package not found. Install with: pip install msal"}))
    sys.exit(1)


def ensure_cache_dir():
    """Ensure cache directory exists."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)


# =============================================================================
# Token Management
# =============================================================================

class TokenManager:
    """Manages authentication tokens for Microsoft Graph API."""
    
    def __init__(self):
        self.access_token: Optional[str] = None
        self.token_expiry: float = 0
        self.refresh_token: Optional[str] = None
        self.authenticated: bool = False
        self.username: Optional[str] = None
        self.load_tokens_from_disk()
    
    def load_tokens_from_disk(self) -> None:
        """Load authentication tokens from disk."""
        if not TOKEN_CACHE_FILE.exists():
            return
        
        try:
            with open(TOKEN_CACHE_FILE, "r") as f:
                token_data = json.load(f)
            
            self.access_token = token_data.get("access_token")
            self.refresh_token = token_data.get("refresh_token")
            self.token_expiry = token_data.get("token_expiry", 0)
            self.authenticated = token_data.get("authenticated", False)
            self.username = token_data.get("username")
            
            # Check if token is expired
            if self.authenticated and self.access_token:
                if time.time() >= self.token_expiry - 60:
                    self.authenticated = False
                    self.access_token = None
        except Exception:
            pass
    
    def save_tokens_to_disk(self) -> None:
        """Save authentication tokens to disk."""
        ensure_cache_dir()
        token_data = {
            "access_token": self.access_token,
            "refresh_token": self.refresh_token,
            "token_expiry": self.token_expiry,
            "authenticated": self.authenticated,
            "username": self.username,
        }
        with open(TOKEN_CACHE_FILE, "w") as f:
            json.dump(token_data, f, indent=2)
    
    def update_token(
        self,
        access_token: str,
        expires_in: int = 3600,
        refresh_token: Optional[str] = None,
        username: Optional[str] = None,
    ) -> None:
        """Update the access token and related information."""
        self.access_token = access_token
        self.token_expiry = time.time() + expires_in
        self.refresh_token = refresh_token or self.refresh_token
        self.authenticated = True
        self.username = username or self.username
        self.save_tokens_to_disk()
    
    def clear_tokens(self) -> None:
        """Clear all authentication tokens."""
        self.access_token = None
        self.token_expiry = 0
        self.refresh_token = None
        self.authenticated = False
        self.username = None
        if TOKEN_CACHE_FILE.exists():
            TOKEN_CACHE_FILE.unlink()
    
    def is_token_valid(self) -> bool:
        """Check if the current token is valid and not expired."""
        if not self.authenticated or not self.access_token:
            return False
        return time.time() < self.token_expiry - 60
    
    def get_token_expiry_info(self) -> dict:
        """Get token expiry information."""
        remaining_seconds = int(self.token_expiry - time.time())
        remaining_minutes = remaining_seconds // 60
        remaining_hours = remaining_minutes // 60
        remaining_minutes_display = remaining_minutes % 60
        
        if remaining_hours > 0:
            display = f"{remaining_hours}h {remaining_minutes_display}m"
        else:
            display = f"{remaining_minutes_display}m"
        
        return {
            "seconds": remaining_seconds,
            "display": display,
        }


# =============================================================================
# Device Flow Management
# =============================================================================

def save_device_flow(flow: Dict[str, Any]) -> None:
    """Save device flow to disk."""
    ensure_cache_dir()
    with open(DEVICE_FLOW_FILE, "w") as f:
        json.dump(flow, f, indent=2)


def load_device_flow() -> Optional[Dict[str, Any]]:
    """Load device flow from disk."""
    if not DEVICE_FLOW_FILE.exists():
        return None
    try:
        with open(DEVICE_FLOW_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return None


def clear_device_flow() -> None:
    """Clear device flow from disk."""
    if DEVICE_FLOW_FILE.exists():
        DEVICE_FLOW_FILE.unlink()


# =============================================================================
# Authentication Functions
# =============================================================================

def create_app(client_id: str) -> PublicClientApplication:
    """Create MSAL PublicClientApplication."""
    return PublicClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )


def start_auth_flow(client_id: str = None, scopes: list = None) -> Dict[str, Any]:
    """
    Start device code flow and return URL and code.
    Does NOT wait for completion.
    
    Returns:
        dict: {"url": "...", "code": "..."} or {"error": "..."}
    """
    if client_id is None:
        client_id = CLIENT_ID or os.environ.get("MS_GRAPH_CLIENT_ID")
    
    if not client_id:
        return {"error": "Client ID is required"}
    
    if scopes is None:
        scopes = DEFAULT_SCOPES
    
    # Check if already authenticated
    token_manager = TokenManager()
    if token_manager.is_token_valid():
        return {
            "status": "already_authenticated",
            "message": "Already authenticated",
            "username": token_manager.username,
            "expires": datetime.fromtimestamp(token_manager.token_expiry).isoformat()
        }
    
    app = create_app(client_id)
    
    # Initiate device code flow
    flow = app.initiate_device_flow(scopes=scopes)
    
    if 'verification_uri' not in flow:
        return {"error": flow.get('error_description', flow.get('error', 'Failed to initiate flow'))}
    
    # Add expires_at timestamp
    expires_in = flow.get('expires_in', 900)
    flow['expires_at'] = time.time() + expires_in
    flow['client_id'] = client_id
    flow['scopes'] = scopes
    
    # Save flow to disk
    save_device_flow(flow)
    
    return {
        "url": flow['verification_uri'],
        "code": flow['user_code'],
        "message": f"To sign in, open {flow['verification_uri']} and enter code {flow['user_code']}"
    }


def complete_auth_flow() -> Dict[str, Any]:
    """
    Complete the pending authentication flow.
    
    Returns:
        dict: {"success": True, ...} or {"error": "..."} or {"status": "pending", ...}
    """
    flow = load_device_flow()
    
    if not flow:
        return {"error": "No pending authentication flow. Run --start first."}
    
    # Check if flow expired
    if flow.get('expires_at') and time.time() > flow['expires_at']:
        clear_device_flow()
        return {"error": "Authentication flow expired. Please start again with --start."}
    
    client_id = flow.get('client_id', CLIENT_ID)
    scopes = flow.get('scopes', DEFAULT_SCOPES)
    
    app = create_app(client_id)
    token_manager = TokenManager()
    
    # Acquire token using device flow
    result = app.acquire_token_by_device_flow(flow)
    
    if 'access_token' in result:
        # Get account info
        accounts = app.get_accounts()
        username = accounts[0].get('username', 'Unknown') if accounts else 'Unknown'
        
        # Update token manager
        token_manager.update_token(
            access_token=result['access_token'],
            expires_in=int(result.get('expires_in', 3600)),
            refresh_token=result.get('refresh_token'),
            username=username,
        )
        
        # Clear device flow
        clear_device_flow()
        
        expiry_info = token_manager.get_token_expiry_info()
        
        return {
            "success": True,
            "username": username,
            "expires": datetime.fromtimestamp(token_manager.token_expiry).isoformat(),
            "expires_in": expiry_info['display']
        }
    else:
        error = result.get('error', 'unknown')
        if error == 'authorization_pending':
            return {"status": "pending", "message": "Waiting for user to complete authentication..."}
        elif error == 'expired':
            clear_device_flow()
            return {"error": "Authentication expired. Please start again."}
        else:
            return {"error": result.get('error_description', error)}


def refresh_token(client_id: str = None, scopes: list = None) -> Dict[str, Any]:
    """
    Refresh the access token using the refresh token.
    
    Returns:
        dict: {"success": True, ...} or {"error": "..."}
    """
    if client_id is None:
        client_id = CLIENT_ID or os.environ.get("MS_GRAPH_CLIENT_ID")
    
    if scopes is None:
        scopes = DEFAULT_SCOPES
    
    token_manager = TokenManager()
    
    if not token_manager.refresh_token:
        return {"error": "No refresh token available. Please login again with --start."}
    
    app = create_app(client_id)
    
    # Build the token request with refresh token
    result = app.acquire_token_by_refresh_token(
        token_manager.refresh_token,
        scopes=scopes
    )
    
    if 'access_token' in result:
        # Get account info
        accounts = app.get_accounts()
        username = accounts[0].get('username', token_manager.username) if accounts else token_manager.username
        
        # Update token manager
        token_manager.update_token(
            access_token=result['access_token'],
            expires_in=int(result.get('expires_in', 3600)),
            refresh_token=result.get('refresh_token'),
            username=username,
        )
        
        expiry_info = token_manager.get_token_expiry_info()
        
        return {
            "success": True,
            "username": username,
            "expires": datetime.fromtimestamp(token_manager.token_expiry).isoformat(),
            "expires_in": expiry_info['display'],
            "message": "Token refreshed successfully."
        }
    else:
        error = result.get('error', 'unknown')
        if error == 'invalid_grant':
            token_manager.clear_tokens()
            return {"error": "Refresh token expired. Please login again with --start."}
        return {"error": result.get('error_description', error)}


def check_status(client_id: str = None) -> Dict[str, Any]:
    """Check authentication status."""
    token_manager = TokenManager()
    
    if token_manager.is_token_valid():
        expiry_info = token_manager.get_token_expiry_info()
        return {
            "authenticated": True,
            "username": token_manager.username,
            "expires": datetime.fromtimestamp(token_manager.token_expiry).isoformat(),
            "expires_in": expiry_info['display']
        }
    elif token_manager.refresh_token:
        # Token expired but we have a refresh token, try to refresh
        refresh_result = refresh_token(client_id)
        if refresh_result.get('success'):
            return {
                "authenticated": True,
                "username": refresh_result.get('username'),
                "expires": refresh_result.get('expires'),
                "expires_in": refresh_result.get('expires_in'),
                "refreshed": True
            }
        else:
            return {"authenticated": False, "message": f"Token expired. {refresh_result.get('error', 'Please login again.')}"}
    else:
        return {"authenticated": False, "message": "Not authenticated or token expired."}


def logout() -> Dict[str, Any]:
    """Clear cached tokens."""
    clear_device_flow()
    token_manager = TokenManager()
    token_manager.clear_tokens()
    return {"success": True, "message": "Logged out successfully."}


def get_access_token(client_id: str = None, scopes: list = None) -> Optional[str]:
    """
    Get access token for Microsoft Graph API.
    
    This function is used by other scripts to get the access token.
    Automatically refreshes the token if expired and a refresh token exists.
    Returns None if not authenticated.
    
    Args:
        client_id: Azure AD application client ID (not used, kept for compatibility)
        scopes: List of scopes (not used, kept for compatibility)
    
    Returns:
        str: Access token or None if not authenticated
    """
    token_manager = TokenManager()
    
    # If token is valid, return it
    if token_manager.is_token_valid():
        return token_manager.access_token
    
    # If token expired but we have a refresh token, try to refresh
    if token_manager.refresh_token:
        refresh_result = refresh_token(client_id, scopes)
        if refresh_result.get('success'):
            return token_manager.access_token
    
    return None


def get_token_manager() -> TokenManager:
    """
    Get the TokenManager instance.
    
    Returns:
        TokenManager: The token manager instance
    """
    return TokenManager()


# =============================================================================
# CLI Entry Point
# =============================================================================

def main():
    """Main entry point for command-line usage."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Microsoft Graph Authentication")
    parser.add_argument("--start", action="store_true", help="Start auth flow (output URL and code)")
    parser.add_argument("--complete", action="store_true", help="Complete auth flow")
    parser.add_argument("--status", action="store_true", help="Check authentication status")
    parser.add_argument("--refresh", action="store_true", help="Refresh access token")
    parser.add_argument("--logout", action="store_true", help="Clear cached tokens")
    parser.add_argument("--client-id", help="Azure AD application client ID")
    
    args = parser.parse_args()
    
    result = {}
    
    if args.logout:
        result = logout()
    elif args.status:
        result = check_status(args.client_id)
    elif args.refresh:
        result = refresh_token(args.client_id)
    elif args.start:
        result = start_auth_flow(args.client_id)
    elif args.complete:
        result = complete_auth_flow()
    else:
        # Default: check status, if not authenticated, start flow
        status = check_status(args.client_id)
        if status.get("authenticated"):
            result = status
        else:
            result = start_auth_flow(args.client_id)
    
    # Output as JSON
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
