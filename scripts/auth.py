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
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any
from functools import wraps
import threading

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

# =============================================================================
# Logging Setup
# =============================================================================

def setup_logging():
    """Setup logging for authentication module."""
    log_file = CACHE_DIR / "auth.log"
    log_file.parent.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger(__name__)

logger = setup_logging()


# =============================================================================
# Utilities
# =============================================================================

def validate_config() -> bool:
    """Validate that required configuration is available."""
    if not TENANT_ID:
        logger.error("TENANT_ID is not configured")
        return False
    if not CLIENT_ID and not os.environ.get("MS_GRAPH_CLIENT_ID"):
        logger.warning("CLIENT_ID is not configured, will use environment variable")
    return True

def retry_on_failure(max_retries: int = 3, delay: float = 1.0):
    """Decorator to retry function on failure."""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            last_error = None
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    last_error = e
                    if attempt < max_retries - 1:
                        logger.warning(f"{func.__name__} failed (attempt {attempt + 1}/{max_retries}): {e}. Retrying in {delay}s...")
                        time.sleep(delay * (attempt + 1))
                    else:
                        logger.error(f"{func.__name__} failed after {max_retries} attempts: {e}")
            raise last_error
        return wrapper
    return decorator


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
        self._lock = threading.Lock()  # Thread safety for token operations

        # Load tokens from disk
        self.load_tokens_from_disk()

        # Clean up expired device flow to prevent authentication errors
        self._cleanup_expired_flow()

    def _cleanup_expired_flow(self) -> None:
        """Clean up expired device flow to prevent authentication errors."""
        try:
            flow = load_device_flow()
            # Check if 'expires_at' key exists and value is set (use 'in' instead of truthy check)
            if flow and 'expires_at' in flow and flow['expires_at'] is not None:
                if time.time() > flow['expires_at']:
                    clear_device_flow()
                    logger.info("清理了过期的设备流程记录")
        except Exception as e:
            logger.warning(f"清理过期流程时出错: {e}")
    
    def load_tokens_from_disk(self) -> None:
        """Load authentication tokens from disk."""
        if not TOKEN_CACHE_FILE.exists():
            logger.debug("Token cache file does not exist")
            return

        try:
            with open(TOKEN_CACHE_FILE, "r", encoding='utf-8') as f:
                token_data = json.load(f)

            self.access_token = token_data.get("access_token")
            self.refresh_token = token_data.get("refresh_token")
            self.token_expiry = token_data.get("token_expiry", 0)
            self.authenticated = token_data.get("authenticated", False)
            self.username = token_data.get("username")

            # Check if token is expired
            if self.authenticated and self.access_token:
                if time.time() >= self.token_expiry - 60:
                    logger.info("Access token expired on load")
                    self.authenticated = False
                    self.access_token = None

            logger.debug(f"Tokens loaded from disk. Authenticated: {self.authenticated}")
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse token cache file: {e}")
        except Exception as e:
            logger.error(f"Failed to load tokens from disk: {e}")
    
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
        try:
            with open(TOKEN_CACHE_FILE, "w", encoding='utf-8') as f:
                json.dump(token_data, f, indent=2)
            logger.debug("Tokens saved to disk")
        except Exception as e:
            logger.error(f"Failed to save tokens to disk: {e}")
            raise
    
    def update_token(
        self,
        access_token: str,
        expires_in: int = 3600,
        refresh_token: Optional[str] = None,
        username: Optional[str] = None,
    ) -> None:
        """Update the access token and related information (thread-safe)."""
        with self._lock:
            self.access_token = access_token
            self.token_expiry = time.time() + expires_in
            self.refresh_token = refresh_token or self.refresh_token
            self.authenticated = True
            self.username = username or self.username
            logger.info(f"Token updated for user: {self.username}, expires in {expires_in}s")
            self.save_tokens_to_disk()
    
    def clear_tokens(self) -> None:
        """Clear all authentication tokens (thread-safe)."""
        with self._lock:
            self.access_token = None
            self.token_expiry = 0
            self.refresh_token = None
            self.authenticated = False
            self.username = None
            if TOKEN_CACHE_FILE.exists():
                try:
                    TOKEN_CACHE_FILE.unlink()
                    logger.info("Tokens cleared successfully")
                except Exception as e:
                    logger.error(f"Failed to clear token cache file: {e}")
    
    def is_token_valid(self) -> bool:
        """Check if the current token is valid and not expired."""
        if not self.authenticated or not self.access_token:
            return False

        # Add 60 seconds buffer to ensure token doesn't expire during use
        is_valid = time.time() < self.token_expiry - 60

        if not is_valid:
            logger.debug("Token is expired or about to expire")

        return is_valid
    
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
    try:
        with open(DEVICE_FLOW_FILE, "w", encoding='utf-8') as f:
            json.dump(flow, f, indent=2)
        logger.debug("Device flow saved to disk")
    except Exception as e:
        logger.error(f"Failed to save device flow: {e}")
        raise


def load_device_flow() -> Optional[Dict[str, Any]]:
    """Load device flow from disk."""
    if not DEVICE_FLOW_FILE.exists():
        return None

    try:
        with open(DEVICE_FLOW_FILE, "r", encoding='utf-8') as f:
            flow = json.load(f)
            logger.debug("Device flow loaded from disk")
            return flow
    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse device flow file: {e}")
        return None
    except Exception as e:
        logger.error(f"Failed to load device flow: {e}")
        return None


def clear_device_flow() -> None:
    """Clear device flow from disk."""
    if DEVICE_FLOW_FILE.exists():
        try:
            DEVICE_FLOW_FILE.unlink()
            logger.debug("Device flow cleared")
        except Exception as e:
            logger.error(f"Failed to clear device flow file: {e}")


# =============================================================================
# Authentication Functions
# =============================================================================

def create_app(client_id: str) -> PublicClientApplication:
    """Create MSAL PublicClientApplication."""
    if not TENANT_ID:
        raise ValueError("TENANT_ID is not configured")

    logger.debug(f"Creating MSAL app for client_id: {client_id}")
    return PublicClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )


@retry_on_failure(max_retries=3, delay=1.0)
def start_auth_flow(client_id: str = None, scopes: list = None) -> Dict[str, Any]:
    """
    Start device code flow and return URL and code.
    Does NOT wait for completion.

    Returns:
        dict: {"url": "...", "code": "..."} or {"error": "..."}
    """
    # Validate configuration
    if not validate_config():
        return {"error": "Configuration error: TENANT_ID not set"}

    if client_id is None:
        client_id = CLIENT_ID or os.environ.get("MS_GRAPH_CLIENT_ID")

    if not client_id:
        return {"error": "Client ID is required. Set CLIENT_ID in config.py or MS_GRAPH_CLIENT_ID environment variable"}

    if scopes is None:
        scopes = DEFAULT_SCOPES

    logger.info(f"Starting auth flow for client_id: {client_id}")

    # Check if already authenticated
    token_manager = TokenManager()
    if token_manager.is_token_valid():
        logger.info(f"Already authenticated as {token_manager.username}")
        return {
            "status": "already_authenticated",
            "message": "Already authenticated",
            "username": token_manager.username,
            "expires": datetime.fromtimestamp(token_manager.token_expiry).isoformat(),
            "expires_in": token_manager.get_token_expiry_info()['display']
        }

    app = create_app(client_id)

    # Initiate device code flow
    flow = app.initiate_device_flow(scopes=scopes)

    if 'verification_uri' not in flow:
        error_msg = flow.get('error_description', flow.get('error', 'Failed to initiate flow'))
        logger.error(f"Failed to initiate device flow: {error_msg}")
        return {"error": error_msg}

    # Add expires_at timestamp
    expires_in = flow.get('expires_in', 900)
    flow['expires_at'] = time.time() + expires_in
    flow['client_id'] = client_id
    flow['scopes'] = scopes

    # Save flow to disk
    save_device_flow(flow)

    logger.info(f"Device flow started. User code: {flow['user_code']}, expires in {expires_in}s")

    return {
        "url": flow['verification_uri'],
        "code": flow['user_code'],
        "message": f"To sign in, open {flow['verification_uri']} and enter code {flow['user_code']}",
        "expires_in": expires_in
    }


@retry_on_failure(max_retries=3, delay=1.0)
def complete_auth_flow() -> Dict[str, Any]:
    """
    Complete the pending authentication flow.

    Returns:
        dict: {"success": True, ...} or {"error": "..."} or {"status": "pending", ...}
    """
    flow = load_device_flow()

    if not flow:
        logger.warning("No pending authentication flow found")
        return {"error": "No pending authentication flow. Run --start first."}

    # Check if flow expired
    if 'expires_at' in flow and flow['expires_at'] is not None and time.time() > flow['expires_at']:
        clear_device_flow()
        logger.warning("Device flow expired")
        return {"error": "Authentication flow expired. Please start again with --start."}

    client_id = flow.get('client_id', CLIENT_ID)
    scopes = flow.get('scopes', DEFAULT_SCOPES)

    logger.info(f"Completing auth flow for client_id: {client_id}")

    app = create_app(client_id)
    token_manager = TokenManager()

    # Acquire token using device flow
    result = app.acquire_token_by_device_flow(flow)

    if 'access_token' in result:
        # Get account info
        accounts = app.get_accounts()
        username = accounts[0].get('username', 'Unknown') if accounts else 'Unknown'

        logger.info(f"Authentication successful for user: {username}")

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
        error_description = result.get('error_description', error)

        if error == 'authorization_pending':
            logger.debug("Authorization pending, waiting for user...")
            return {"status": "pending", "message": "Waiting for user to complete authentication..."}
        elif error == 'expired':
            clear_device_flow()
            logger.warning("Device flow expired during completion")
            return {"error": "Authentication expired. Please start again."}
        elif error == 'authorization_declined':
            clear_device_flow()
            logger.warning("Authorization declined by user")
            return {"error": "Authentication was declined. Please try again."}
        else:
            logger.error(f"Authentication failed: {error_description}")
            return {"error": error_description, "code": error}


@retry_on_failure(max_retries=3, delay=1.0)
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
        logger.warning("No refresh token available")
        return {"error": "No refresh token available. Please login again with --start."}

    logger.info(f"Refreshing token for user: {token_manager.username}")

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

        logger.info(f"Token refreshed successfully for user: {username}")

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
        error_description = result.get('error_description', error)

        if error == 'invalid_grant':
            logger.warning("Refresh token invalid or expired, clearing tokens")
            token_manager.clear_tokens()
            return {"error": "Refresh token expired or invalid. Please login again with --start."}

        logger.error(f"Failed to refresh token: {error_description}")
        return {"error": error_description, "code": error}


def check_status(client_id: str = None) -> Dict[str, Any]:
    """
    Check authentication status with automatic token refresh.

    Args:
        client_id: Client ID (optional, for token refresh)

    Returns:
        dict: Authentication status information
    """
    logger.debug("Checking authentication status")
    token_manager = TokenManager()

    if token_manager.is_token_valid():
        expiry_info = token_manager.get_token_expiry_info()
        logger.info(f"Authenticated as {token_manager.username}, token expires in {expiry_info['display']}")
        return {
            "authenticated": True,
            "username": token_manager.username,
            "expires": datetime.fromtimestamp(token_manager.token_expiry).isoformat(),
            "expires_in": expiry_info['display']
        }
    elif token_manager.refresh_token:
        # Token expired but we have a refresh token, try to refresh
        logger.info("Token expired, attempting to refresh...")
        refresh_result = refresh_token(client_id)
        if refresh_result.get('success'):
            return {
                "authenticated": True,
                "username": refresh_result.get('username'),
                "expires": refresh_result.get('expires'),
                "expires_in": refresh_result.get('expires_in'),
                "refreshed": True,
                "message": "令牌已自动刷新"
            }
        else:
            error_msg = refresh_result.get('error', 'Unknown error')
            logger.warning(f"Token refresh failed: {error_msg}")
            return {
                "authenticated": False,
                "message": f"令牌过期且刷新失败: {error_msg}",
                "action": "请重新登录: python auth.py --start"
            }
    else:
        logger.info("Not authenticated - no token available")
        return {
            "authenticated": False,
            "message": "Not authenticated or token expired.",
            "action": "请运行: python auth.py --start"
        }


def logout() -> Dict[str, Any]:
    """Clear cached tokens."""
    logger.info("Logging out...")
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
        client_id: Azure AD application client ID (optional, for token refresh)
        scopes: List of scopes (optional, for token refresh)

    Returns:
        str: Access token or None if not authenticated
    """
    token_manager = TokenManager()

    # If token is valid, return it
    if token_manager.is_token_valid():
        logger.debug("Returning valid access token")
        return token_manager.access_token

    # If token expired but we have a refresh token, try to refresh
    if token_manager.refresh_token:
        logger.info("Token expired, attempting refresh...")
        refresh_result = refresh_token(client_id, scopes)
        if refresh_result.get('success'):
            logger.info("Token refreshed successfully")
            return token_manager.access_token
        else:
            logger.error("Failed to refresh token")

    logger.warning("No valid access token available")
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
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose logging")

    args = parser.parse_args()

    # Set logging level based on verbose flag
    if args.verbose:
        logger.setLevel(logging.DEBUG)

    result = {}

    try:
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
    except KeyboardInterrupt:
        logger.info("Operation cancelled by user")
        result = {"error": "Operation cancelled"}
    except Exception as e:
        logger.exception(f"Unexpected error: {e}")
        result = {"error": f"Unexpected error: {str(e)}"}

    # Output as JSON
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
