#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Microsoft Graph User Operations Module

Provides user and contact search and query operations.

Usage:
    python user_operations.py search <query>
    python user_operations.py get <user_id_or_email>
    python user_operations.py manager <user_id_or_email>
    python user_operations.py directreports <user_id_or_email>
    python user_operations.py contacts [--search <query>]
"""

import os
import sys
import argparse
import json
from pathlib import Path
from typing import List, Optional, Dict, Any

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Add parent directory to path for config import
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import configuration and auth
from config import GRAPH_API_BASE
from auth import get_access_token

# Try to import requests
try:
    import requests
except ImportError:
    print("Error: requests package not found.")
    print("Install with: pip install requests")
    sys.exit(1)


def get_headers(token: str) -> Dict[str, str]:
    """Get authorization headers for API requests."""
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }


# =============================================================================
# SEARCH USERS
# =============================================================================

def search_users(
    query: str,
    limit: int = 25,
    token: str = None,
    search_fields: list = None,
    office: str = None
) -> List[Dict[str, Any]]:
    """
    Search for users in the organization.
    
    Args:
        query: Search query (name, email, etc.)
        limit: Maximum number of results
        token: Access token
        search_fields: Fields to search in ['displayName', 'mail', 'givenName', 'surname']
        office: Filter by office location (partial match)
    
    Returns:
        List of user objects
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/users"
    
    # Default search fields
    if search_fields is None:
        search_fields = ['displayName', 'mail', 'userPrincipalName', 'givenName', 'surname']
    
    # Build filter for searching
    filter_parts = []
    for field in search_fields:
        filter_parts.append(f"startsWith({field},'{query}')")
    
    filter_query = " or ".join(filter_parts)
    
    # Add office filter if specified (filter by email domain after getting results)
    office_lower = office.lower() if office else None
    
    params = {
        "$filter": filter_query,
        "$top": limit,
        "$select": "id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones"
    }
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to search users: {response.status_code} - {response.text}")
    
    data = response.json()
    users = data.get("value", [])
    
    # Client-side filter by office/email domain
    if office_lower:
        users = [
            u for u in users
            if office_lower in (u.get('mail') or '').lower()
            or office_lower in (u.get('officeLocation') or '').lower()
        ]
    
    return users[:limit]


def list_users(
    limit: int = 25,
    filter_query: str = None,
    token: str = None
) -> List[Dict[str, Any]]:
    """
    List users in the organization.
    
    Args:
        limit: Maximum number of results
        filter_query: OData filter query
        token: Access token
    
    Returns:
        List of user objects
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/users"
    
    params = {
        "$top": limit,
        "$select": "id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation"
    }
    
    if filter_query:
        params["$filter"] = filter_query
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to list users: {response.status_code} - {response.text}")
    
    data = response.json()
    return data.get("value", [])


# =============================================================================
# GET USER
# =============================================================================

def get_user(user_id: str, token: str = None) -> Dict[str, Any]:
    """
    Get a specific user by ID or email.
    
    Args:
        user_id: User ID or email/UPN
        token: Access token
    
    Returns:
        User object
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/users/{user_id}"
    
    response = requests.get(url, headers=get_headers(token))
    
    if response.status_code != 200:
        raise Exception(f"Failed to get user: {response.status_code} - {response.text}")
    
    return response.json()


def get_me(token: str = None) -> Dict[str, Any]:
    """
    Get the current authenticated user.
    
    Args:
        token: Access token
    
    Returns:
        Current user object
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me"
    
    response = requests.get(url, headers=get_headers(token))
    
    if response.status_code != 200:
        raise Exception(f"Failed to get current user: {response.status_code} - {response.text}")
    
    return response.json()


# =============================================================================
# MANAGER
# =============================================================================

def get_manager(user_id: str = None, token: str = None) -> Dict[str, Any]:
    """
    Get the manager of a user.
    
    Args:
        user_id: User ID or email (uses 'me' if not provided)
        token: Access token
    
    Returns:
        Manager user object
    """
    if token is None:
        token = get_access_token()
    
    if user_id:
        url = f"{GRAPH_API_BASE}/users/{user_id}/manager"
    else:
        url = f"{GRAPH_API_BASE}/me/manager"
    
    response = requests.get(url, headers=get_headers(token))
    
    if response.status_code == 404:
        return None  # No manager assigned
    elif response.status_code != 200:
        raise Exception(f"Failed to get manager: {response.status_code} - {response.text}")
    
    return response.json()


# =============================================================================
# DIRECT REPORTS
# =============================================================================

def get_direct_reports(user_id: str = None, token: str = None) -> List[Dict[str, Any]]:
    """
    Get direct reports of a user.
    
    Args:
        user_id: User ID or email (uses 'me' if not provided)
        token: Access token
    
    Returns:
        List of direct report user objects
    """
    if token is None:
        token = get_access_token()
    
    if user_id:
        url = f"{GRAPH_API_BASE}/users/{user_id}/directReports"
    else:
        url = f"{GRAPH_API_BASE}/me/directReports"
    
    params = {
        "$select": "id,displayName,mail,userPrincipalName,jobTitle,department"
    }
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to get direct reports: {response.status_code} - {response.text}")
    
    data = response.json()
    return data.get("value", [])


# =============================================================================
# CONTACTS (PERSONAL)
# =============================================================================

def list_contacts(
    folder_id: str = None,
    limit: int = 25,
    token: str = None
) -> List[Dict[str, Any]]:
    """
    List personal contacts.
    
    Args:
        folder_id: Specific contact folder ID
        limit: Maximum number of results
        token: Access token
    
    Returns:
        List of contact objects
    """
    if token is None:
        token = get_access_token()
    
    if folder_id:
        url = f"{GRAPH_API_BASE}/me/contactFolders/{folder_id}/contacts"
    else:
        url = f"{GRAPH_API_BASE}/me/contacts"
    
    params = {
        "$top": limit,
        "$select": "id,displayName,emailAddresses,mobilePhone,companyName"
    }
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to list contacts: {response.status_code} - {response.text}")
    
    data = response.json()
    return data.get("value", [])


def search_contacts(
    query: str,
    limit: int = 25,
    token: str = None
) -> List[Dict[str, Any]]:
    """
    Search personal contacts.
    
    Args:
        query: Search query
        limit: Maximum number of results
        token: Access token
    
    Returns:
        List of matching contact objects
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/contacts"
    
    filter_query = f"startsWith(displayName,'{query}') or contains(emailAddresses/any(a:a/address),'{query}')"
    
    params = {
        "$filter": filter_query,
        "$top": limit,
        "$select": "id,displayName,emailAddresses,mobilePhone,companyName"
    }
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to search contacts: {response.status_code} - {response.text}")
    
    data = response.json()
    return data.get("value", [])


def get_contact(contact_id: str, token: str = None) -> Dict[str, Any]:
    """
    Get a specific contact by ID.
    
    Args:
        contact_id: Contact ID
        token: Access token
    
    Returns:
        Contact object
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/contacts/{contact_id}"
    
    response = requests.get(url, headers=get_headers(token))
    
    if response.status_code != 200:
        raise Exception(f"Failed to get contact: {response.status_code} - {response.text}")
    
    return response.json()


# =============================================================================
# CONTACT FOLDERS
# =============================================================================

def list_contact_folders(token: str = None) -> List[Dict[str, Any]]:
    """List all contact folders."""
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/contactFolders"
    
    response = requests.get(url, headers=get_headers(token))
    
    if response.status_code != 200:
        raise Exception(f"Failed to list contact folders: {response.status_code} - {response.text}")
    
    data = response.json()
    return data.get("value", [])


# =============================================================================
# PEOPLE (SUGGESTED)
# =============================================================================

def get_people(
    query: str = None,
    limit: int = 25,
    token: str = None
) -> List[Dict[str, Any]]:
    """
    Get people relevant to the user (suggested contacts).
    
    Args:
        query: Optional search query
        limit: Maximum number of results
        token: Access token
    
    Returns:
        List of person objects
    """
    if token is None:
        token = get_access_token()
    
    url = f"{GRAPH_API_BASE}/me/people"
    
    params = {
        "$top": limit,
        "$select": "id,displayName,emailAddresses,jobTitle,department"
    }
    
    if query:
        params["$search"] = f'"{query}"'
    
    response = requests.get(url, headers=get_headers(token), params=params)
    
    if response.status_code != 200:
        raise Exception(f"Failed to get people: {response.status_code} - {response.text}")
    
    data = response.json()
    return data.get("value", [])


# =============================================================================
# DISPLAY HELPERS
# =============================================================================

def display_user_list(users: List[Dict]):
    """Display a list of users in a readable format."""
    print(f"\n{'='*120}")
    print(f"{'Name':<25} {'Email':<35} {'Title':<30} {'Office':<20} {'Phone':<18}")
    print(f"{'='*120}")
    
    for user in users:
        name = (user.get('displayName') or 'Unknown')[:25]
        email = (user.get('mail') or user.get('userPrincipalName') or '')[:35]
        title = (user.get('jobTitle') or '')[:30]
        office = (user.get('officeLocation') or '')[:20]
        phone = (user.get('mobilePhone') or '')[:18]
        
        print(f"{name:<25} {email:<35} {title:<30} {office:<20} {phone:<18}")
    
    print(f"{'='*120}")
    print(f"Total: {len(users)} users")


def display_user(user: Dict):
    """Display a single user in detail."""
    print(f"\n{'='*80}")
    print(f"Display Name: {user.get('displayName', 'N/A')}")
    print(f"Email: {user.get('mail') or user.get('userPrincipalName', 'N/A')}")
    print(f"User ID: {user.get('id', 'N/A')}")
    print(f"Job Title: {user.get('jobTitle', 'N/A')}")
    print(f"Department: {user.get('department', 'N/A')}")
    print(f"Office Location: {user.get('officeLocation', 'N/A')}")
    print(f"Mobile Phone: {user.get('mobilePhone', 'N/A')}")
    print(f"Business Phones: {user.get('businessPhones', [])}")
    print(f"{'='*80}")


def display_contact_list(contacts: List[Dict]):
    """Display a list of contacts."""
    print(f"\n{'='*80}")
    print(f"{'Name':<35} {'Email':<35} {'Phone':<20}")
    print(f"{'='*80}")
    
    for contact in contacts:
        name = contact.get('displayName', 'Unknown')[:35]
        emails = contact.get('emailAddresses', [])
        email = emails[0].get('address', '') if emails else ''
        phone = contact.get('mobilePhone', '')[:20]
        
        print(f"{name:<35} {email:<35} {phone:<20}")
    
    print(f"{'='*80}")
    print(f"Total: {len(contacts)} contacts")


def display_contact(contact: Dict):
    """Display a single contact in detail."""
    print(f"\n{'='*80}")
    print(f"Name: {contact.get('displayName', 'N/A')}")
    print(f"Contact ID: {contact.get('id', 'N/A')}")
    print(f"Company: {contact.get('companyName', 'N/A')}")
    print(f"Job Title: {contact.get('jobTitle', 'N/A')}")
    
    emails = contact.get('emailAddresses', [])
    if emails:
        print(f"\nEmail Addresses:")
        for e in emails:
            print(f"  - {e.get('name', '')} <{e.get('address', '')}>")
    
    phones = contact.get('businessPhones', [])
    mobile = contact.get('mobilePhone')
    home = contact.get('homePhones', [])
    
    if phones or mobile or home:
        print(f"\nPhones:")
        for p in phones:
            print(f"  Business: {p}")
        if mobile:
            print(f"  Mobile: {mobile}")
        for p in home:
            print(f"  Home: {p}")
    
    print(f"{'='*80}")


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Microsoft Graph User Operations")
    subparsers = parser.add_subparsers(dest="command", required=True)
    
    # Global --json flag
    parser.add_argument("--json", action="store_true", help="Output in JSON format")
    
    # Search users command
    search_parser = subparsers.add_parser("search", help="Search users")
    search_parser.add_argument("query", help="Search query")
    search_parser.add_argument("--limit", type=int, default=25, help="Max results")
    search_parser.add_argument("--name-only", action="store_true", help="Search only by given name (first name)")
    search_parser.add_argument("--office", help="Filter by office location (e.g., 'Philippines', 'PH')")
    search_parser.add_argument("--detail", action="store_true", help="Show detailed information for each user")
    
    # Get user command
    get_parser = subparsers.add_parser("get", help="Get a user")
    get_parser.add_argument("user_id", nargs="?", help="User ID or email (optional, defaults to 'me')")
    
    # Manager command
    manager_parser = subparsers.add_parser("manager", help="Get user's manager")
    manager_parser.add_argument("user_id", nargs="?", help="User ID or email")
    
    # Direct reports command
    reports_parser = subparsers.add_parser("directreports", help="Get direct reports")
    reports_parser.add_argument("user_id", nargs="?", help="User ID or email")
    
    # Contacts commands
    contacts_parser = subparsers.add_parser("contacts", help="List contacts")
    contacts_parser.add_argument("--search", dest="query", help="Search query")
    contacts_parser.add_argument("--folder", help="Folder ID")
    contacts_parser.add_argument("--limit", type=int, default=25, help="Max results")
    
    # People command
    people_parser = subparsers.add_parser("people", help="Get suggested people")
    people_parser.add_argument("--search", dest="query", help="Search query")
    people_parser.add_argument("--limit", type=int, default=25, help="Max results")
    
    # Contact folders command
    subparsers.add_parser("folders", help="List contact folders")
    
    args = parser.parse_args()
    
    try:
        if args.command == "search":
            search_fields = ['givenName'] if args.name_only else None
            users = search_users(args.query, args.limit, search_fields=search_fields, office=args.office)
            if args.json:
                print(json.dumps({"success": True, "users": users, "total": len(users)}, indent=2, default=str))
            elif args.detail:
                for user in users:
                    display_user(user)
            else:
                display_user_list(users)
        
        elif args.command == "get":
            if args.user_id:
                user = get_user(args.user_id)
            else:
                user = get_me()
            if args.json:
                print(json.dumps({"success": True, "user": user}, indent=2, default=str))
            else:
                display_user(user)
        
        elif args.command == "manager":
            manager = get_manager(args.user_id)
            if args.json:
                print(json.dumps({"success": True, "manager": manager}, indent=2, default=str))
            elif manager:
                display_user(manager)
            else:
                print("No manager assigned.")
        
        elif args.command == "directreports":
            reports = get_direct_reports(args.user_id)
            if args.json:
                print(json.dumps({"success": True, "directReports": reports, "total": len(reports)}, indent=2, default=str))
            else:
                display_user_list(reports)
        
        elif args.command == "contacts":
            if args.query:
                contacts = search_contacts(args.query, args.limit)
            else:
                contacts = list_contacts(args.folder, args.limit)
            if args.json:
                print(json.dumps({"success": True, "contacts": contacts, "total": len(contacts)}, indent=2, default=str))
            else:
                display_contact_list(contacts)
        
        elif args.command == "people":
            people = get_people(args.query, args.limit)
            if args.json:
                print(json.dumps({"success": True, "people": people, "total": len(people)}, indent=2, default=str))
            else:
                display_user_list(people)
        
        elif args.command == "folders":
            folders = list_contact_folders()
            if args.json:
                print(json.dumps({"success": True, "folders": folders, "total": len(folders)}, indent=2, default=str))
            else:
                print(f"\nContact Folders ({len(folders)}):")
                for f in folders:
                    print(f"  - {f.get('displayName', 'Unknown')} (ID: {f.get('id')})")
    
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
