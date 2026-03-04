"""Shared fixtures for SimpleOrgChart tests."""

from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List
from unittest.mock import patch

import pytest

# Ensure the project root is importable
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

# Set required env vars BEFORE any app module import triggers app_main
os.environ.setdefault("ADMIN_PASSWORD", "test-password-for-ci-only")
os.environ.setdefault("SECRET_KEY", "test-secret-key")


@pytest.fixture()
def sample_employees() -> List[Dict[str, Any]]:
    """Small employee list for hierarchy tests."""
    return [
        {
            "id": "1",
            "name": "Alice CEO",
            "email": "alice@example.com",
            "title": "Chief Executive Officer",
            "department": "Executive",
            "managerId": None,
        },
        {
            "id": "2",
            "name": "Bob VP",
            "email": "bob@example.com",
            "title": "VP Engineering",
            "department": "Engineering",
            "managerId": "1",
        },
        {
            "id": "3",
            "name": "Carol Dev",
            "email": "carol@example.com",
            "title": "Software Engineer",
            "department": "Engineering",
            "managerId": "2",
        },
        {
            "id": "4",
            "name": "Dave Orphan",
            "email": "dave@example.com",
            "title": "Analyst",
            "department": "Finance",
            "managerId": None,
        },
    ]


@pytest.fixture()
def sample_login_records() -> List[Dict[str, Any]]:
    """Records for last-login filter tests."""
    return [
        {
            "name": "Active User",
            "email": "active@example.com",
            "accountEnabled": True,
            "licenseCount": 1,
            "userType": "Member",
            "mailboxType": "user",
            "neverSignedIn": False,
            "daysSinceLastActivity": 5,
        },
        {
            "name": "Disabled User",
            "email": "disabled@example.com",
            "accountEnabled": False,
            "licenseCount": 1,
            "userType": "Member",
            "mailboxType": "user",
            "neverSignedIn": False,
            "daysSinceLastActivity": 120,
        },
        {
            "name": "Guest User",
            "email": "guest@example.com",
            "accountEnabled": True,
            "licenseCount": 0,
            "userType": "Guest",
            "mailboxType": "user",
            "neverSignedIn": False,
            "daysSinceLastActivity": 30,
        },
        {
            "name": "Shared Mailbox",
            "email": "shared@example.com",
            "accountEnabled": True,
            "licenseCount": 0,
            "userType": "Member",
            "mailboxType": "shared",
            "isSharedMailbox": True,
            "neverSignedIn": True,
            "daysSinceLastActivity": None,
        },
        {
            "name": "Room Mailbox",
            "email": "room@example.com",
            "accountEnabled": True,
            "licenseCount": 0,
            "userType": "Member",
            "mailboxType": "room",
            "neverSignedIn": True,
            "daysSinceLastActivity": None,
        },
        {
            "name": "Unlicensed Member",
            "email": "unlicensed@example.com",
            "accountEnabled": True,
            "licenseCount": 0,
            "userType": "Member",
            "mailboxType": "user",
            "neverSignedIn": False,
            "daysSinceLastActivity": 60,
        },
    ]
