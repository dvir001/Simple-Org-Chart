"""Email configuration and SMTP settings for SimpleOrgChart."""

from __future__ import annotations

import json
import logging
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List

from .config import DATA_DIR

logger = logging.getLogger(__name__)

EMAIL_CONFIG_FILE = DATA_DIR / "email_config.json"

# Default email configuration
DEFAULT_EMAIL_CONFIG: Dict[str, Any] = {
    "enabled": False,
    "recipientEmail": "",
    "frequency": "weekly",  # daily, weekly, monthly
    "dayOfWeek": "monday",  # for daily and weekly schedules
    "dayOfMonth": "first",  # first, last - for monthly schedules
    "fileTypes": ["pdf"],  # svg, png, pdf, xlsx
    "includeReports": [],  # List of report types to include as attachments
    "lastSent": None,  # ISO timestamp of last sent email
}

def get_smtp_config() -> Dict[str, Any]:
    """Load SMTP configuration from environment variables."""
    # Get encryption setting (TLS, SSL, or None)
    encryption = os.environ.get("SMTP_ENCRYPTION", "TLS").upper()
    if encryption not in ("TLS", "SSL", "NONE"):
        encryption = "TLS"  # Default to TLS if invalid value
    
    return {
        "server": os.environ.get("SMTP_SERVER", ""),
        "port": int(os.environ.get("SMTP_PORT", "587")),
        "username": os.environ.get("SMTP_USERNAME", ""),
        "password": os.environ.get("SMTP_PASSWORD", ""),
        "fromAddress": os.environ.get("SMTP_FROM_ADDRESS", ""),
        "encryption": encryption,
    }


def is_smtp_configured() -> bool:
    """Check if SMTP settings are properly configured."""
    config = get_smtp_config()
    required_fields = ["server", "port", "username", "password", "fromAddress"]
    return all(config.get(field) for field in required_fields)


def load_email_config() -> Dict[str, Any]:
    """Load email configuration from disk or return defaults."""
    if EMAIL_CONFIG_FILE.exists():
        try:
            with EMAIL_CONFIG_FILE.open("r", encoding="utf-8") as handle:
                stored = json.load(handle)
                # Merge with defaults to ensure all fields exist
                merged = DEFAULT_EMAIL_CONFIG.copy()
                merged.update(stored)
                return merged
        except Exception as error:
            logger.error("Error loading email config: %s", error)
    
    return DEFAULT_EMAIL_CONFIG.copy()


def save_email_config(config: Dict[str, Any]) -> bool:
    """Save email configuration to disk."""
    EMAIL_CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    
    # Preserve lastSent from existing on-disk config when the caller
    # (e.g. the configure page) does not supply it.
    existing = load_email_config()
    
    # Merge with defaults
    persisted = DEFAULT_EMAIL_CONFIG.copy()
    persisted.update(config)
    
    if 'lastSent' not in config:
        persisted['lastSent'] = existing.get('lastSent')
    
    # Validate file types
    valid_file_types = {"svg", "png", "pdf", "xlsx"}
    persisted["fileTypes"] = [
        ft for ft in persisted.get("fileTypes", []) 
        if ft in valid_file_types
    ]
    
    # Validate frequency
    if persisted.get("frequency") not in ("daily", "weekly", "monthly"):
        persisted["frequency"] = "weekly"
    
    # Validate day of week
    valid_days = {"monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"}
    if persisted.get("dayOfWeek", "").lower() not in valid_days:
        persisted["dayOfWeek"] = "monday"
    
    # Validate day of month
    if persisted.get("dayOfMonth") not in ("first", "last"):
        persisted["dayOfMonth"] = "first"
    
    logger.info("Saving email configuration to: %s", EMAIL_CONFIG_FILE)
    try:
        with EMAIL_CONFIG_FILE.open("w", encoding="utf-8") as handle:
            json.dump(persisted, handle, indent=2)
        logger.info("Email configuration saved successfully")
        return True
    except Exception as error:
        logger.error("Error saving email config to %s: %s", EMAIL_CONFIG_FILE, error)
        return False


def get_report_types() -> List[str]:
    """Get list of available report types for email attachments."""
    return [
        "missing_manager",
        "disabled_with_license",
        "filtered_with_license",
        "filtered_users",
        "disabled_users",
        "last_login",
        "recently_disabled",
        "recently_hired",
    ]


def _is_scheduled_day(
    now: datetime,
    frequency: str,
    day_of_week: str,
    day_of_month: str,
) -> bool:
    """Return True if *today* is a valid send day for the configured schedule."""
    if frequency == 'daily':
        return True

    if frequency == 'weekly':
        return now.strftime('%A').lower() == day_of_week

    if frequency == 'monthly':
        if day_of_month == 'first':
            return now.day == 1
        if day_of_month == 'last':
            tomorrow = now + timedelta(days=1)
            return tomorrow.day == 1

    return False


def _already_sent_this_period(
    now: datetime,
    last_sent: datetime,
    frequency: str,
) -> bool:
    """Return True if an email was already sent during the current calendar period."""
    if frequency == 'daily':
        # Same UTC calendar day → already sent
        return last_sent.date() >= now.date()

    if frequency == 'weekly':
        # Sent fewer than ~7 days ago on or after the most recent scheduled day
        return (now - last_sent) < timedelta(days=6, hours=23)

    if frequency == 'monthly':
        # Same calendar month (UTC) → already sent
        return (last_sent.year, last_sent.month) >= (now.year, now.month)

    return False


def should_send_email_now() -> bool:
    """
    Check if an email should be sent based on the current schedule.

    The check is **calendar-based**: it first verifies that today matches
    the configured schedule day, then verifies that no email has already
    been sent for the current period.  The ``lastSent`` timestamp is
    persisted in ``email_config.json`` so the decision survives restarts
    and container rebuilds as long as the ``data/`` directory is retained.

    Returns:
        True if email should be sent, False otherwise
    """
    email_config = load_email_config()

    if not email_config.get('enabled'):
        return False

    frequency = email_config.get('frequency', 'weekly')
    day_of_week = email_config.get('dayOfWeek', 'monday').lower()
    day_of_month = email_config.get('dayOfMonth', 'first')

    now = datetime.now(timezone.utc)

    # 1. Is today a valid send day for this schedule?
    if not _is_scheduled_day(now, frequency, day_of_week, day_of_month):
        return False

    # 2. Have we already sent during this period?
    last_sent_str = email_config.get('lastSent')
    if not last_sent_str:
        logger.info("Email never sent before and today matches schedule — sending now")
        return True

    try:
        last_sent = datetime.fromisoformat(last_sent_str.replace('Z', '+00:00'))
    except (ValueError, AttributeError):
        logger.warning("Invalid lastSent timestamp: %s; treating as never sent", last_sent_str)
        return True

    if _already_sent_this_period(now, last_sent, frequency):
        return False

    logger.info("Schedule period elapsed since last send (%s) — sending now", last_sent_str)
    return True


def mark_email_sent() -> None:
    """Update the lastSent timestamp in email configuration."""
    email_config = load_email_config()
    email_config['lastSent'] = datetime.now(timezone.utc).isoformat()
    save_email_config(email_config)
    logger.info("Marked email as sent at %s", email_config['lastSent'])


__all__ = [
    "DEFAULT_EMAIL_CONFIG",
    "EMAIL_CONFIG_FILE",
    "get_smtp_config",
    "is_smtp_configured",
    "load_email_config",
    "save_email_config",
    "get_report_types",
    "should_send_email_now",
    "mark_email_sent",
    "_is_scheduled_day",
    "_already_sent_this_period",
]
