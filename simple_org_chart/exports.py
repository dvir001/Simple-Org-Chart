"""Export utilities for SimpleOrgChart."""

from __future__ import annotations

import logging
from datetime import datetime
from typing import Optional

logger = logging.getLogger(__name__)


def format_hire_date(date_string: Optional[str]) -> str:
    """Format an ISO date string to a readable format."""
    if not date_string:
        return ''
    try:
        dt = datetime.fromisoformat(date_string.replace('Z', '+00:00'))
        return dt.strftime('%Y-%m-%d')
    except Exception:
        return date_string


__all__ = [
    'format_hire_date',
]
