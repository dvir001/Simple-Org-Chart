"""Configuration and path helpers for SimpleOrgChart."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Dict

logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
CONFIG_DIR = BASE_DIR / "config"
STATIC_DIR = BASE_DIR / "static"
TEMPLATE_DIR = BASE_DIR / "templates"

REPO_DIR = BASE_DIR / "repositories"
SETTINGS_FILE = CONFIG_DIR / "app_settings.json"
DATA_FILE = DATA_DIR / "employee_data.json"
MISSING_MANAGER_FILE = DATA_DIR / "missing_manager_records.json"
EMPLOYEE_LIST_FILE = DATA_DIR / "employee_list.json"
DISABLED_LICENSE_FILE = DATA_DIR / "disabled_with_license_records.json"
FILTERED_LICENSE_FILE = DATA_DIR / "filtered_with_license_records.json"
FILTERED_USERS_FILE = DATA_DIR / "filtered_user_records.json"
DISABLED_USERS_FILE = DATA_DIR / "disabled_user_records.json"
LAST_LOGIN_FILE = DATA_DIR / "last_login_records.json"
RECENTLY_DISABLED_FILE = DATA_DIR / "recently_disabled_employees.json"
RECENTLY_HIRED_FILE = DATA_DIR / "recently_hired_employees.json"


def ensure_directories() -> None:
    """Ensure that the application's data, config, static, and repo directories exist."""
    for target in (DATA_DIR, CONFIG_DIR, STATIC_DIR, REPO_DIR):
        try:
            target.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            logger.warning("Failed to create directory %s: %s", target, error)


def as_posix_env(mapping: Dict[str, Path]) -> Dict[str, str]:
    """Helper to convert path constants into environment-style strings."""
    return {key: str(value) for key, value in mapping.items()}


__all__ = [
    "BASE_DIR",
    "DATA_DIR",
    "CONFIG_DIR",
    "STATIC_DIR",
    "TEMPLATE_DIR",
    "REPO_DIR",
    "SETTINGS_FILE",
    "DATA_FILE",
    "MISSING_MANAGER_FILE",
    "EMPLOYEE_LIST_FILE",
    "DISABLED_LICENSE_FILE",
    "FILTERED_LICENSE_FILE",
    "FILTERED_USERS_FILE",
    "DISABLED_USERS_FILE",
    "LAST_LOGIN_FILE",
    "RECENTLY_DISABLED_FILE",
    "RECENTLY_HIRED_FILE",
    "ensure_directories",
    "as_posix_env",
]
