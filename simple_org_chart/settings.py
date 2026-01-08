"""Settings management utilities for SimpleOrgChart."""

from __future__ import annotations

import json
import logging
import os
import re
from typing import Any, Dict, Iterable, Set

from .config import SETTINGS_FILE

logger = logging.getLogger(__name__)

TOP_LEVEL_USER_EMAIL = os.environ.get("TOP_LEVEL_USER_EMAIL", "")
TOP_LEVEL_USER_ID = os.environ.get("TOP_LEVEL_USER_ID")

DEFAULT_SETTINGS: Dict[str, Any] = {
    "chartTitle": "DB Auto Org Chart",
    "headerColor": "#0078D4",
    "logoPath": "/static/icon.png",
    "faviconPath": "/favicon.ico",
    "nodeColors": {
        "level0": "#90EE90",
        "level1": "#FFFFE0",
        "level2": "#E0F2FF",
        "level3": "#FFE4E1",
        "level4": "#E8DFF5",
        "level5": "#FFEAA7",
        "level6": "#FAD7FF",
        "level7": "#D7F8FF",
    },
    "autoUpdateEnabled": True,
    "updateTime": "20:00",
    "updateTimezone": "UTC",
    "collapseLevel": "2",
    "searchAutoExpand": True,
    "searchHighlight": True,
    "showNames": True,
    "showDepartments": True,
    "showJobTitles": True,
    "showOffice": False,
    "showEmployeeCount": True,
    "showProfileImages": True,
    "printOrientation": "landscape",
    "printSize": "a4",
    "exportXlsxColumns": {
        "name": "show",
        "title": "show",
        "department": "show",
        "email": "show",
        "phone": "show",
        "businessPhone": "show",
        "hireDate": "admin",
        "country": "show",
        "state": "show",
        "city": "show",
        "office": "show",
        "manager": "show",
    },
    "topUserEmail": TOP_LEVEL_USER_EMAIL.strip(),
    "newEmployeeMonths": 3,
    "multiLineChildrenEnabled": True,
    "multiLineChildrenThreshold": 20,
    "compactSiblingSpacingEnabled": False,
    "hideDisabledUsers": True,
    "hideGuestUsers": True,
    "hideNoTitle": True,
    "hideConsultantGroup": True,
    "ignoredEmployees": "",
    "ignoredDepartments": "",
    "ignoredTitles": "",
    "customDirectoryContacts": "",
}

_filter_legacy_split_re = re.compile(r"\s*[;,]+\s*")
_trim_edge_punct = re.compile(r"^[\s\-–—|]+|[\s\-–—|]+$")
_hex_six_pattern = re.compile(r"^[0-9a-fA-F]{6}$")


def _normalize_hex_color(value: Any, fallback: str) -> str:
    def _coerce(candidate: Any) -> str | None:
        if not isinstance(candidate, str):
            return None
        text = candidate.strip()
        if not text:
            return None
        if text.startswith("#"):
            text = text[1:]
        if _hex_six_pattern.fullmatch(text):
            return f"#{text.upper()}"
        return None

    normalized_value = _coerce(value)
    if normalized_value:
        return normalized_value

    normalized_fallback = _coerce(fallback)
    if normalized_fallback:
        return normalized_fallback

    return "#000000"


def translate_placeholder(key: str, default: str | None = None, **kwargs: Any) -> str:
    """Basic translation helper used by templates until full i18n is wired."""
    if default is not None:
        try:
            return default.format(**kwargs)
        except Exception:
            return default
    return key


def _apply_environment_overrides(settings: Dict[str, Any]) -> Dict[str, Any]:
    updated = settings.copy()
    updated["topUserEmail"] = TOP_LEVEL_USER_EMAIL.strip() if TOP_LEVEL_USER_EMAIL else updated.get("topUserEmail", "")
    return updated


def load_settings() -> Dict[str, Any]:
    """Load persisted settings or fall back to defaults."""
    if SETTINGS_FILE.exists():
        try:
            with SETTINGS_FILE.open("r", encoding="utf-8") as handle:
                stored = json.load(handle)
        except Exception as error:  # noqa: BLE001 - log and fall back
            logger.error("Error loading settings: %s", error)
        else:
            merged = DEFAULT_SETTINGS.copy()
            merged.update(stored)
            merged.pop("highlightNewEmployees", None)
            merged["headerColor"] = _normalize_hex_color(
                merged.get("headerColor"),
                DEFAULT_SETTINGS["headerColor"],
            )

            default_node_colors = DEFAULT_SETTINGS["nodeColors"].copy()
            stored_node_colors = stored.get("nodeColors")
            if isinstance(stored_node_colors, dict):
                default_node_colors.update(stored_node_colors)

            merged["nodeColors"] = {
                level: _normalize_hex_color(color, DEFAULT_SETTINGS["nodeColors"].get(level, "#000000"))
                for level, color in default_node_colors.items()
            }
            return _apply_environment_overrides(merged)

    return _apply_environment_overrides(DEFAULT_SETTINGS)


def save_settings(settings: Dict[str, Any]) -> bool:
    """Persist settings to disk, returning True on success."""
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)

    # Update stored defaults with provided overrides
    persisted = DEFAULT_SETTINGS.copy()
    persisted.update(settings)
    persisted.pop("highlightNewEmployees", None)

    persisted["headerColor"] = _normalize_hex_color(
        persisted.get("headerColor"),
        DEFAULT_SETTINGS["headerColor"],
    )

    default_node_colors = DEFAULT_SETTINGS["nodeColors"].copy()
    provided_node_colors = settings.get("nodeColors")
    if isinstance(provided_node_colors, dict):
        default_node_colors.update(provided_node_colors)

    persisted["nodeColors"] = {
        level: _normalize_hex_color(color, DEFAULT_SETTINGS["nodeColors"].get(level, "#000000"))
        for level, color in default_node_colors.items()
    }

    logger.info("Attempting to save settings to: %s", SETTINGS_FILE)
    try:
        with SETTINGS_FILE.open("w", encoding="utf-8") as handle:
            json.dump(persisted, handle, indent=2)
    except Exception as error:  # noqa: BLE001 - mirror legacy behaviour
        logger.error("Error saving settings to %s: %s", SETTINGS_FILE, error)
        return False

    logger.info("Settings saved successfully. File exists: %s", SETTINGS_FILE.exists())
    return True


def normalize_filter_value(value: Any) -> str:
    if not value:
        return ""
    cleaned = _trim_edge_punct.sub("", str(value))
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip().lower()


def parse_filter_values(raw_value: Any) -> Set[str]:
    if raw_value is None:
        return set()

    values: Iterable[Any] | None
    if isinstance(raw_value, str):
        text = raw_value.strip()
        if not text:
            return set()
        if text.startswith("["):
            try:
                decoded = json.loads(text)
                if isinstance(decoded, (list, tuple, set)):
                    values = list(decoded)
                else:
                    values = None
            except json.JSONDecodeError:
                values = None
        else:
            values = _filter_legacy_split_re.split(text)
    elif isinstance(raw_value, (list, tuple, set)):
        values = list(raw_value)
    else:
        values = None

    if not values:
        return set()

    normalized: Set[str] = set()
    for part in values:
        normalized_value = normalize_filter_value(part)
        if normalized_value:
            normalized.add(normalized_value)
    return normalized


def parse_ignored_departments(settings: Dict[str, Any]) -> Set[str]:
    return parse_filter_values(settings.get("ignoredDepartments", ""))


def parse_ignored_titles(settings: Dict[str, Any]) -> Set[str]:
    return parse_filter_values(settings.get("ignoredTitles", ""))


def parse_ignored_employees(settings: Dict[str, Any]) -> Set[str]:
    return parse_filter_values(settings.get("ignoredEmployees", ""))


def department_is_ignored(department: str, ignored_set: Set[str]) -> bool:
    if not ignored_set:
        return False
    normalized = normalize_filter_value(department)
    return normalized in ignored_set


def employee_is_ignored(name: str | None, email: str | None, user_principal_name: str | None, ignored_values: Set[str]) -> bool:
    if not ignored_values:
        return False

    candidates: Set[str] = set()

    for value in (name, email, user_principal_name):
        normalized = normalize_filter_value(value)
        if normalized:
            candidates.add(normalized)

    contact_values = [value for value in (email, user_principal_name) if value]
    for contact in contact_values:
        if name:
            combos = (
                f"{name} <{contact}>",
                f"{name} ({contact})",
                f"{name} - {contact}",
                f"{contact} ({name})",
                f"{contact} - {name}",
            )
            for combo in combos:
                normalized = normalize_filter_value(combo)
                if normalized:
                    candidates.add(normalized)

    return any(candidate in ignored_values for candidate in candidates)


__all__ = [
    "DEFAULT_SETTINGS",
    "TOP_LEVEL_USER_EMAIL",
    "TOP_LEVEL_USER_ID",
    "department_is_ignored",
    "employee_is_ignored",
    "load_settings",
    "normalize_filter_value",
    "parse_filter_values",
    "parse_ignored_departments",
    "parse_ignored_employees",
    "parse_ignored_titles",
    "save_settings",
    "translate_placeholder",
]
