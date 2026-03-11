"""Settings management utilities for SimpleOrgChart."""

from __future__ import annotations

import contextlib
import json
import logging
import os
import re
import tempfile
import threading
from typing import Any, Callable, Dict, Iterable, Set, Union

from .config import SETTINGS_FILE


class _InterProcessSettingsFileLock:
    """Combined thread + process lock for protecting SETTINGS_FILE.

    Serialises concurrent access across both threads within a single worker
    and across multiple Gunicorn worker processes by combining a
    ``threading.Lock`` with an ``fcntl.flock`` advisory file lock.  On
    platforms where ``fcntl`` is unavailable (e.g. Windows) the class
    falls back gracefully to thread-only locking.
    """

    def __init__(
        self,
        lock_file_path: Union[str, bytes, "os.PathLike[str]", Callable[[], Union[str, bytes, "os.PathLike[str]"]]],
    ) -> None:
        # Accept either a static path (str/bytes/Path) or a zero-argument
        # callable that returns the path at acquire-time.  The callable form
        # lets callers that re-bind the module-level SETTINGS_FILE (e.g. in
        # tests via monkeypatch) have the lock always protect the file that is
        # actually being accessed rather than the path that was current at
        # import time.
        if callable(lock_file_path):
            self._get_lock_file_path = lock_file_path
        else:
            _static = lock_file_path
            self._get_lock_file_path = lambda: _static
        self._thread_lock = threading.Lock()
        self._fd: int | None = None

    def acquire(self, blocking: bool = True) -> bool:
        acquired = self._thread_lock.acquire(blocking)
        if not acquired:
            return False
        fd = None
        try:
            import fcntl
            lock_file_path = self._get_lock_file_path()
            fd = os.open(lock_file_path, os.O_CREAT | os.O_RDWR, 0o600)
            flags = fcntl.LOCK_EX | (0 if blocking else fcntl.LOCK_NB)
            fcntl.flock(fd, flags)
            self._fd = fd
        except ImportError:
            # fcntl not available (Windows); thread lock is sufficient.
            if fd is not None:
                os.close(fd)
        except BlockingIOError:
            if fd is not None:
                os.close(fd)
            self._thread_lock.release()
            return False
        except Exception:
            if fd is not None:
                os.close(fd)
            self._thread_lock.release()
            raise
        return True

    def release(self) -> None:
        try:
            if self._fd is not None:
                try:
                    import fcntl
                    fcntl.flock(self._fd, fcntl.LOCK_UN)
                except (ImportError, OSError):
                    pass
                finally:
                    try:
                        os.close(self._fd)
                    except OSError:
                        pass
                    self._fd = None
        finally:
            self._thread_lock.release()

    def __enter__(self) -> "_InterProcessSettingsFileLock":
        self.acquire()
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.release()


# Shared inter-process lock protecting all reads/writes to SETTINGS_FILE.
# Both settings.py and email_config.py import this lock so that concurrent
# writes from different code paths (scheduler, HTTP handlers, workers) are
# serialised across all Gunicorn worker processes.
#
# The factory callable reads SETTINGS_FILE from this module's namespace at
# acquire-time rather than computing the path once at import time.  This
# allows test-time monkeypatching of ``simple_org_chart.settings.SETTINGS_FILE``
# to take effect so the lock file is always co-located with the settings file
# actually being accessed.
def _settings_lock_file_factory() -> str:
    # Accessing the module-level name SETTINGS_FILE here (rather than via
    # globals()) is sufficient: Python resolves global names at call-time via
    # LOAD_GLOBAL, so monkeypatching ``simple_org_chart.settings.SETTINGS_FILE``
    # in tests will be reflected when this factory is invoked.
    return os.path.join(os.fspath(SETTINGS_FILE.parent), f"{SETTINGS_FILE.name}.lock")


_settings_file_lock = _InterProcessSettingsFileLock(_settings_lock_file_factory)

logger = logging.getLogger(__name__)

DEFAULT_SETTINGS: Dict[str, Any] = {
    "chartTitle": "Simple Org Chart",
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
    "collapseLevel": "2",
    "searchAutoExpand": True,
    "searchHighlight": True,
    "searchHighlightDuration": 5,
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
    "topLevelUserEmail": "",
    "topLevelUserId": "",
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
    "userScannerEnabled": False,
    "teamsPresenceEnabled": False,
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
    """Apply environment variable overrides (currently a pass-through)."""
    return settings.copy()


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
        with _settings_file_lock:
            # Read the current file so we keep all other keys intact
            existing: Dict[str, Any] = {}
            if SETTINGS_FILE.exists():
                try:
                    with SETTINGS_FILE.open("r", encoding="utf-8") as handle:
                        loaded = json.load(handle)
                        if isinstance(loaded, dict):
                            existing = loaded
                except Exception:
                    pass

            existing.update(persisted)

            # Atomic write: write to a temp file then replace
            tmp_fd, tmp_path = tempfile.mkstemp(
                dir=SETTINGS_FILE.parent, suffix=".tmp"
            )
            try:
                with os.fdopen(tmp_fd, "w", encoding="utf-8") as tmp_handle:
                    json.dump(existing, tmp_handle, indent=2)
                os.replace(tmp_path, SETTINGS_FILE)
            except Exception:
                with contextlib.suppress(OSError):
                    os.unlink(tmp_path)
                raise
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
    "_settings_file_lock",
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
