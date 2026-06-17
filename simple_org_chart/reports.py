"""Report cache helpers for SimpleOrgChart."""

from __future__ import annotations

import json
import logging
import os
from datetime import datetime, timedelta, timezone
from typing import Callable, Iterable, List, Optional, Sequence, Type

import simple_org_chart.config as app_config
from simple_org_chart.msgraph import parse_graph_datetime

logger = logging.getLogger(__name__)


MISSING_MANAGER_FILE = str(app_config.MISSING_MANAGER_FILE)
DISABLED_LICENSE_FILE = str(app_config.DISABLED_LICENSE_FILE)
DISABLED_USERS_FILE = str(app_config.DISABLED_USERS_FILE)
RECENTLY_DISABLED_FILE = str(app_config.RECENTLY_DISABLED_FILE)
RECENTLY_HIRED_FILE = str(app_config.RECENTLY_HIRED_FILE)
LAST_LOGIN_FILE = str(app_config.LAST_LOGIN_FILE)
FILTERED_LICENSE_FILE = str(app_config.FILTERED_LICENSE_FILE)
FILTERED_USERS_FILE = str(app_config.FILTERED_USERS_FILE)
MISSING_PHOTO_FILE = str(app_config.MISSING_PHOTO_FILE)
DIRTY_DATA_FILE = str(app_config.DIRTY_DATA_FILE)
MISSING_HIRE_DATE_FILE = str(app_config.MISSING_HIRE_DATE_FILE)


class ReportCacheManager:
    """Centralised helper for loading cached report data."""

    def __init__(self, refresh_callback: Optional[Callable[[], None]] = None) -> None:
        self._refresh_callback = refresh_callback

    def load_json(
        self,
        path: str,
        *,
        refresh: bool = False,
        description: str = "report cache",
        expected_type: Optional[Type] = list,
    ):
        """Load a JSON payload from disk, optionally refreshing first."""
        if not path:
            logger.error("No path provided for %s", description)
            return [] if expected_type is list else None

        if refresh or not os.path.exists(path):
            if refresh:
                logger.info("Refreshing %s", description)
            if self._refresh_callback is not None:
                try:
                    self._refresh_callback()
                except Exception as exc:  # pragma: no cover - defensive
                    logger.error("Failed to refresh %s: %s", description, exc)
            else:
                logger.warning("No refresh callback configured; cannot refresh %s", description)

        if not os.path.exists(path):
            logger.warning("%s not found at %s", description, path)
            return [] if expected_type is list else None

        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except json.JSONDecodeError as decode_error:
            logger.error("Failed to parse %s at %s: %s", description, path, decode_error)
            return [] if expected_type is list else None
        except Exception as error:  # pragma: no cover - I/O errors
            logger.error("Unexpected error loading %s at %s: %s", description, path, error)
            return [] if expected_type is list else None

        if expected_type is not None and not isinstance(data, expected_type):
            logger.warning("Unexpected payload type for %s; expected %s", description, expected_type.__name__)
            return [] if expected_type is list else None

        return data


def load_missing_manager_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        MISSING_MANAGER_FILE,
        refresh=force_refresh,
        description="missing manager report cache",
    )


def load_missing_photo_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        MISSING_PHOTO_FILE,
        refresh=force_refresh,
        description="missing photo report cache",
    )


def load_missing_hire_date_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        MISSING_HIRE_DATE_FILE,
        refresh=force_refresh,
        description="missing hire date report cache",
    )


def load_disabled_license_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        DISABLED_LICENSE_FILE,
        refresh=force_refresh,
        description="disabled licensed users report cache",
    )


def load_disabled_users_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        DISABLED_USERS_FILE,
        refresh=force_refresh,
        description="disabled users report cache",
    )


def load_recently_disabled_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        RECENTLY_DISABLED_FILE,
        refresh=force_refresh,
        description="recently disabled employees report cache",
    )


def load_recently_hired_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        RECENTLY_HIRED_FILE,
        refresh=force_refresh,
        description="recently hired employees report cache",
    )


def load_last_login_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        LAST_LOGIN_FILE,
        refresh=force_refresh,
        description="last sign-in report cache",
    )


def load_filtered_license_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        FILTERED_LICENSE_FILE,
        refresh=force_refresh,
        description="filtered licensed users report cache",
    )


def load_filtered_user_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        FILTERED_USERS_FILE,
        refresh=force_refresh,
        description="filtered users report cache",
    )


def apply_disabled_filters(
    records: Optional[Sequence[dict]],
    *,
    licensed_only: bool = False,
    recent_days: Optional[int] = None,
    include_guests: bool = False,
    include_members: bool = True,
):
    if not records:
        return []

    cutoff = None
    if recent_days and recent_days > 0:
        cutoff = datetime.now(timezone.utc) - timedelta(days=recent_days)

    filtered: List[dict] = []

    for record in records:
        user_type = (record.get("userType") or "").lower()

        if user_type == "guest" and not include_guests:
            continue
        if user_type == "member" and not include_members:
            continue

        if licensed_only and (record.get("licenseCount") or 0) == 0:
            continue

        if cutoff is not None:
            observed = parse_graph_datetime(
                record.get("firstSeenDisabledAt")
                or record.get("disabledDate")
            )
            if not observed or observed < cutoff:
                continue

        filtered.append(record)

    return filtered


def calculate_license_totals(records: Optional[Iterable[dict]]):
    return sum((record.get("licenseCount") or 0) for record in records or [])


def _resolve_mailbox_categories(record: dict) -> tuple[bool, bool, bool]:
    mailbox_type_raw = record.get("mailboxType")
    mailbox_type_value = str(mailbox_type_raw).strip().lower() if mailbox_type_raw is not None else ""

    shared_flag = record.get("isSharedMailbox")
    if isinstance(shared_flag, str):
        lowered_flag = shared_flag.strip().lower()
        if lowered_flag in {"true", "1", "yes", "on"}:
            shared_flag = True
        elif lowered_flag in {"false", "0", "no", "off"}:
            shared_flag = False

    is_shared_mailbox = False
    if isinstance(shared_flag, bool):
        is_shared_mailbox = shared_flag
    elif shared_flag:
        is_shared_mailbox = True

    if not is_shared_mailbox and mailbox_type_value.startswith("shared"):
        is_shared_mailbox = True

    is_room_equipment_mailbox = False
    if not is_shared_mailbox and mailbox_type_value:
        if mailbox_type_value in {"room", "equipment"}:
            is_room_equipment_mailbox = True
        elif any(mailbox_type_value.startswith(prefix) for prefix in ("room", "equipment")):
            is_room_equipment_mailbox = True

    is_user_mailbox = not is_shared_mailbox and not is_room_equipment_mailbox

    return is_user_mailbox, is_shared_mailbox, is_room_equipment_mailbox


def apply_last_login_filters(
    records: Optional[Sequence[dict]],
    *,
    include_user_mailboxes: bool = True,
    include_shared_mailboxes: bool = True,
    include_room_equipment_mailboxes: bool = True,
    include_enabled: bool = True,
    include_disabled: bool = True,
    include_licensed: bool = True,
    include_unlicensed: bool = True,
    include_members: bool = True,
    include_guests: bool = True,
    include_never_signed_in: bool = True,
    inactive_days: Optional[str] = None,
    inactive_days_max: Optional[str] = None,
    include_hidden_from_address_list: bool = True,
    include_visible_in_address_list: bool = True,
    include_with_mailbox: bool = True,
    include_without_mailbox: bool = True,
    include_with_manager: bool = True,
    include_without_manager: bool = True,
):
    if not records:
        return []

    inactive_threshold = None
    require_never_signed_in = False
    inactive_max_threshold = None

    if inactive_days not in (None, "", "none"):
        if isinstance(inactive_days, str) and inactive_days.lower() == "never":
            require_never_signed_in = True
        else:
            try:
                inactive_threshold = int(inactive_days)  # type: ignore[arg-type]
            except (TypeError, ValueError):
                inactive_threshold = None

    if inactive_days_max not in (None, "", "none"):
        try:
            inactive_max_threshold = int(inactive_days_max)  # type: ignore[arg-type]
        except (TypeError, ValueError):
            inactive_max_threshold = None

    filtered: List[dict] = []

    for record in records:
        is_user_mailbox, is_shared_mailbox, is_room_equipment_mailbox = _resolve_mailbox_categories(record)

        if is_user_mailbox and not include_user_mailboxes:
            continue
        if is_shared_mailbox and not include_shared_mailboxes:
            continue
        if is_room_equipment_mailbox and not include_room_equipment_mailboxes:
            continue

        account_enabled = record.get("accountEnabled", True)
        if account_enabled and not include_enabled:
            continue
        if not account_enabled and not include_disabled:
            continue

        license_count = record.get("licenseCount") or 0
        if license_count > 0 and not include_licensed:
            continue
        if license_count == 0 and not include_unlicensed:
            continue

        if is_user_mailbox:
            user_type = (record.get("userType") or "").lower()
            if user_type == "member" and not include_members:
                continue
            if user_type == "guest" and not include_guests:
                continue

        never_signed_in = bool(record.get("neverSignedIn"))
        if never_signed_in and not include_never_signed_in:
            continue
        if require_never_signed_in and not never_signed_in:
            continue

        if inactive_threshold is not None:
            days_since = record.get("daysSinceLastActivity")
            if days_since is None or days_since < inactive_threshold:
                continue

        if inactive_max_threshold is not None:
            days_since = record.get("daysSinceLastActivity")
            if days_since is None or days_since > inactive_max_threshold:
                continue

        # A shared/room/equipment mailbox is a mailbox by definition.
        has_mailbox = bool(record.get("hasMailbox", True)) or is_shared_mailbox or is_room_equipment_mailbox

        # GAL visibility is a mailbox concept; the two populations are filtered independently.
        # Mailbox users: apply the hidden/visible toggles directly.
        # Non-mailbox users: they have no GAL status; exclude them when either GAL toggle is
        # active (narrowing to a specific group) so they don't bleed into GAL-specific results.
        if has_mailbox:
            hidden = bool(record.get("hiddenFromAddressLists"))
            if hidden and not include_hidden_from_address_list:
                continue
            if not hidden and not include_visible_in_address_list:
                continue
        elif not include_hidden_from_address_list or not include_visible_in_address_list:
            continue

        if has_mailbox and not include_with_mailbox:
            continue
        if not has_mailbox and not include_without_mailbox:
            continue

        has_manager = bool(record.get("hasManager", True))
        if has_manager and not include_with_manager:
            continue
        if not has_manager and not include_without_manager:
            continue

        filtered.append(record)

    return filtered


def apply_filtered_user_filters(
    records: Optional[Sequence[dict]],
    *,
    include_user_mailboxes: bool = True,
    include_shared_mailboxes: bool = True,
    include_room_equipment_mailboxes: bool = True,
    include_enabled: bool = True,
    include_disabled: bool = True,
    include_licensed: bool = True,
    include_unlicensed: bool = True,
    include_members: bool = True,
    include_guests: bool = True,
    include_hidden_from_address_list: bool = True,
    include_visible_in_address_list: bool = True,
    include_with_mailbox: bool = True,
    include_without_mailbox: bool = True,
    include_with_manager: bool = True,
    include_without_manager: bool = True,
):
    if not records:
        return []

    filtered: List[dict] = []

    for record in records:
        is_user_mailbox, is_shared_mailbox, is_room_equipment_mailbox = _resolve_mailbox_categories(record)

        if is_user_mailbox and not include_user_mailboxes:
            continue
        if is_shared_mailbox and not include_shared_mailboxes:
            continue
        if is_room_equipment_mailbox and not include_room_equipment_mailboxes:
            continue

        account_enabled = record.get("accountEnabled", True)
        if account_enabled and not include_enabled:
            continue
        if not account_enabled and not include_disabled:
            continue

        license_count = record.get("licenseCount") or 0
        if license_count > 0 and not include_licensed:
            continue
        if license_count == 0 and not include_unlicensed:
            continue

        user_type = (record.get("userType") or "").lower()

        if user_type == "guest" and not include_guests:
            continue
        if user_type == "member" and not include_members:
            continue

        # A shared/room/equipment mailbox is a mailbox by definition.
        has_mailbox = bool(record.get("hasMailbox", True)) or is_shared_mailbox or is_room_equipment_mailbox

        # GAL visibility is a mailbox concept; the two populations are filtered independently.
        # Mailbox users: apply the hidden/visible toggles directly.
        # Non-mailbox users: they have no GAL status; exclude them when either GAL toggle is
        # active (narrowing to a specific group) so they don't bleed into GAL-specific results.
        if has_mailbox:
            hidden = bool(record.get("hiddenFromAddressLists"))
            if hidden and not include_hidden_from_address_list:
                continue
            if not hidden and not include_visible_in_address_list:
                continue
        elif not include_hidden_from_address_list or not include_visible_in_address_list:
            continue

        if has_mailbox and not include_with_mailbox:
            continue
        if not has_mailbox and not include_without_mailbox:
            continue

        has_manager = bool(record.get("hasManager", True))
        if has_manager and not include_with_manager:
            continue
        if not has_manager and not include_without_manager:
            continue

        filtered.append(record)

    return filtered


def apply_missing_manager_filters(
    records: Optional[Sequence[dict]],
    *,
    include_user_mailboxes: bool = True,
    include_shared_mailboxes: bool = True,
    include_room_equipment_mailboxes: bool = True,
    include_enabled: bool = True,
    include_disabled: bool = True,
    include_licensed: bool = True,
    include_unlicensed: bool = True,
    include_members: bool = True,
    include_guests: bool = True,
    include_hidden_from_address_list: bool = True,
    include_visible_in_address_list: bool = True,
    include_with_mailbox: bool = True,
    include_without_mailbox: bool = True,
    include_with_manager: bool = True,
    include_without_manager: bool = True,
):
    return apply_filtered_user_filters(
        records,
        include_user_mailboxes=include_user_mailboxes,
        include_shared_mailboxes=include_shared_mailboxes,
        include_room_equipment_mailboxes=include_room_equipment_mailboxes,
        include_enabled=include_enabled,
        include_disabled=include_disabled,
        include_licensed=include_licensed,
        include_unlicensed=include_unlicensed,
        include_members=include_members,
        include_guests=include_guests,
        include_hidden_from_address_list=include_hidden_from_address_list,
        include_visible_in_address_list=include_visible_in_address_list,
        include_with_mailbox=include_with_mailbox,
        include_without_mailbox=include_without_mailbox,
        include_with_manager=include_with_manager,
        include_without_manager=include_without_manager,
    )


def apply_missing_hire_date_filters(
    records: Optional[Sequence[dict]],
    *,
    include_user_mailboxes: bool = True,
    include_shared_mailboxes: bool = True,
    include_room_equipment_mailboxes: bool = True,
    include_enabled: bool = True,
    include_disabled: bool = True,
    include_licensed: bool = True,
    include_unlicensed: bool = True,
    include_members: bool = True,
    include_guests: bool = True,
    include_hidden_from_address_list: bool = True,
    include_visible_in_address_list: bool = True,
    include_with_mailbox: bool = True,
    include_without_mailbox: bool = True,
    include_with_manager: bool = True,
    include_without_manager: bool = True,
):
    return apply_filtered_user_filters(
        records,
        include_user_mailboxes=include_user_mailboxes,
        include_shared_mailboxes=include_shared_mailboxes,
        include_room_equipment_mailboxes=include_room_equipment_mailboxes,
        include_enabled=include_enabled,
        include_disabled=include_disabled,
        include_licensed=include_licensed,
        include_unlicensed=include_unlicensed,
        include_members=include_members,
        include_guests=include_guests,
        include_hidden_from_address_list=include_hidden_from_address_list,
        include_visible_in_address_list=include_visible_in_address_list,
        include_with_mailbox=include_with_mailbox,
        include_without_mailbox=include_without_mailbox,
        include_with_manager=include_with_manager,
        include_without_manager=include_without_manager,
    )


def apply_missing_photo_filters(
    records: Optional[Sequence[dict]],
    *,
    include_user_mailboxes: bool = True,
    include_shared_mailboxes: bool = True,
    include_room_equipment_mailboxes: bool = True,
    include_enabled: bool = True,
    include_disabled: bool = True,
    include_licensed: bool = True,
    include_unlicensed: bool = True,
    include_members: bool = True,
    include_guests: bool = True,
    include_hidden_from_address_list: bool = True,
    include_visible_in_address_list: bool = True,
    include_with_mailbox: bool = True,
    include_without_mailbox: bool = True,
    include_with_manager: bool = True,
    include_without_manager: bool = True,
):
    return apply_filtered_user_filters(
        records,
        include_user_mailboxes=include_user_mailboxes,
        include_shared_mailboxes=include_shared_mailboxes,
        include_room_equipment_mailboxes=include_room_equipment_mailboxes,
        include_enabled=include_enabled,
        include_disabled=include_disabled,
        include_licensed=include_licensed,
        include_unlicensed=include_unlicensed,
        include_members=include_members,
        include_guests=include_guests,
        include_hidden_from_address_list=include_hidden_from_address_list,
        include_visible_in_address_list=include_visible_in_address_list,
        include_with_mailbox=include_with_mailbox,
        include_without_mailbox=include_without_mailbox,
        include_with_manager=include_with_manager,
        include_without_manager=include_without_manager,
    )


def apply_tagpicker_filters(
    records: Sequence[dict],
    *,
    filter_titles: Optional[List[str]] = None,
    filter_titles_mode: str = "exclude",
    filter_departments: Optional[List[str]] = None,
    filter_departments_mode: str = "exclude",
    filter_countries: Optional[List[str]] = None,
    filter_countries_mode: str = "exclude",
    filter_states: Optional[List[str]] = None,
    filter_states_mode: str = "exclude",
) -> List[dict]:
    """Apply optional title/department/country/state include/exclude filters."""
    if not records:
        return []

    title_set = {v.strip().lower() for v in (filter_titles or []) if v and v.strip()}
    dept_set = {v.strip().lower() for v in (filter_departments or []) if v and v.strip()}
    country_set = {v.strip().lower() for v in (filter_countries or []) if v and v.strip()}
    state_set = {v.strip().lower() for v in (filter_states or []) if v and v.strip()}

    # Nothing to filter
    if not title_set and not dept_set and not country_set and not state_set:
        return list(records)

    filtered: List[dict] = []
    for record in records:
        title_val = (record.get("title") or "").strip().lower()
        dept_val = (record.get("department") or "").strip().lower()
        country_val = (record.get("country") or "").strip().lower()
        state_val = (record.get("state") or "").strip().lower()

        if title_set:
            matched = title_val in title_set
            if filter_titles_mode == "include" and not matched:
                continue
            if filter_titles_mode == "exclude" and matched:
                continue

        if dept_set:
            matched = dept_val in dept_set
            if filter_departments_mode == "include" and not matched:
                continue
            if filter_departments_mode == "exclude" and matched:
                continue

        if country_set:
            matched = country_val in country_set
            if filter_countries_mode == "include" and not matched:
                continue
            if filter_countries_mode == "exclude" and matched:
                continue

        if state_set:
            matched = state_val in state_set
            if filter_states_mode == "include" and not matched:
                continue
            if filter_states_mode == "exclude" and matched:
                continue

        filtered.append(record)

    return filtered



# ---------------------------------------------------------------------------
# Dirty data detection
# ---------------------------------------------------------------------------

_DIRTY_DATA_FIELDS = [
    ("name", "Name"),
    ("title", "Title"),
    ("department", "Department"),
    ("email", "Email"),
    ("phone", "Mobile Phone"),
    ("businessPhone", "Business Phone"),
    ("location", "Office Location"),
    ("city", "City"),
    ("state", "State / Province"),
    ("country", "Country"),
    ("usageLocation", "Usage Location"),
]


def _check_field_whitespace(value: object) -> Optional[str]:
    """Return a human-readable issue string if *value* has whitespace problems."""
    if not isinstance(value, str) or not value:
        return None
    if value != value.strip():
        return "leading/trailing spaces"
    if "  " in value:
        return "consecutive spaces"
    return None


def _check_email(value: object) -> list:
    """Return a list of issue strings for email-specific problems."""
    if not isinstance(value, str) or not value:
        return []
    domain = value.lower().rsplit("@", 1)[-1]
    if domain == "onmicrosoft.com" or domain.endswith(".onmicrosoft.com"):
        return []
    problems = []
    ws = _check_field_whitespace(value)
    if ws:
        problems.append(ws)
    if value != value.lower():
        problems.append("uppercase letters")
    return problems


def detect_dirty_data_records(employees: Iterable[dict]) -> List[dict]:
    """Scan employee records for fields with whitespace data-quality issues.

    Returns records that have at least one issue, each augmented with an
    ``issues`` list of ``{"field": <label>, "value": <raw>, "problem": <desc>}``
    entries.
    """
    results: List[dict] = []
    for emp in employees or []:
        issues = []
        for field_key, field_label in _DIRTY_DATA_FIELDS:
            raw = emp.get(field_key)
            if field_key == "email":
                problems = _check_email(raw)
            else:
                problem = _check_field_whitespace(raw)
                problems = [problem] if problem else []
            for problem in problems:
                issues.append({
                    "field": field_label,
                    "fieldKey": field_key,
                    "value": raw,
                    "problem": problem,
                })
        if issues:
            results.append({
                "id": emp.get("id"),
                "name": emp.get("name") or "",
                "title": emp.get("title") or "",
                "department": emp.get("department") or "",
                "email": emp.get("email") or "",
                "country": emp.get("country") or "",
                "accountEnabled": emp.get("accountEnabled", True),
                "userType": emp.get("userType") or "",
                "licenseCount": emp.get("licenseCount", 0),
                "licenseSkus": emp.get("licenseSkus", []),
                "licenseSkuIds": emp.get("licenseSkuIds", []),
                "mailboxType": emp.get("mailboxType"),
                "isSharedMailbox": emp.get("isSharedMailbox"),
                "issues": issues,
                "issueCount": len(issues),
                "issueFields": ", ".join(i["field"] for i in issues),
            })
    return results


def load_dirty_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        DIRTY_DATA_FILE,
        refresh=force_refresh,
        description="dirty data report cache",
    )


def apply_dirty_data_filters(
    records: Optional[Sequence[dict]],
    *,
    include_enabled: bool = True,
    include_disabled: bool = True,
    include_licensed: bool = True,
    include_unlicensed: bool = True,
    include_members: bool = True,
    include_guests: bool = True,
) -> List[dict]:
    """Simple filter for the dirty-data report (no mailbox/GAL/manager toggles needed)."""
    if not records:
        return []

    filtered: List[dict] = []
    for record in records:
        account_enabled = record.get("accountEnabled", True)
        if account_enabled and not include_enabled:
            continue
        if not account_enabled and not include_disabled:
            continue

        license_count = record.get("licenseCount") or 0
        if license_count > 0 and not include_licensed:
            continue
        if license_count == 0 and not include_unlicensed:
            continue

        user_type = (record.get("userType") or "").lower()
        if user_type == "member" and not include_members:
            continue
        if user_type == "guest" and not include_guests:
            continue

        filtered.append(record)

    return filtered


__all__ = [
    "ReportCacheManager",
    "apply_dirty_data_filters",
    "apply_disabled_filters",
    "apply_filtered_user_filters",
    "apply_last_login_filters",
    "apply_missing_hire_date_filters",
    "apply_missing_manager_filters",
    "apply_missing_photo_filters",
    "apply_tagpicker_filters",
    "calculate_license_totals",
    "detect_dirty_data_records",
    "load_dirty_data",
    "load_disabled_license_data",
    "load_disabled_users_data",
    "load_filtered_license_data",
    "load_filtered_user_data",
    "load_last_login_data",
    "load_missing_hire_date_data",
    "load_missing_manager_data",
    "load_missing_photo_data",
    "load_recently_disabled_data",
    "load_recently_hired_data",
]
