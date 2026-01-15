"""Data update status and refresh logic for SimpleOrgChart."""

from __future__ import annotations

import json
import logging
import os
import threading
from datetime import datetime, timedelta, timezone
from typing import Any, Callable, Dict, List, Optional

import simple_org_chart.config as app_config
from simple_org_chart.msgraph import (
    calculate_days_since,
    collect_disabled_users,
    collect_last_login_records,
    datetime_to_iso,
    fetch_all_employees,
    get_access_token,
    parse_graph_datetime,
    _enrich_mailbox_metadata,
)
from simple_org_chart.settings import (
    department_is_ignored,
    employee_is_ignored,
    load_settings,
    parse_ignored_departments,
    parse_ignored_employees,
)
from simple_org_chart.hierarchy import (
    build_org_hierarchy,
    collect_missing_manager_records,
)

logger = logging.getLogger(__name__)

DATA_DIR = str(app_config.DATA_DIR)
DATA_FILE = str(app_config.DATA_FILE)
EMPLOYEE_LIST_FILE = str(app_config.EMPLOYEE_LIST_FILE)
MISSING_MANAGER_FILE = str(app_config.MISSING_MANAGER_FILE)
DISABLED_LICENSE_FILE = str(app_config.DISABLED_LICENSE_FILE)
DISABLED_USERS_FILE = str(app_config.DISABLED_USERS_FILE)
RECENTLY_DISABLED_FILE = str(app_config.RECENTLY_DISABLED_FILE)
RECENTLY_HIRED_FILE = str(app_config.RECENTLY_HIRED_FILE)
LAST_LOGIN_FILE = str(app_config.LAST_LOGIN_FILE)
FILTERED_LICENSE_FILE = str(app_config.FILTERED_LICENSE_FILE)
FILTERED_USERS_FILE = str(app_config.FILTERED_USERS_FILE)
DATA_UPDATE_STATUS_FILE = os.path.join(DATA_DIR, 'data_update_status.json')

_DATA_UPDATE_STATUS_LOCK = threading.Lock()
_CURRENT_DATA_UPDATE_STATUS: Dict[str, Any] = {'state': 'idle'}
_APP_STARTUP_COMPLETE = False


def _write_data_update_status(payload: Dict[str, Any]) -> Dict[str, Any]:
    global _CURRENT_DATA_UPDATE_STATUS
    with _DATA_UPDATE_STATUS_LOCK:
        _CURRENT_DATA_UPDATE_STATUS = payload
        try:
            with open(DATA_UPDATE_STATUS_FILE, 'w') as status_file:
                json.dump(payload, status_file, indent=2)
        except Exception as error:
            logger.warning("Failed to write data update status: %s", error)
    return payload


def load_data_update_status() -> Dict[str, Any]:
    """Load data update status from disk, resetting stale running states."""
    global _CURRENT_DATA_UPDATE_STATUS, _APP_STARTUP_COMPLETE
    stale_override = None
    with _DATA_UPDATE_STATUS_LOCK:
        if os.path.exists(DATA_UPDATE_STATUS_FILE):
            try:
                with open(DATA_UPDATE_STATUS_FILE, 'r') as status_file:
                    data = json.load(status_file)
                if isinstance(data, dict):
                    _CURRENT_DATA_UPDATE_STATUS = data
            except Exception as error:
                logger.warning("Failed to load data update status: %s", error)

        state = (_CURRENT_DATA_UPDATE_STATUS or {}).get('state')
        if state == 'running':
            # On first load after app startup, any "running" status is stale
            # because the previous process is gone after a restart
            if not _APP_STARTUP_COMPLETE:
                stale_override = {
                    'state': 'idle',
                    'success': False,
                    'finishedAt': datetime.now(timezone.utc).isoformat(),
                    'error': 'Previous update was interrupted by application restart.',
                }
            else:
                # During normal operation, check elapsed time
                started_text = (_CURRENT_DATA_UPDATE_STATUS or {}).get('startedAt')
                try:
                    started_dt = datetime.fromisoformat(started_text) if started_text else None
                    if started_dt and started_dt.tzinfo is None:
                        started_dt = started_dt.replace(tzinfo=timezone.utc)
                except Exception:
                    started_dt = None

                if started_dt:
                    elapsed = datetime.now(timezone.utc) - started_dt.astimezone(timezone.utc)
                    if elapsed > timedelta(hours=2):
                        stale_override = {
                            'state': 'idle',
                            'success': False,
                            'finishedAt': datetime.now(timezone.utc).isoformat(),
                            'error': 'Previous update status appeared stuck; automatically reset.',
                        }
                else:
                    stale_override = {
                        'state': 'idle',
                        'success': False,
                        'finishedAt': datetime.now(timezone.utc).isoformat(),
                        'error': 'Previous update status was invalid; automatically reset.',
                    }

        if stale_override:
            last_success = (_CURRENT_DATA_UPDATE_STATUS or {}).get('lastSuccessAt')
            if last_success:
                stale_override['lastSuccessAt'] = last_success
            _CURRENT_DATA_UPDATE_STATUS = stale_override

        current_snapshot = dict(_CURRENT_DATA_UPDATE_STATUS)

    if stale_override:
        _write_data_update_status(stale_override)

    return current_snapshot


def mark_data_update_running(source: str = 'unknown') -> Dict[str, Any]:
    """Mark data update as running."""
    previous_status = load_data_update_status()
    status = {
        'state': 'running',
        'source': source,
        'startedAt': datetime.now(timezone.utc).isoformat(),
    }
    if isinstance(previous_status, dict) and previous_status.get('lastSuccessAt'):
        status['lastSuccessAt'] = previous_status['lastSuccessAt']
    return _write_data_update_status(status)


def mark_data_update_finished(
    success: bool = True,
    error: Optional[str] = None,
    source: str = 'unknown',
) -> Dict[str, Any]:
    """Mark data update as finished."""
    status: Dict[str, Any] = {
        'state': 'idle',
        'success': bool(success),
        'finishedAt': datetime.now(timezone.utc).isoformat(),
        'source': source,
    }
    if error:
        status['error'] = str(error)
    if success:
        status['lastSuccessAt'] = status['finishedAt']
    else:
        previous_status = load_data_update_status()
        last_success = previous_status.get('lastSuccessAt') if isinstance(previous_status, dict) else None
        if last_success:
            status['lastSuccessAt'] = last_success
    return _write_data_update_status(status)


def mark_startup_complete() -> None:
    """Mark that app startup is complete (for stale status detection)."""
    global _APP_STARTUP_COMPLETE
    _APP_STARTUP_COMPLETE = True


def collect_recently_disabled_employees(
    records: List[Dict[str, Any]],
    days: int = 365,
) -> List[Dict[str, Any]]:
    """Filter disabled user records to those disabled within the last N days."""
    if not records:
        return []

    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    recent = []

    for record in records:
        observed_value = (
            record.get('firstSeenDisabledAt')
            or record.get('disabledDate')
        )
        disabled_at = parse_graph_datetime(observed_value)
        if not disabled_at or disabled_at < cutoff:
            continue

        updated = record.copy()
        updated['disabledDate'] = datetime_to_iso(disabled_at)
        updated['disabledDays'] = calculate_days_since(disabled_at)
        if not updated.get('firstSeenDisabledAt'):
            updated['firstSeenDisabledAt'] = updated['disabledDate']
        recent.append(updated)

    recent.sort(key=lambda item: item.get('disabledDate') or '')
    return recent


def collect_recently_hired_employees(
    employees: List[Dict[str, Any]],
    days: int = 365,
) -> List[Dict[str, Any]]:
    """Filter employees to those hired within the last N days."""
    if not employees:
        return []

    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    manager_lookup = {emp.get('id'): emp for emp in employees if emp.get('id')}
    recent = []

    for employee in employees:
        hire_date = parse_graph_datetime(employee.get('hireDate') or employee.get('employeeHireDate'))
        if not hire_date or hire_date < cutoff:
            continue

        record = {
            'id': employee.get('id'),
            'name': employee.get('name'),
            'title': employee.get('title'),
            'department': employee.get('department'),
            'email': employee.get('email'),
            'userPrincipalName': employee.get('userPrincipalName'),
            'phone': employee.get('phone') or '',
            'businessPhone': employee.get('businessPhone') or '',
            'location': employee.get('location') or employee.get('officeLocation') or '',
            'hireDate': datetime_to_iso(hire_date),
            'daysSinceHire': calculate_days_since(hire_date),
            'managerName': '',
        }

        manager_id = employee.get('managerId')
        if manager_id and manager_id in manager_lookup:
            record['managerName'] = manager_lookup[manager_id].get('name') or ''

        recent.append(record)

    recent.sort(key=lambda item: item.get('hireDate') or '')
    return recent


def _load_fetch_all_employees_fallback() -> tuple:
    """Load cached employee data when Graph API fetch fails."""
    from simple_org_chart.hierarchy import load_cached_employees

    cached_employees = load_cached_employees() or []

    def _load_cached_list(path, description):
        if not os.path.exists(path):
            logger.debug(f"No cached {description} found at {path}")
            return []
        try:
            with open(path, 'r') as cache_file:
                data = json.load(cache_file)
        except Exception as error:
            logger.error(f"Failed to load cached {description} from {path}: {error}")
            return []

        if not isinstance(data, list):
            logger.warning(f"Cached {description} at {path} is not a list; ignoring contents")
            return []

        return data

    cached_filtered_with_license = _load_cached_list(FILTERED_LICENSE_FILE, 'filtered licensed users')
    cached_filtered_users = _load_cached_list(FILTERED_USERS_FILE, 'filtered users')

    return cached_employees, cached_filtered_with_license, cached_filtered_users


def update_employee_data(source: str = 'unknown') -> None:
    """Refresh all employee data from Microsoft Graph API."""
    success = False
    error_message = None
    existing_status = load_data_update_status()
    if existing_status.get('state') == 'running':
        logger.info(
            "Data update already running (current source: %s); skipping %s request",
            existing_status.get('source', 'unknown'),
            source,
        )
        return
    mark_data_update_running(source=source)
    try:
        # Ensure data directory exists and is writable
        if not os.path.exists(DATA_DIR):
            os.makedirs(DATA_DIR, exist_ok=True)
            logger.info(f"Created data directory: {DATA_DIR}")

        # Test if we can write to the data directory
        test_file = os.path.join(DATA_DIR, 'test_write.tmp')
        try:
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
        except Exception as e:
            logger.error(f"Cannot write to data directory {DATA_DIR}: {e}")
            error_message = f"Cannot write to data directory {DATA_DIR}: {e}"
            return

        logger.info(f"[{datetime.now()}] Starting employee data update...")

        token = get_access_token()
        if not token:
            logger.error("Unable to refresh employee data because access token retrieval failed")
            error_message = "Access token retrieval failed"
            return

        settings = load_settings()
        months_threshold = settings.get('newEmployeeMonths', 3)

        employees, filtered_with_license, filtered_users = fetch_all_employees(
            token=token,
            settings=settings,
            fallback_loader=_load_fetch_all_employees_fallback,
        )

        if employees:
            ignored_employee_set = parse_ignored_employees(settings)
            ignored_department_set = parse_ignored_departments(settings)

            if ignored_employee_set:
                before = len(employees)
                employees = [
                    emp for emp in employees
                    if not employee_is_ignored(
                        emp.get('name'),
                        emp.get('email'),
                        emp.get('userPrincipalName'),
                        ignored_employee_set
                    )
                ]
                if before != len(employees):
                    logger.info(f"Filtered ignored employees; {before}->{len(employees)} remaining")

            if ignored_department_set:
                before = len(employees)
                employees = [
                    emp for emp in employees
                    if not department_is_ignored(emp.get('department'), ignored_department_set)
                ]
                logger.info(
                    f"Filtered ignored departments {sorted(list(ignored_department_set))}; {before}->{len(employees)} employees"
                )

            try:
                with open(EMPLOYEE_LIST_FILE, 'w') as employee_cache:
                    json.dump(employees, employee_cache, indent=2)
                logger.info(f"Cached {len(employees)} employees for session-specific hierarchy builds")
            except Exception as cache_error:
                logger.error(f"Failed to write employee cache: {cache_error}")

            hierarchy = build_org_hierarchy(employees, settings=settings)

            if filtered_users:
                combined_by_id: dict[str, dict] = {}
                for record in employees:
                    record_id = record.get('id')
                    if record_id:
                        combined_by_id[str(record_id)] = record
                    else:
                        combined_by_id[f'anon-emp-{id(record)}'] = record
                for record in filtered_users:
                    candidate = dict(record)
                    candidate.setdefault('children', [])
                    record_id = candidate.get('id')
                    if record_id:
                        combined_by_id[str(record_id)] = candidate
                    else:
                        combined_by_id[f'anon-filtered-{id(record)}'] = candidate
                missing_source_records = list(combined_by_id.values())
            else:
                missing_source_records = employees

            missing_records = collect_missing_manager_records(missing_source_records, hierarchy, settings)

            if missing_records:
                enrichment_headers = {
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                }
                _enrich_mailbox_metadata(enrichment_headers, missing_records, max_lookups=0)

            if hierarchy:
                def update_new_status(node):
                    if node.get('hireDate'):
                        try:
                            hire_date = datetime.fromisoformat(node['hireDate'])
                            if hire_date.tzinfo:
                                cutoff_date = datetime.now(hire_date.tzinfo) - timedelta(days=months_threshold * 30)
                            else:
                                cutoff_date = datetime.now() - timedelta(days=months_threshold * 30)
                            node['isNewEmployee'] = hire_date > cutoff_date
                        except Exception:
                            node['isNewEmployee'] = False
                    else:
                        node['isNewEmployee'] = False

                    for child in node.get('children', []) or []:
                        update_new_status(child)

                update_new_status(hierarchy)

                with open(DATA_FILE, 'w') as f:
                    json.dump(hierarchy, f, indent=2)
                logger.info(f"[{datetime.now()}] Successfully updated employee data. Total employees: {len(employees)}")

                try:
                    with open(MISSING_MANAGER_FILE, 'w') as report_file:
                        json.dump(missing_records, report_file, indent=2)
                    logger.info(f"Updated missing manager report cache with {len(missing_records)} records")
                except Exception as report_error:
                    logger.error(f"Failed to write missing manager report cache: {report_error}")
            else:
                logger.error(f"[{datetime.now()}] Could not build hierarchy from employee data")

            try:
                recently_hired_records = collect_recently_hired_employees(employees, days=365)
                with open(RECENTLY_HIRED_FILE, 'w') as report_file:
                    json.dump(recently_hired_records, report_file, indent=2)
                logger.info(
                    f"Updated recently hired employees report cache with {len(recently_hired_records)} records"
                )
            except Exception as report_error:
                logger.error(f"Failed to write recently hired employees report cache: {report_error}")
        else:
            logger.error(f"[{datetime.now()}] No employees fetched from Graph API")

        try:
            filtered_user_records = filtered_users or []
            with open(FILTERED_USERS_FILE, 'w') as report_file:
                json.dump(filtered_user_records, report_file, indent=2)
            logger.info(
                f"Updated filtered users report cache with {len(filtered_user_records)} records"
            )
        except Exception as report_error:
            logger.error(f"Failed to write filtered users report cache: {report_error}")

        try:
            filtered_license_records = filtered_with_license or []
            with open(FILTERED_LICENSE_FILE, 'w') as report_file:
                json.dump(filtered_license_records, report_file, indent=2)
            logger.info(
                f"Updated filtered licensed users report cache with {len(filtered_license_records)} records"
            )
        except Exception as report_error:
            logger.error(f"Failed to write filtered licensed users report cache: {report_error}")

        try:
            last_login_records = collect_last_login_records(token=token)
            with open(LAST_LOGIN_FILE, 'w') as report_file:
                json.dump(last_login_records, report_file, indent=2)
            logger.info(
                f"Updated last sign-in report cache with {len(last_login_records)} records"
            )
        except Exception as report_error:
            logger.error(f"Failed to write last sign-in report cache: {report_error}")

        try:
            existing_disabled_records = []
            if os.path.exists(DISABLED_USERS_FILE):
                try:
                    with open(DISABLED_USERS_FILE, 'r') as previous_file:
                        data = json.load(previous_file)
                        if isinstance(data, list):
                            existing_disabled_records = data
                except Exception as previous_error:
                    logger.warning(f"Unable to load existing disabled users cache: {previous_error}")

            disabled_user_records = collect_disabled_users(
                token=token,
                previous_records=existing_disabled_records
            ) or []
            with open(DISABLED_USERS_FILE, 'w') as report_file:
                json.dump(disabled_user_records, report_file, indent=2)
            logger.info(
                f"Updated disabled users report cache with {len(disabled_user_records)} records"
            )
        except Exception as report_error:
            logger.error(f"Failed to write disabled users report cache: {report_error}")

        try:
            disabled_license_records = [
                record for record in disabled_user_records if (record.get('licenseCount') or 0) > 0
            ]

            with open(DISABLED_LICENSE_FILE, 'w') as report_file:
                json.dump(disabled_license_records, report_file, indent=2)
            logger.info(
                f"Updated disabled licensed users report cache with {len(disabled_license_records)} records"
            )
        except Exception as report_error:
            logger.error(f"Failed to write disabled licensed users report cache: {report_error}")

        try:
            recently_disabled_records = collect_recently_disabled_employees(disabled_user_records, days=365)
            with open(RECENTLY_DISABLED_FILE, 'w') as report_file:
                json.dump(recently_disabled_records, report_file, indent=2)
            logger.info(
                f"Updated recently disabled employees report cache with {len(recently_disabled_records)} records"
            )
        except Exception as report_error:
            logger.error(f"Failed to write recently disabled employees report cache: {report_error}")
        success = True
    except Exception as e:
        error_message = str(e)
        logger.error(f"[{datetime.now()}] Error updating employee data: {e}")
    finally:
        if success:
            mark_data_update_finished(success=True, source=source)
        else:
            mark_data_update_finished(success=False, error=error_message or 'Unknown error', source=source)


# Initialize on module load
load_data_update_status()
mark_startup_complete()


__all__ = [
    'load_data_update_status',
    'mark_data_update_running',
    'mark_data_update_finished',
    'mark_startup_complete',
    'collect_recently_disabled_employees',
    'collect_recently_hired_employees',
    'update_employee_data',
    '_load_fetch_all_employees_fallback',
]
