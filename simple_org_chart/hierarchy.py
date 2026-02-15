"""Org hierarchy building and employee utilities for SimpleOrgChart."""

from __future__ import annotations

import json
import logging
import os
from typing import Any, Dict, List, Optional

import simple_org_chart.config as app_config
from simple_org_chart.settings import load_settings, normalize_filter_value

logger = logging.getLogger(__name__)

DATA_FILE = str(app_config.DATA_FILE)
EMPLOYEE_LIST_FILE = str(app_config.EMPLOYEE_LIST_FILE)


def build_org_hierarchy(
    employees: List[Dict[str, Any]],
    *,
    top_user_email_override: Optional[str] = None,
    settings: Optional[Dict[str, Any]] = None,
) -> Optional[Dict[str, Any]]:
    """Build an org chart hierarchy from a flat list of employees."""
    if not employees:
        return None

    if settings is None:
        settings = load_settings()

    # Read from settings (topLevelUserEmail / topLevelUserId)
    settings_top_user_email = (settings.get('topLevelUserEmail') or '').strip()
    settings_top_user_id = (settings.get('topLevelUserId') or '').strip()

    if top_user_email_override is not None:
        chosen_top_user = (top_user_email_override or '').strip()
    else:
        chosen_top_user = settings_top_user_email

    top_user_email = (chosen_top_user or '').strip() or None

    # Debug logging
    logger.info(f"Settings topLevelUserEmail: '{settings_top_user_email}'")
    logger.info(f"Settings topLevelUserId: '{settings_top_user_id}'")
    if top_user_email_override is not None:
        logger.info(f"Session override topUserEmail: '{top_user_email_override}'")
    logger.info(f"Final top_user_email: '{top_user_email}'")

    emp_dict = {emp['id']: emp.copy() for emp in employees}

    for emp_id in emp_dict:
        if 'children' not in emp_dict[emp_id]:
            emp_dict[emp_id]['children'] = []

    # First, check if a specific top-level user is configured
    root = None
    if top_user_email:
        logger.info(f"Searching for user with email: '{top_user_email}' among {len(employees)} employees")
        for emp in employees:
            if emp.get('email') == top_user_email:
                root = emp_dict[emp['id']]
                logger.info(f"Found and using configured top-level user by email: {root['name']} ({root.get('email')})")
                break
        else:
            logger.warning(f"Could not find user with email '{top_user_email}' in employee list")

    # Fallback to settings user ID if no email-based selection was made
    if not root and settings_top_user_id and settings_top_user_id in emp_dict:
        root = emp_dict[settings_top_user_id]
        logger.info(f"Using fallback settings top-level user by ID: {root['name']}")

    if root:
        # If a specific root is configured, build hierarchy with that person at the top
        # Clear any existing manager relationship for the root user
        root['managerId'] = None

        # Build the hierarchy normally but ensure the selected root has no manager
        for emp in employees:
            emp_copy = emp_dict[emp['id']]
            if emp_copy['id'] == root['id']:
                continue  # Skip the root user in hierarchy building

            if emp['managerId'] and emp['managerId'] in emp_dict:
                manager = emp_dict[emp['managerId']]
                if emp_copy not in manager['children']:
                    manager['children'].append(emp_copy)

        # Remove the selected root from anyone's children list (in case they were someone's subordinate)
        for emp_id, emp in emp_dict.items():
            emp['children'] = [child for child in emp['children'] if child['id'] != root['id']]

        return root
    else:
        # Auto-detect root using existing logic
        root_candidates = []

        # Build normal manager-employee relationships
        for emp in employees:
            emp_copy = emp_dict[emp['id']]
            if emp['managerId'] and emp['managerId'] in emp_dict:
                manager = emp_dict[emp['managerId']]
                if emp_copy not in manager['children']:
                    manager['children'].append(emp_copy)
            else:
                if not emp['managerId'] and emp_copy not in root_candidates:
                    root_candidates.append(emp_copy)

        # Auto-detect root
        if root_candidates:
            ceo_keywords = ['chief executive', 'ceo', 'president', 'chair', 'director', 'head']
            for candidate in root_candidates:
                title_lower = (candidate.get('title') or '').lower()
                if any(keyword in title_lower for keyword in ceo_keywords):
                    root = candidate
                    logger.info(
                        "Auto-detected top-level user based on title keywords among %d root candidates",
                        len(root_candidates),
                    )
                    break

            if not root and root_candidates:
                root = root_candidates[0]
                logger.info(
                    "Using first root candidate as top-level (no title keyword match found; %d candidates considered)",
                    len(root_candidates),
                )
        else:
            max_reports = 0
            for emp_id, emp in emp_dict.items():
                if len(emp['children']) > max_reports:
                    max_reports = len(emp['children'])
                    root = emp

            if root:
                logger.info(f"Using person with most reports as top-level: {root['name']} ({max_reports} reports)")

        if not root and employees:
            root = emp_dict[employees[0]['id']]
            logger.info("Using first employee in list as root (no explicit or inferred top-level user found)")

        return root


def collect_missing_manager_records(
    employees: List[Dict[str, Any]],
    hierarchy_root: Optional[Dict[str, Any]] = None,
    settings: Optional[Dict[str, Any]] = None,
    top_user_email_override: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Collect employees missing a valid manager reference."""
    if not employees:
        return []

    employee_index = {emp['id']: emp for emp in employees if emp.get('id')}
    visited = set()

    def traverse(node):
        node_id = node.get('id')
        if not node_id or node_id in visited:
            return
        visited.add(node_id)
        for child in node.get('children', []):
            traverse(child)

    if hierarchy_root:
        traverse(hierarchy_root)

    root_ids = set()
    top_user_email = None

    if hierarchy_root and hierarchy_root.get('id'):
        root_ids.add(hierarchy_root['id'])

    if settings is None:
        settings = load_settings()

    if top_user_email_override is not None:
        top_user_email = (top_user_email_override or '').strip().lower() or None
    elif settings:
        top_user_email = (settings.get('topLevelUserEmail') or '').strip().lower() or None

    missing_records = []

    for emp in employees:
        emp_id = emp.get('id')
        manager_id = emp.get('managerId')
        manager_name = ''
        reason = None

        if emp_id and emp_id in root_ids:
            continue

        if top_user_email:
            email = (emp.get('email') or '').strip().lower()
            if email and email == top_user_email:
                continue

        if manager_id and manager_id in employee_index:
            manager_name = employee_index[manager_id].get('name') or ''

        if not manager_id:
            reason = 'no_manager'
        elif manager_id not in employee_index:
            reason = 'manager_not_found'
        elif emp_id not in visited:
            reason = 'detached'

        if reason:
            filter_reasons = list(emp.get('filterReasons') or [])
            effective_reason = reason
            if filter_reasons:
                effective_reason = 'filtered'

            missing_records.append({
                'id': emp_id,
                'name': emp.get('name'),
                'title': emp.get('title'),
                'department': emp.get('department'),
                'email': emp.get('email'),
                'phone': emp.get('phone'),
                'businessPhone': emp.get('businessPhone'),
                'location': emp.get('location') or emp.get('officeLocation') or '',
                'managerName': manager_name,
                'reason': effective_reason,
                'missingReason': reason,
                'filterReasons': filter_reasons,
                'accountEnabled': emp.get('accountEnabled', True),
                'userType': (emp.get('userType') or '').lower(),
                'licenseCount': emp.get('licenseCount') or 0,
                'licenseSkus': list(emp.get('licenseSkus') or []),
                'licenseSkuIds': list(emp.get('licenseSkuIds') or []),
                'mailboxType': emp.get('mailboxType'),
                'isSharedMailbox': emp.get('isSharedMailbox'),
            })

    missing_records.sort(key=lambda item: (item.get('department') or '', item.get('name') or ''))
    return missing_records


def load_cached_employees() -> Optional[List[Dict[str, Any]]]:
    """Load cached employee list from disk."""
    if os.path.exists(EMPLOYEE_LIST_FILE):
        try:
            with open(EMPLOYEE_LIST_FILE, 'r') as cache_file:
                return json.load(cache_file)
        except Exception as e:
            logger.error(f"Failed to read employee cache {EMPLOYEE_LIST_FILE}: {e}")
    return None


def flatten_hierarchy_to_employee_list(root_node: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Walk a hierarchy tree and return a flat list of employees."""
    employees = []

    def _walk(node):
        if not isinstance(node, dict):
            return

        entry = {k: v for k, v in node.items() if k != 'children'}
        entry['children'] = []
        employees.append(entry)

        for child in node.get('children', []) or []:
            _walk(child)

    if root_node:
        _walk(root_node)

    return employees


def get_employee_list_for_metadata() -> List[Dict[str, Any]]:
    """Get flat employee list, preferring cached list over hierarchy walk."""
    employees = load_cached_employees()
    if employees:
        return employees

    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r') as data_file:
                hierarchy = json.load(data_file)
            if hierarchy:
                return flatten_hierarchy_to_employee_list(hierarchy)
        except Exception as error:
            logger.error(f"Failed to read hierarchy for metadata: {error}")

    return []


def collect_unique_field_values(
    employees: Optional[List[Dict[str, Any]]],
    field_name: str,
) -> List[str]:
    """Collect unique, deduplicated values for a given field."""
    unique = {}
    for employee in employees or []:
        value = (employee.get(field_name) or '').strip()
        if not value:
            continue
        key = value.lower()
        if key not in unique:
            unique[key] = value

    return sorted(unique.values(), key=lambda item: item.lower())


def collect_employee_option_labels(
    employees: Optional[List[Dict[str, Any]]],
) -> List[str]:
    """Collect employee labels for dropdown options (name + email)."""
    options = {}
    for employee in employees or []:
        name = (employee.get('name') or '').strip()
        email = (employee.get('email') or '').strip()
        user_principal_name = (employee.get('userPrincipalName') or '').strip()

        contact = email or user_principal_name

        if not name and not contact:
            continue

        if name and contact:
            label = f"{name} <{contact}>"
        else:
            label = name or contact

        primary_key = normalize_filter_value(contact) or normalize_filter_value(name) or normalize_filter_value(label)
        if not primary_key:
            continue

        if primary_key not in options:
            options[primary_key] = label

    return sorted(options.values(), key=lambda item: item.lower())


__all__ = [
    'build_org_hierarchy',
    'collect_missing_manager_records',
    'load_cached_employees',
    'flatten_hierarchy_to_employee_list',
    'get_employee_list_for_metadata',
    'collect_unique_field_values',
    'collect_employee_option_labels',
]
