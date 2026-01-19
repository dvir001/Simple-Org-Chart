"""Export utilities for SimpleOrgChart (MicroSIP directory, etc.)."""

from __future__ import annotations

import logging
from datetime import datetime
from typing import Any, Dict, List, Optional, Set

from simple_org_chart.settings import load_settings

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


def _sanitize_contact_number(value: Any) -> str:
    """Return only the digits from a phone-like string."""
    if not value:
        return ''
    digits = ''.join(ch for ch in str(value) if ch.isdigit())
    return digits


def _split_name_parts(display_name: Optional[str]) -> tuple:
    """Split a full name into MicroSIP-friendly first/last fields."""
    if not display_name:
        return '', ''

    parts = [segment for segment in str(display_name).strip().split() if segment]
    if not parts:
        return '', ''

    if len(parts) == 1:
        return parts[0], ''

    return ' '.join(parts[:-1]), parts[-1]


def _parse_custom_directory_contacts(raw_text: Optional[str]) -> List[Dict[str, str]]:
    """Parse custom directory contacts from settings text."""
    contacts = []
    if not raw_text:
        return contacts

    lines = str(raw_text).splitlines()
    for index, line in enumerate(lines, start=1):
        stripped = line.strip()
        if not stripped or stripped.startswith('#'):
            continue

        if ',' not in stripped:
            logger.debug("Skipping custom contact line %s without comma delimiter", index)
            continue

        name_part, number_part = stripped.split(',', 1)
        name = name_part.strip()
        raw_number = number_part.strip()

        sanitized_number = _sanitize_contact_number(raw_number)
        if not sanitized_number:
            logger.debug("Skipping custom contact line %s without numeric phone digits", index)
            continue

        contacts.append({
            'name': name,
            'raw_number': raw_number,
            'sanitized_number': sanitized_number,
        })

    return contacts


def build_microsip_directory_items(
    employees: Optional[List[Dict[str, Any]]],
    *,
    settings: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """Transform cached employee records into MicroSIP directory entries."""
    items = []
    used_numbers: Set[str] = set()
    employee_numbers: Set[str] = set()
    fallback_counter = 1

    def _next_generated_number() -> str:
        nonlocal fallback_counter
        while True:
            candidate = f"{fallback_counter:03}"
            fallback_counter += 1
            if candidate not in used_numbers:
                return candidate

    for employee in sorted(employees or [], key=lambda item: (item.get('name') or '').lower()):
        business_phone = (employee.get('businessPhone') or '').strip()
        mobile_phone = (employee.get('phone') or '').strip()

        sanitized_business = _sanitize_contact_number(business_phone)
        sanitized_mobile = _sanitize_contact_number(mobile_phone)

        if not sanitized_business and not sanitized_mobile:
            continue

        preferred_number = sanitized_business or sanitized_mobile
        if not preferred_number or preferred_number in used_numbers:
            preferred_number = _next_generated_number()

        used_numbers.add(preferred_number)
        employee_numbers.add(preferred_number)

        full_name = (employee.get('name') or '').strip()
        first_name, last_name = _split_name_parts(full_name)
        email = (employee.get('email') or employee.get('userPrincipalName') or '').strip()
        comment_parts = []
        if employee.get('title'):
            comment_parts.append(employee.get('title'))
        if employee.get('department'):
            comment_parts.append(employee.get('department'))
        comment = ' - '.join(comment_parts)

        items.append({
            'number': preferred_number,
            'name': full_name or email or preferred_number,
            'firstname': first_name,
            'lastname': last_name,
            'phone': business_phone,
            'mobile': mobile_phone,
            'email': email,
            'address': (employee.get('fullAddress') or employee.get('officeLocation') or ''),
            'city': (employee.get('city') or ''),
            'state': (employee.get('state') or ''),
            'zip': '',
            'comment': comment,
            'presence': 0,
            'starred': 0,
            'info': employee.get('userPrincipalName') or '',
        })

    settings = settings or load_settings()
    custom_contacts = _parse_custom_directory_contacts((settings or {}).get('customDirectoryContacts'))
    custom_numbers: Set[str] = set()

    for contact in custom_contacts:
        preferred_number = contact['sanitized_number']

        if preferred_number:
            if preferred_number in custom_numbers:
                preferred_number = None
            else:
                custom_numbers.add(preferred_number)
                if preferred_number in employee_numbers:
                    logger.debug("Custom contact reusing existing employee number %s", preferred_number)

        if not preferred_number:
            preferred_number = _next_generated_number()

        used_numbers.add(preferred_number)

        full_name = (contact.get('name') or '').strip()
        first_name, last_name = _split_name_parts(full_name)

        items.append({
            'number': preferred_number,
            'name': full_name or contact['raw_number'] or preferred_number,
            'firstname': first_name,
            'lastname': last_name,
            'phone': contact['raw_number'],
            'mobile': '',
            'email': '',
            'address': '',
            'city': '',
            'state': '',
            'zip': '',
            'comment': '',
            'presence': 0,
            'starred': 0,
            'info': '',
        })

    return items


def build_yealink_phonebook_xml(
    employees: Optional[List[Dict[str, Any]]],
    *,
    settings: Optional[Dict[str, Any]] = None,
    title: str = "Organization Directory",
) -> str:
    """Generate Yealink-compatible remote phonebook XML from employee records.
    
    Compatible with Yealink T31P, T33G, T46U and similar models that support
    the YealinkIPPhoneDirectory XML schema for remote phonebooks.
    """
    from xml.etree.ElementTree import Element, SubElement, tostring

    settings = settings or load_settings()
    custom_contacts = _parse_custom_directory_contacts((settings or {}).get('customDirectoryContacts'))

    root = Element('YealinkIPPhoneDirectory')

    title_el = SubElement(root, 'Title')
    title_el.text = title

    used_numbers: Set[str] = set()

    # Process employees sorted alphabetically
    for employee in sorted(employees or [], key=lambda e: (e.get('name') or '').lower()):
        business_phone = (employee.get('businessPhone') or '').strip()
        mobile_phone = (employee.get('phone') or '').strip()

        if not business_phone and not mobile_phone:
            continue

        full_name = (employee.get('name') or '').strip()
        if not full_name:
            continue

        entry = SubElement(root, 'DirectoryEntry')

        name_el = SubElement(entry, 'Name')
        name_el.text = full_name

        # Add business phone as primary
        if business_phone:
            phone_el = SubElement(entry, 'Telephone')
            phone_el.text = business_phone
            used_numbers.add(_sanitize_contact_number(business_phone))

        # Add mobile as secondary if different from business
        if mobile_phone:
            sanitized_mobile = _sanitize_contact_number(mobile_phone)
            if sanitized_mobile not in used_numbers:
                mobile_el = SubElement(entry, 'Telephone')
                mobile_el.text = mobile_phone
                used_numbers.add(sanitized_mobile)

    # Process custom contacts
    for contact in custom_contacts:
        raw_number = contact.get('raw_number', '')
        sanitized = contact.get('sanitized_number', '')
        name = (contact.get('name') or '').strip()

        if not raw_number or not name:
            continue

        if sanitized in used_numbers:
            continue

        used_numbers.add(sanitized)

        entry = SubElement(root, 'DirectoryEntry')
        name_el = SubElement(entry, 'Name')
        name_el.text = name
        phone_el = SubElement(entry, 'Telephone')
        phone_el.text = raw_number

    return '<?xml version="1.0" encoding="UTF-8"?>\n' + tostring(root, encoding='unicode')


__all__ = [
    'format_hire_date',
    'build_microsip_directory_items',
    'build_yealink_phonebook_xml',
]
