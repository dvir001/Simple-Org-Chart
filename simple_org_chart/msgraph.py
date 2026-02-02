"""Microsoft Graph helpers for SimpleOrgChart."""

from __future__ import annotations

import logging
import os
import time
from datetime import datetime, timedelta, timezone
from typing import Callable, Iterable, Optional, Sequence, Tuple

import requests

from simple_org_chart.settings import (
    department_is_ignored,
    employee_is_ignored,
    load_settings,
    normalize_filter_value,
    parse_ignored_departments,
    parse_ignored_employees,
    parse_ignored_titles,
)


logger = logging.getLogger(__name__)

# Graph API endpoints (configurable via .env)
GRAPH_API_ENDPOINT = os.environ.get('GRAPH_API_ENDPOINT', 'https://graph.microsoft.com/v1.0')
GRAPH_API_BETA_ENDPOINT = os.environ.get('GRAPH_API_BETA_ENDPOINT', 'https://graph.microsoft.com/beta')

EmployeeTriple = Tuple[list[dict], list[dict], list[dict]]
FallbackLoader = Callable[[], EmployeeTriple]


def _enrich_mailbox_metadata(
    headers: dict,
    records: Iterable[dict],
    *,
    max_lookups: Optional[int] = 200,
) -> None:
    record_map: dict[str, list[dict]] = {}
    for record in records:
        if not isinstance(record, dict):
            continue
        user_id = record.get("id")
        if not user_id:
            continue
        record_map.setdefault(str(user_id), []).append(record)

    if not record_map:
        return

    effective_limit = None if max_lookups is None or max_lookups <= 0 else max_lookups
    lookups_performed = 0

    for user_id, record_group in record_map.items():
        if any((rec.get("mailboxType") or "").strip() for rec in record_group):
            continue
        if effective_limit is not None and lookups_performed >= effective_limit:
            break

        lookup_url = (
            f"{GRAPH_API_BETA_ENDPOINT}/users/{user_id}/mailboxSettings"
            "?$select=userPurpose"
        )

        if "ConsistencyLevel" in headers:
            enrichment_headers = headers
        else:
            enrichment_headers = dict(headers)
            enrichment_headers["ConsistencyLevel"] = "eventual"

        try:
            response = requests.get(lookup_url, headers=enrichment_headers, timeout=10)
        except requests.RequestException as exc:
            logger.debug("Failed to enrich mailbox settings for %s: %s", user_id, exc)
            continue

        if response.status_code in {401, 403}:
            logger.info(
                "Skipping mailbox enrichment; permission denied (status %s)",
                response.status_code,
            )
            break

        if response.status_code == 404:
            continue

        try:
            response.raise_for_status()
        except requests.HTTPError as exc:
            logger.debug("Mailbox settings lookup failed for %s: %s", user_id, exc)
            continue

        payload = response.json() or {}
        mailbox_purpose_raw = (payload.get("userPurpose") or "").strip()
        if not mailbox_purpose_raw:
            continue

        mailbox_purpose = mailbox_purpose_raw.lower()
        is_shared_mailbox = mailbox_purpose.startswith("shared") if mailbox_purpose else None

        for record in record_group:
            record["mailboxType"] = mailbox_purpose_raw
            record["isSharedMailbox"] = is_shared_mailbox

        lookups_performed += 1

    if lookups_performed:
        logger.info("Enriched mailbox metadata for %s users", lookups_performed)


def parse_graph_datetime(value: object) -> Optional[datetime]:
    """Cast a variety of Graph timestamp formats into aware datetimes."""
    if not value:
        return None

    dt: Optional[datetime] = None

    if isinstance(value, datetime):
        dt = value
    elif isinstance(value, (int, float)):
        try:
            dt = datetime.fromtimestamp(value, tz=timezone.utc)
        except Exception:  # pragma: no cover - defensive
            return None
    elif isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        try:
            if text.endswith("Z"):
                text = text[:-1] + "+00:00"
            dt = datetime.fromisoformat(text)
        except ValueError:
            for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
                try:
                    dt = datetime.strptime(text, fmt)
                    break
                except ValueError:
                    dt = None
            if dt is None:
                return None
    else:
        return None

    if dt is None:
        return None

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)

    return dt


def datetime_to_iso(dt: Optional[datetime]) -> Optional[str]:
    """Return ISO-8601 string in UTC for a datetime."""
    if not isinstance(dt, datetime):
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).isoformat()


def calculate_days_since(moment: object) -> Optional[int]:
    dt = parse_graph_datetime(moment)
    if not dt:
        return None
    now = datetime.now(timezone.utc)
    delta = now - dt.astimezone(timezone.utc)
    return max(delta.days, 0)


def _graph_credentials() -> Tuple[Optional[str], Optional[str], Optional[str]]:
    return (
        os.environ.get("AZURE_TENANT_ID"),
        os.environ.get("AZURE_CLIENT_ID"),
        os.environ.get("AZURE_CLIENT_SECRET"),
    )


def get_access_token(
    *,
    tenant_id: Optional[str] = None,
    client_id: Optional[str] = None,
    client_secret: Optional[str] = None,
) -> Optional[str]:
    """Exchange cached credentials for an application token."""

    tenant_id, client_id, client_secret = _resolve_credentials(tenant_id, client_id, client_secret)
    if not all([tenant_id, client_id, client_secret]):
        logger.error("AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET must be configured")
        return None

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }

    try:
        token_response = requests.post(token_url, data=token_data, timeout=10)
        token_response.raise_for_status()
        return token_response.json().get("access_token")
    except requests.RequestException as exc:  # pragma: no cover - network failures
        logger.error("Error getting access token: %s", exc)
        return None


def _resolve_credentials(
    tenant_id: Optional[str], client_id: Optional[str], client_secret: Optional[str]
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    if tenant_id and client_id and client_secret:
        return tenant_id, client_id, client_secret
    env_tenant, env_client, env_secret = _graph_credentials()
    return tenant_id or env_tenant, client_id or env_client, client_secret or env_secret


def fetch_employee_photo(user_id: str, token: str) -> Optional[bytes]:
    """Download an employee photo from Microsoft Graph."""
    try:
        photo_url = f"{GRAPH_API_ENDPOINT}/users/{user_id}/photo/$value"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(photo_url, headers=headers, timeout=10)
        if response.status_code == 200:
            return response.content
        logger.debug("No photo found for user %s (status %s)", user_id, response.status_code)
        return None
    except requests.RequestException as exc:  # pragma: no cover - network failures
        logger.debug("Error fetching photo for user %s: %s", user_id, exc)
        return None


def fetch_subscribed_sku_map(token: str) -> dict[str, str]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    sku_map: dict[str, str] = {}
    skus_url = f"{GRAPH_API_ENDPOINT}/subscribedSkus?$select=skuId,skuPartNumber"

    try:
        while skus_url:
            response = requests.get(skus_url, headers=headers, timeout=10)
            response.raise_for_status()
            data = response.json()
            for sku in data.get("value", []):
                sku_id = sku.get("skuId")
                if not sku_id:
                    continue
                key = str(sku_id).lower()
                sku_map[key] = sku.get("skuPartNumber") or str(sku_id)
            skus_url = data.get("@odata.nextLink")
    except requests.RequestException as exc:  # pragma: no cover
        logger.warning("Failed to load subscribed SKUs: %s", exc)
    except Exception as exc:  # pragma: no cover - unexpected formats
        logger.warning("Unexpected error loading subscribed SKUs: %s", exc)

    return sku_map


def fetch_all_employees(
    *,
    token: Optional[str] = None,
    settings: Optional[dict] = None,
    fallback_loader: Optional[FallbackLoader] = None,
) -> EmployeeTriple:
    token = token or get_access_token()

    if not token:
        logger.error("Failed to get access token")
        if fallback_loader:
            logger.warning("Using cached employee data because access token retrieval failed")
            return fallback_loader()
        return ([], [], [])

    settings = settings or load_settings()

    hide_disabled_users = settings.get("hideDisabledUsers", True)
    hide_guest_users = settings.get("hideGuestUsers", True)
    hide_no_title = settings.get("hideNoTitle", True)
    ignored_title_values = parse_ignored_titles(settings)
    ignored_employee_values = parse_ignored_employees(settings)
    ignored_department_values = parse_ignored_departments(settings)
    new_employee_months = settings.get("newEmployeeMonths", 3)

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    employees: list[dict] = []
    filtered_with_license: list[dict] = []
    filtered_users: list[dict] = []
    fetch_failed = False

    sku_map = fetch_subscribed_sku_map(token)

    select_fields = (
        "id,displayName,jobTitle,department,mail,userPrincipalName,mobilePhone,"
        "businessPhones,officeLocation,city,state,country,usageLocation,streetAddress,"
        "postalCode,employeeHireDate,accountEnabled,userType,assignedLicenses"
    )
    users_url = (
        f"{GRAPH_API_ENDPOINT}/users?$select={select_fields}"
        f"&$expand=manager($select=id,displayName)"
    )

    while users_url:
        try:
            response = requests.get(users_url, headers=headers, timeout=15)
            response.raise_for_status()
            data = response.json()
            if "value" not in data:
                break
            for user in data["value"]:
                display_name = user.get("displayName") or ""
                primary_email = user.get("mail") or ""
                user_principal_name = user.get("userPrincipalName") or ""
                job_title_val = user.get("jobTitle") or ""
                lowered_title = normalize_filter_value(job_title_val)
                department_val = user.get("department") or ""
                business_phones = user.get("businessPhones") or []
                if isinstance(business_phones, list):
                    business_phone = next((phone for phone in business_phones if phone), "")
                else:
                    business_phone = business_phones or ""

                assigned_licenses = user.get("assignedLicenses") or []
                user_type = (user.get("userType") or "").lower()

                license_sku_ids: list[str] = []
                license_labels: list[str] = []
                if assigned_licenses:
                    seen_labels: set[str] = set()
                    for license_entry in assigned_licenses:
                        sku_id = license_entry.get("skuId")
                        if not sku_id:
                            continue
                        sku_key = str(sku_id).lower()
                        license_sku_ids.append(str(sku_id))
                        friendly_name = (
                            sku_map.get(sku_key)
                            or sku_map.get(sku_key.upper())
                            or str(sku_id)
                        )
                        normalized_label = friendly_name.lower()
                        if normalized_label not in seen_labels:
                            seen_labels.add(normalized_label)
                            license_labels.append(friendly_name)
                    license_labels.sort(key=lambda item: item.lower())

                filtered_reasons: list[str] = []
                if hide_disabled_users and not user.get("accountEnabled", True):
                    filtered_reasons.append("filter_disabled")
                    logger.debug(
                        "Filtering disabled user: %s (accountEnabled=%s)",
                        display_name,
                        user.get("accountEnabled"),
                    )
                if hide_guest_users and user_type == "guest":
                    filtered_reasons.append("filter_guest")
                if hide_no_title and job_title_val.strip() == "":
                    filtered_reasons.append("filter_no_title")
                if ignored_title_values and lowered_title in ignored_title_values:
                    filtered_reasons.append("filter_ignored_title")
                if department_is_ignored(department_val, ignored_department_values):
                    filtered_reasons.append("filter_ignored_department")
                if employee_is_ignored(
                    display_name,
                    primary_email,
                    user_principal_name,
                    ignored_employee_values,
                ):
                    filtered_reasons.append("filter_ignored_employee")

                if filtered_reasons:
                    base_record = {
                        "id": user.get("id"),
                        "name": display_name or "Unknown",
                        "title": job_title_val or "No Title",
                        "department": department_val or "No Department",
                        "email": primary_email or user_principal_name or "",
                        "userPrincipalName": user_principal_name,
                        "phone": user.get("mobilePhone") or "",
                        "businessPhone": business_phone,
                        "location": user.get("officeLocation") or "",
                        "city": user.get("city") or "",
                        "state": user.get("state") or "",
                        "country": user.get("country") or "",
                        "usageLocation": user.get("usageLocation") or "",
                        "accountEnabled": user.get("accountEnabled", True),
                        "userType": user_type,
                        "filterReasons": filtered_reasons,
                        "licenseCount": len(license_sku_ids),
                        "licenseSkus": license_labels,
                        "licenseSkuIds": license_sku_ids,
                        "mailboxType": None,
                        "isSharedMailbox": None,
                        "managerId": user.get("manager", {}).get("id") if user.get("manager") else None,
                        "children": [],
                    }
                    filtered_users.append(base_record)
                    if license_sku_ids:
                        filtered_with_license.append(dict(base_record))
                    continue

                if display_name:
                    hire_date_str = user.get("employeeHireDate")
                    is_new = False
                    hire_date = None
                    if hire_date_str:
                        try:
                            if "T" in hire_date_str:
                                hire_date = datetime.fromisoformat(hire_date_str.replace("Z", "+00:00"))
                            else:
                                hire_date = datetime.strptime(hire_date_str, "%Y-%m-%d")
                                hire_date = hire_date.replace(tzinfo=None)
                            if hire_date.tzinfo:
                                cutoff_date = datetime.now(hire_date.tzinfo) - timedelta(days=new_employee_months * 30)
                            else:
                                cutoff_date = datetime.now() - timedelta(days=new_employee_months * 30)
                            is_new = hire_date > cutoff_date
                        except Exception as exc:  # pragma: no cover - defensive
                            logger.warning("Error parsing hire date for user %s: %s", user.get("displayName"), exc)

                    address_components: list[str] = []
                    if user.get("streetAddress"):
                        address_components.append(user.get("streetAddress"))
                    if user.get("city"):
                        address_components.append(user.get("city"))
                    if user.get("state"):
                        address_components.append(user.get("state"))
                    if user.get("postalCode"):
                        address_components.append(user.get("postalCode"))
                    if user.get("country"):
                        address_components.append(user.get("country"))

                    full_address = ", ".join(address_components) if address_components else ""
                    email_value = primary_email or user_principal_name or ""

                    employees.append(
                        {
                            "id": user.get("id"),
                            "name": display_name or "Unknown",
                            "title": user.get("jobTitle") or "No Title",
                            "department": department_val or "No Department",
                            "email": email_value,
                            "phone": user.get("mobilePhone") or "",
                            "businessPhone": business_phone,
                            "location": user.get("officeLocation") or "",
                            "officeLocation": user.get("officeLocation") or "",
                            "city": user.get("city") or "",
                            "state": user.get("state") or "",
                            "country": user.get("country") or "",
                            "fullAddress": full_address,
                            "managerId": user.get("manager", {}).get("id") if user.get("manager") else None,
                            "employeeHireDate": hire_date_str,
                            "hireDate": hire_date.isoformat() if hire_date else None,
                            "isNewEmployee": is_new,
                            "photoUrl": f"/api/photo/{user.get('id')}",
                            "userPrincipalName": user_principal_name,
                            "children": [],
                            "accountEnabled": user.get("accountEnabled", True),
                            "userType": user.get("userType") or "",
                            "usageLocation": user.get("usageLocation") or "",
                            "licenseCount": len(license_sku_ids),
                            "licenseSkus": list(license_labels),
                            "licenseSkuIds": list(license_sku_ids),
                            "mailboxType": None,
                            "isSharedMailbox": None,
                        }
                    )
            users_url = data.get("@odata.nextLink")
        except requests.RequestException as exc:
            fetch_failed = True
            logger.error("Error fetching employees: %s", exc)
            status_code = getattr(getattr(exc, "response", None), "status_code", None)
            if status_code == 401:
                logger.error("Authentication failed. Please check your credentials.")
            elif status_code == 403:
                logger.error("Permission denied. Ensure User.Read.All permission is granted.")
            break
        except Exception as exc:  # pragma: no cover - defensive
            fetch_failed = True
            logger.error("Unexpected error: %s", exc)
            break

    # Count filtered reasons for logging
    disabled_count = sum(1 for u in filtered_users if "filter_disabled" in (u.get("filterReasons") or []))
    guest_count = sum(1 for u in filtered_users if "filter_guest" in (u.get("filterReasons") or []))
    no_title_count = sum(1 for u in filtered_users if "filter_no_title" in (u.get("filterReasons") or []))

    logger.info(
        "Fetched %s employees from Graph API (filtered total %s: %s disabled, %s guests, %s no-title; licensed filtered %s)",
        len(employees),
        len(filtered_users),
        disabled_count,
        guest_count,
        no_title_count,
        len(filtered_with_license),
    )

    if (fetch_failed or not employees) and fallback_loader:
        fallback_employees, fallback_filtered_with_license, fallback_filtered_users = fallback_loader()
        if fallback_employees:
            logger.warning(
                "Using cached employee fallback after Graph fetch %s (%s records)",
                "failure" if fetch_failed else "returning no data",
                len(fallback_employees),
            )
            employees = fallback_employees
            if fallback_filtered_with_license:
                filtered_with_license = fallback_filtered_with_license
            if fallback_filtered_users:
                filtered_users = fallback_filtered_users
        elif fetch_failed:
            logger.warning("Graph fetch failed and no cached employee data is available")

    if filtered_users:
        _enrich_mailbox_metadata(headers, filtered_users, max_lookups=0)
        if filtered_with_license:
            mailbox_lookup: dict[str, tuple[Optional[str], Optional[bool]]] = {}
            for record in filtered_users:
                user_id = record.get("id")
                if not user_id:
                    continue
                mailbox_lookup[str(user_id)] = (
                    record.get("mailboxType"),
                    record.get("isSharedMailbox"),
                )

            for record in filtered_with_license:
                user_id = record.get("id")
                if not user_id:
                    continue
                mailbox_type, shared_flag = mailbox_lookup.get(str(user_id), (None, None))
                if mailbox_type is not None:
                    record["mailboxType"] = mailbox_type
                if shared_flag is not None or "isSharedMailbox" in record:
                    record["isSharedMailbox"] = shared_flag

    return employees, filtered_with_license, filtered_users


def collect_last_login_records(*, token: Optional[str] = None) -> list[dict]:
    token = token or get_access_token()
    if not token:
        logger.error("Failed to get access token for last sign-in report")
        return []

    sku_map = fetch_subscribed_sku_map(token)

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "ConsistencyLevel": "eventual",
    }

    base_fields = (
        "id,displayName,jobTitle,department,mail,userPrincipalName,"
        "signInActivity,accountEnabled,userType,assignedLicenses"
    )

    def build_users_url(select_fields: str) -> str:
        return f"{GRAPH_API_BETA_ENDPOINT}/users?$select={select_fields}&$top=999"

    users_url = build_users_url(base_fields)

    now_utc = datetime.now(timezone.utc)
    records: list[dict] = []

    def _format_datetime(dt: Optional[datetime]) -> Optional[str]:
        if not dt:
            return None
        if not dt.tzinfo:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc).isoformat()

    def _map_licenses(license_entries: Optional[Iterable[dict]]) -> Tuple[list[str], list[str]]:
        license_entries = license_entries or []
        if not license_entries:
            return [], []
        sku_ids: list[str] = []
        labels: list[str] = []
        seen_labels: set[str] = set()
        for entry in license_entries:
            sku_id = entry.get("skuId")
            if not sku_id:
                continue
            sku_ids.append(str(sku_id))
            lookup_key = str(sku_id).lower()
            friendly = (
                sku_map.get(lookup_key)
                or sku_map.get(lookup_key.upper())
                or str(sku_id)
            )
            normalized = friendly.lower()
            if normalized not in seen_labels:
                seen_labels.add(normalized)
                labels.append(friendly)
        labels.sort(key=lambda item: item.lower())
        return sku_ids, labels

    while users_url:
        try:
            response = requests.get(users_url, headers=headers, timeout=20)
        except requests.RequestException as exc:
            logger.error("Failed to fetch sign-in activity: %s", exc)
            break

        if response.status_code == 429:
            retry_after = response.headers.get("Retry-After")
            delay = 5
            try:
                parsed = int(retry_after)
                delay = max(parsed, delay)
            except Exception:
                pass
            logger.warning("Graph throttled sign-in activity request; retrying in %s seconds", delay)
            time.sleep(delay)
            continue

        try:
            response.raise_for_status()
        except requests.HTTPError as exc:
            status_code = getattr(exc.response, "status_code", None)
            logger.error(
                "Graph error fetching sign-in activity (status %s): %s",
                status_code,
                exc,
            )
            break

        payload = response.json()

        for user in payload.get("value", []):
            sign_in = user.get("signInActivity") or {}

            last_combined = parse_graph_datetime(sign_in.get("lastSignInDateTime"))
            last_interactive = parse_graph_datetime(sign_in.get("lastInteractiveSignInDateTime"))
            last_non_interactive = parse_graph_datetime(sign_in.get("lastNonInteractiveSignInDateTime"))

            observed_dates = [dt for dt in (last_combined, last_interactive, last_non_interactive) if dt]
            most_recent = max(observed_dates) if observed_dates else None

            sku_ids, license_labels = _map_licenses(user.get("assignedLicenses"))

            mailbox_settings = user.get("mailboxSettings") or {}
            mailbox_purpose_raw = (mailbox_settings.get("userPurpose") or "").strip()
            mailbox_purpose = mailbox_purpose_raw.lower()
            is_shared_mailbox = None
            if mailbox_purpose:
                is_shared_mailbox = mailbox_purpose.startswith("shared")

            user_id = user.get("id")

            record = {
                "id": user_id,
                "name": user.get("displayName") or "Unknown",
                "title": user.get("jobTitle") or "No Title",
                "department": user.get("department") or "No Department",
                "email": user.get("mail") or user.get("userPrincipalName") or "",
                "accountEnabled": user.get("accountEnabled", True),
                "userType": (user.get("userType") or "").lower(),
                "licenseCount": len(sku_ids),
                "licenseSkus": license_labels,
                "licenseSkuIds": sku_ids,
                "mailboxType": mailbox_purpose_raw or None,
                "isSharedMailbox": is_shared_mailbox,
                "lastActivityDate": _format_datetime(most_recent),
                "daysSinceLastActivity": int((now_utc - most_recent).days) if most_recent else None,
                "lastInteractiveSignIn": _format_datetime(last_interactive),
                "daysSinceInteractiveSignIn": int((now_utc - last_interactive).days) if last_interactive else None,
                "lastNonInteractiveSignIn": _format_datetime(last_non_interactive),
                "daysSinceNonInteractiveSignIn": int((now_utc - last_non_interactive).days) if last_non_interactive else None,
                "neverSignedIn": not observed_dates,
            }
            records.append(record)

        users_url = payload.get("@odata.nextLink")

    if records:
        _enrich_mailbox_metadata(headers, records)

    logger.info("Collected %s last sign-in records", len(records))
    return records


def _collect_disabled_users(*, token: Optional[str] = None) -> list[dict]:
    token = token or get_access_token()
    if not token:
        logger.error("Failed to get access token for disabled user reports")
        return []

    sku_map = fetch_subscribed_sku_map(token)

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    select_fields = (
        "id,displayName,jobTitle,department,mail,userPrincipalName,mobilePhone,"
        "businessPhones,officeLocation,city,state,country,usageLocation,streetAddress,"
        "postalCode,employeeHireDate,employeeLeaveDateTime,accountEnabled,userType,assignedLicenses"
    )

    users_url = f"{GRAPH_API_ENDPOINT}/users?$select={select_fields}&$filter=accountEnabled eq false"
    records: list[dict] = []

    while users_url:
        try:
            response = requests.get(users_url, headers=headers, timeout=15)
            response.raise_for_status()
            data = response.json()
            for user in data.get("value", []):
                display_name = user.get("displayName") or ""
                primary_email = user.get("mail") or ""
                user_principal_name = user.get("userPrincipalName") or ""
                job_title_val = user.get("jobTitle") or ""
                department_val = user.get("department") or "No Department"

                business_phones = user.get("businessPhones") or []
                if isinstance(business_phones, list):
                    business_phone = next((phone for phone in business_phones if phone), "")
                else:
                    business_phone = business_phones or ""

                assigned_licenses = user.get("assignedLicenses") or []
                license_sku_ids: list[str] = []
                license_labels: list[str] = []
                for license_entry in assigned_licenses:
                    sku_id = license_entry.get("skuId")
                    if not sku_id:
                        continue
                    sku_key = str(sku_id).lower()
                    license_sku_ids.append(str(sku_id))
                    friendly_name = (
                        sku_map.get(sku_key)
                        or sku_map.get(sku_key.upper())
                        or str(sku_id)
                    )
                    license_labels.append(friendly_name)
                license_labels = sorted(set(license_labels), key=lambda item: item.lower())

                disabled_at = parse_graph_datetime(user.get("employeeLeaveDateTime"))
                disabled_iso = datetime_to_iso(disabled_at) if disabled_at else None
                hire_date = parse_graph_datetime(user.get("employeeHireDate"))

                records.append(
                    {
                        "id": user.get("id"),
                        "name": display_name or "Unknown",
                        "title": job_title_val or "No Title",
                        "department": department_val,
                        "email": primary_email or user_principal_name or "",
                        "userPrincipalName": user_principal_name,
                        "phone": user.get("mobilePhone") or "",
                        "businessPhone": business_phone,
                        "location": user.get("officeLocation") or "",
                        "city": user.get("city") or "",
                        "state": user.get("state") or "",
                        "country": user.get("country") or "",
                        "usageLocation": user.get("usageLocation") or "",
                        "accountEnabled": user.get("accountEnabled", True),
                        "userType": (user.get("userType") or "").lower(),
                        "licenseCount": len(license_sku_ids),
                        "licenseSkus": license_labels,
                        "licenseSkuIds": license_sku_ids,
                        "hireDate": datetime_to_iso(hire_date) if hire_date else None,
                        "disabledDate": disabled_iso,
                        "disabledDays": calculate_days_since(disabled_at),
                    }
                )
            users_url = data.get("@odata.nextLink")
        except requests.RequestException as exc:
            logger.error("Error fetching disabled users: %s", exc)
            break
        except Exception as exc:  # pragma: no cover
            logger.error("Unexpected error while collecting disabled user data: %s", exc)
            break

    logger.info("Collected %s disabled users", len(records))
    return records


def collect_disabled_users(
    *,
    token: Optional[str] = None,
    previous_records: Optional[Sequence[dict]] = None,
) -> list[dict]:
    raw_records = _collect_disabled_users(token=token)

    previous_map: dict[str, dict] = {}
    if previous_records:
        for entry in previous_records:
            entry_id = entry.get("id")
            if entry_id:
                previous_map[entry_id] = entry

    now_iso = datetime_to_iso(datetime.now(timezone.utc))

    for record in raw_records:
        record_id = record.get("id")
        existing = previous_map.get(record_id) if record_id else None
        observed_source = record.get("disabledDate")
        existing_observed = None
        if existing:
            existing_observed = existing.get("firstSeenDisabledAt") or existing.get("disabledDate")
        if observed_source:
            first_seen = observed_source
        elif existing_observed:
            first_seen = existing_observed
        else:
            first_seen = now_iso
        record["firstSeenDisabledAt"] = first_seen
        if not record.get("disabledDate"):
            record["disabledDate"] = first_seen
        record["disabledDays"] = calculate_days_since(first_seen)

    return raw_records


def collect_disabled_licensed_users(
    *,
    token: Optional[str] = None,
    previous_records: Optional[Sequence[dict]] = None,
) -> list[dict]:
    raw_records = collect_disabled_users(token=token, previous_records=previous_records)
    licensed_records = [record for record in raw_records if (record.get("licenseCount") or 0) > 0]
    logger.info("Filtered %s disabled users with active licenses", len(licensed_records))
    return licensed_records


__all__ = [
    "calculate_days_since",
    "collect_disabled_licensed_users",
    "collect_disabled_users",
    "collect_last_login_records",
    "datetime_to_iso",
    "fetch_all_employees",
    "fetch_employee_photo",
    "fetch_subscribed_sku_map",
    "get_access_token",
    "parse_graph_datetime",
]
