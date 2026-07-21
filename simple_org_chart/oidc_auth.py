"""Microsoft Entra OpenID Connect authentication helpers."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import msal
import requests

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
OIDC_SCOPES = ["User.Read", "GroupMember.Read.All"]


@dataclass(frozen=True)
class OidcConfig:
    tenant_id: str
    client_id: str
    client_secret: str
    reader_group_id: str
    privileged_group_id: str
    admin_group_id: str

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"


def build_client(config: OidcConfig) -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        config.client_id,
        authority=config.authority,
        client_credential=config.client_secret,
    )


def begin_login(config: OidcConfig, redirect_uri: str) -> dict[str, Any]:
    return build_client(config).initiate_auth_code_flow(
        OIDC_SCOPES,
        redirect_uri=redirect_uri,
    )


def complete_login(
    config: OidcConfig,
    flow: dict[str, Any],
    callback_args: dict[str, str],
) -> dict[str, Any]:
    return build_client(config).acquire_token_by_auth_code_flow(flow, callback_args)


def fetch_user_and_groups(
    access_token: str,
    configured_group_ids: set[str],
) -> tuple[dict[str, Any], set[str]]:
    headers = {"Authorization": f"Bearer {access_token}"}
    user_response = requests.get(
        f"{GRAPH_BASE_URL}/me?$select=id,displayName,userPrincipalName,mail",
        headers=headers,
        timeout=15,
    )
    user_response.raise_for_status()

    normalized_group_ids = {
        group_id.strip().lower()
        for group_id in configured_group_ids
        if group_id and group_id.strip()
    }
    group_response = requests.post(
        f"{GRAPH_BASE_URL}/me/checkMemberGroups",
        headers={**headers, "Content-Type": "application/json"},
        json={"groupIds": sorted(normalized_group_ids)},
        timeout=15,
    )
    group_response.raise_for_status()
    group_ids = {
        str(group_id).lower()
        for group_id in group_response.json().get("value", [])
    }

    return user_response.json(), group_ids


def resolve_role(config: OidcConfig, group_ids: set[str]) -> str | None:
    normalized_ids = {group_id.lower() for group_id in group_ids}
    if config.admin_group_id.strip().lower() in normalized_ids:
        return "admin"
    if config.privileged_group_id.strip().lower() in normalized_ids:
        return "privileged"
    if config.reader_group_id.strip().lower() in normalized_ids:
        return "reader"
    return None