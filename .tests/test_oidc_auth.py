"""Tests for Microsoft Entra authentication helpers."""

from unittest.mock import Mock, patch

from simple_org_chart.oidc_auth import OidcConfig, fetch_user_and_groups, resolve_role


def _config():
    return OidcConfig(
        tenant_id="tenant",
        client_id="client",
        client_secret="secret",
        reader_group_id="reader-group",
        privileged_group_id="privileged-group",
        admin_group_id="admin-group",
    )


def test_resolve_role_rejects_user_outside_access_groups():
    assert resolve_role(_config(), {"unrelated-group"}) is None


def test_resolve_role_matches_reader_group():
    assert resolve_role(_config(), {"reader-group"}) == "reader"


def test_resolve_role_matches_privileged_group_case_insensitively():
    assert resolve_role(_config(), {"PRIVILEGED-GROUP"}) == "privileged"


def test_admin_group_takes_precedence():
    assert resolve_role(_config(), {"privileged-group", "admin-group"}) == "admin"


def test_resolve_role_ignores_config_whitespace():
    config = OidcConfig(
        tenant_id="tenant",
        client_id="client",
        client_secret="secret",
        reader_group_id=" reader-group ",
        privileged_group_id=" privileged-group ",
        admin_group_id=" admin-group ",
    )
    assert resolve_role(config, {"privileged-group"}) == "privileged"


@patch("simple_org_chart.oidc_auth.requests.post")
@patch("simple_org_chart.oidc_auth.requests.get")
def test_fetch_user_and_groups_checks_only_configured_groups(mock_get, mock_post):
    user_response = Mock()
    user_response.json.return_value = {"id": "user", "displayName": "Test User"}
    mock_get.return_value = user_response
    group_response = Mock()
    group_response.json.return_value = {"value": ["PRIVILEGED-GROUP"]}
    mock_post.return_value = group_response

    user, group_ids = fetch_user_and_groups(
        "token",
        {" privileged-group ", "admin-group"},
    )

    assert user["id"] == "user"
    assert group_ids == {"privileged-group"}
    mock_post.assert_called_once_with(
        "https://graph.microsoft.com/v1.0/me/checkMemberGroups",
        headers={
            "Authorization": "Bearer token",
            "Content-Type": "application/json",
        },
        json={"groupIds": ["admin-group", "privileged-group"]},
        timeout=15,
    )