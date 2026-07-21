"""Tests for simple_org_chart.auth – decorators & path sanitisation."""

from __future__ import annotations

from io import BytesIO
import json
import pytest
from flask import Flask
from unittest.mock import patch
from openpyxl import load_workbook
from simple_org_chart.auth import (
    privileged_login_required,
    require_auth,
    require_privileged,
    sanitize_next_path,
)
from simple_org_chart.app_main import app as org_chart_app


# ---------------------------------------------------------------------------
# sanitize_next_path
# ---------------------------------------------------------------------------


class TestSanitizeNextPath:
    """Validate redirect-path sanitisation against open-redirect attacks."""

    @pytest.mark.parametrize(
        "raw, expected",
        [
            (None, ""),
            ("", ""),
            ("   ", ""),
            ("configure", "configure"),
            ("/configure", "configure"),
            ("reports", "reports"),
            ("/reports", "reports"),
        ],
    )
    def test_valid_paths(self, raw, expected):
        assert sanitize_next_path(raw) == expected

    @pytest.mark.parametrize(
        "raw",
        [
            "http://evil.com",
            "https://evil.com",
            "//evil.com",
            "http://evil.com/configure",
            "/path with spaces",
            "/path?query=1",
            "/path#fragment",
            "/<script>alert(1)</script>",
        ],
    )
    def test_malicious_paths_rejected(self, raw):
        assert sanitize_next_path(raw) == ""


@pytest.fixture()
def role_client():
    app = Flask(__name__)
    app.config.update(SECRET_KEY="test", TESTING=True)

    @app.get("/login")
    def login():
        return "login"

    @app.get("/privileged-api")
    @require_privileged
    def privileged_api():
        return "ok"

    @app.get("/admin-api")
    @require_auth
    def admin_api():
        return "ok"

    @app.get("/reports")
    @privileged_login_required
    def reports():
        return "ok"

    return app.test_client()


def _set_role(client, role=None):
    with client.session_transaction() as user_session:
        user_session["authenticated"] = True
        if role:
            user_session["role"] = role


class TestRoleAuthorization:
    def test_unauthenticated_api_returns_401(self, role_client):
        assert role_client.get("/privileged-api").status_code == 401

    def test_reader_cannot_access_privileged_or_admin_api(self, role_client):
        _set_role(role_client, "reader")
        assert role_client.get("/privileged-api").status_code == 403
        assert role_client.get("/admin-api").status_code == 403

    def test_privileged_role_cannot_access_admin_api(self, role_client):
        _set_role(role_client, "privileged")
        assert role_client.get("/privileged-api").status_code == 200
        assert role_client.get("/admin-api").status_code == 403

    def test_admin_role_can_access_all_tiers(self, role_client):
        _set_role(role_client, "admin")
        assert role_client.get("/privileged-api").status_code == 200
        assert role_client.get("/admin-api").status_code == 200

    def test_legacy_authenticated_session_is_admin(self, role_client):
        _set_role(role_client)
        assert role_client.get("/admin-api").status_code == 200

    def test_reader_gets_forbidden_report_page(self, role_client):
        _set_role(role_client, "reader")
        assert role_client.get("/reports").status_code == 403


@pytest.fixture()
def app_client():
    org_chart_app.config.update(TESTING=True)
    return org_chart_app.test_client()


class TestApplicationRouteTiers:
    def test_reader_cannot_open_reports_or_configure(self, app_client):
        _set_role(app_client, "reader")
        assert app_client.get("/reports").status_code == 403
        assert app_client.get("/configure").status_code == 403

    def test_privileged_user_can_open_reports_but_not_configure(self, app_client):
        _set_role(app_client, "privileged")
        assert app_client.get("/reports").status_code == 200
        assert app_client.get("/configure").status_code == 403

    def test_admin_can_open_configure(self, app_client):
        _set_role(app_client, "admin")
        assert app_client.get("/configure").status_code == 200

    def test_auth_check_returns_capabilities(self, app_client):
        _set_role(app_client, "privileged")
        response = app_client.get("/api/auth-check")
        assert response.status_code == 200
        assert response.get_json() == {
            "authenticated": True,
            "authType": "simple",
            "canAccessReports": True,
            "canAccessRestrictedXlsx": True,
            "canAdminister": False,
            "canSync": True,
            "role": "privileged",
        }

    def test_graph_capabilities_uses_persisted_sync_data(self, app_client, tmp_path):
        capabilities_file = tmp_path / "graph_capabilities.json"
        capabilities_file.write_text(
            json.dumps({"mailbox_settings_read": True}),
            encoding="utf-8",
        )
        _set_role(app_client, "privileged")

        with patch(
            "simple_org_chart.app_main.GRAPH_CAPABILITIES_FILE",
            str(capabilities_file),
        ):
            response = app_client.get("/api/graph-capabilities")

        assert response.status_code == 200
        assert response.get_json() == {
            "available": True,
            "mailbox_settings_read": True,
        }

    def test_oidc_auth_check_returns_current_user_identity(self, app_client):
        with app_client.session_transaction() as user_session:
            user_session["authenticated"] = True
            user_session["role"] = "reader"
            user_session["oidc_session_version"] = 3
            user_session["oidc_user_id"] = "graph-user-id"
            user_session["oidc_user_email"] = "user@example.com"

        with patch("simple_org_chart.app_main.AUTH_TYPE", "oidc"):
            response = app_client.get("/api/auth-check")

        assert response.status_code == 200
        assert response.get_json()["user"] == {
            "id": "graph-user-id",
            "email": "user@example.com",
        }

    def test_privileged_permissions_can_be_disabled_independently(self, app_client):
        _set_role(app_client, "privileged")
        settings = {
            "oidcPrivilegedPermissions": {
                "reports": False,
                "sync": False,
                "restrictedXlsx": True,
            }
        }
        with patch("simple_org_chart.app_main.load_settings", return_value=settings):
            auth_response = app_client.get("/api/auth-check")
            assert auth_response.get_json()["canAccessReports"] is False
            assert auth_response.get_json()["canSync"] is False
            assert auth_response.get_json()["canAccessRestrictedXlsx"] is True
            assert app_client.get("/reports").status_code == 403
            assert app_client.post("/api/update-now").status_code == 403

    def test_admin_permissions_cannot_be_disabled(self, app_client):
        _set_role(app_client, "admin")
        settings = {
            "oidcPrivilegedPermissions": {
                "reports": False,
                "sync": False,
                "restrictedXlsx": False,
            }
        }
        with patch("simple_org_chart.app_main.load_settings", return_value=settings):
            response = app_client.get("/api/auth-check")
        assert response.get_json()["canAccessReports"] is True
        assert response.get_json()["canSync"] is True
        assert response.get_json()["canAccessRestrictedXlsx"] is True

    def test_restricted_xlsx_columns_follow_permission(self, app_client, tmp_path):
        data_file = tmp_path / "employee_data.json"
        data_file.write_text(
            json.dumps({"name": "Test User", "hireDate": "2024-01-01", "children": []}),
            encoding="utf-8",
        )
        settings = {
            "hideDisabledUsers": False,
            "hideGuestUsers": False,
            "hideNoTitle": False,
            "ignoredDepartments": "",
            "exportXlsxColumns": {
                "name": "show",
                "hireDate": "admin",
            },
            "oidcPrivilegedPermissions": {"restrictedXlsx": False},
        }
        _set_role(app_client, "privileged")

        with (
            patch("simple_org_chart.app_main.DATA_FILE", str(data_file)),
            patch("simple_org_chart.app_main.load_settings", return_value=settings),
        ):
            response = app_client.get("/api/export-xlsx")

        workbook = load_workbook(BytesIO(response.data))
        headers = [cell.value for cell in workbook.active[1]]
        assert "Name" in headers
        assert "Hire Date" not in headers

    def test_reader_cannot_trigger_sync(self, app_client):
        _set_role(app_client, "reader")
        assert app_client.post("/api/update-now").status_code == 403

    def test_privileged_user_cannot_export_admin_settings(self, app_client):
        _set_role(app_client, "privileged")
        assert app_client.get("/api/settings/export").status_code == 403

    def test_privileged_user_cannot_mutate_admin_settings(self, app_client):
        _set_role(app_client, "privileged")
        assert app_client.post("/api/settings", json={}).status_code == 403

    def test_oidc_callback_denies_user_outside_all_access_groups(self, app_client):
        with app_client.session_transaction() as user_session:
            user_session["oidc_flow"] = {"state": "test-state"}

        with (
            patch("simple_org_chart.app_main.AUTH_TYPE", "oidc"),
            patch(
                "simple_org_chart.app_main.complete_login",
                return_value={"access_token": "token"},
            ),
            patch(
                "simple_org_chart.app_main.fetch_user_and_groups",
                return_value=({"userPrincipalName": "user@example.com"}, set()),
            ),
        ):
            response = app_client.get("/auth/callback?state=test-state&code=code")

        assert response.status_code == 403
        with app_client.session_transaction() as user_session:
            assert "authenticated" not in user_session

    def test_oidc_mode_invalidates_session_from_old_access_rules(self, app_client):
        _set_role(app_client, "reader")

        with patch("simple_org_chart.app_main.AUTH_TYPE", "oidc"):
            response = app_client.get("/")

        assert response.status_code == 302
        assert response.headers["Location"].endswith("/login")
        with app_client.session_transaction() as user_session:
            assert "authenticated" not in user_session

    def test_oidc_config_page_shows_groups_without_logout(self, app_client):
        with app_client.session_transaction() as user_session:
            user_session["authenticated"] = True
            user_session["role"] = "admin"
            user_session["oidc_session_version"] = 3

        with patch("simple_org_chart.app_main.AUTH_TYPE", "oidc"):
            response = app_client.get("/configure")

        page = response.get_data(as_text=True)
        assert response.status_code == 200
        assert 'id="oidcReaderGroupId"' in page
        assert 'id="oidcPrivilegedGroupId"' in page
        assert 'id="oidcAdminGroupId"' in page
        assert 'data-config-action="logout"' not in page
