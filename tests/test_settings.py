"""Tests for simple_org_chart.settings – parsing, normalisation, persistence."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict
from unittest.mock import patch

import pytest
from simple_org_chart.settings import (
    DEFAULT_SETTINGS,
    department_is_ignored,
    employee_is_ignored,
    load_settings,
    normalize_filter_value,
    parse_filter_values,
    parse_ignored_departments,
    parse_ignored_employees,
    parse_ignored_titles,
    save_settings,
    translate_placeholder,
)


# ---------------------------------------------------------------------------
# normalize_filter_value
# ---------------------------------------------------------------------------


class TestNormalizeFilterValue:
    @pytest.mark.parametrize(
        "raw, expected",
        [
            ("  Hello  ", "hello"),
            ("  -—Leading   ", "leading"),
            ("", ""),
            (None, ""),
            ("Finance", "finance"),
            ("  IT  Support  ", "it support"),
            ("--dashes--", "dashes"),
        ],
    )
    def test_normalise(self, raw, expected):
        assert normalize_filter_value(raw) == expected


# ---------------------------------------------------------------------------
# parse_filter_values
# ---------------------------------------------------------------------------


class TestParseFilterValues:
    def test_comma_separated(self):
        result = parse_filter_values("HR, Finance, IT")
        assert result == {"hr", "finance", "it"}

    def test_semicolon_separated(self):
        result = parse_filter_values("HR; Finance; IT")
        assert result == {"hr", "finance", "it"}

    def test_json_array(self):
        result = parse_filter_values('["HR", "Finance"]')
        assert result == {"hr", "finance"}

    def test_empty_string(self):
        assert parse_filter_values("") == set()

    def test_none(self):
        assert parse_filter_values(None) == set()

    def test_list_input(self):
        assert parse_filter_values(["One", "Two"]) == {"one", "two"}

    def test_set_input(self):
        assert parse_filter_values({"AlReady"}) == {"already"}


# ---------------------------------------------------------------------------
# parse_ignored_* helpers
# ---------------------------------------------------------------------------


class TestParseIgnored:
    def test_departments(self):
        settings = {"ignoredDepartments": "HR, Finance"}
        assert parse_ignored_departments(settings) == {"hr", "finance"}

    def test_titles(self):
        settings = {"ignoredTitles": "Intern; Contractor"}
        assert parse_ignored_titles(settings) == {"intern", "contractor"}

    def test_employees(self):
        settings = {"ignoredEmployees": "john.doe@example.com"}
        assert parse_ignored_employees(settings) == {"john.doe@example.com"}

    def test_missing_key_returns_empty(self):
        assert parse_ignored_departments({}) == set()


# ---------------------------------------------------------------------------
# department_is_ignored / employee_is_ignored
# ---------------------------------------------------------------------------


class TestIgnoreChecks:
    def test_department_in_set(self):
        assert department_is_ignored("Finance", {"finance", "hr"}) is True

    def test_department_not_in_set(self):
        assert department_is_ignored("Engineering", {"finance", "hr"}) is False

    def test_empty_set_never_ignored(self):
        assert department_is_ignored("Finance", set()) is False

    def test_employee_by_email(self):
        ignored = parse_filter_values("alice@example.com")
        assert employee_is_ignored("Alice", "alice@example.com", None, ignored) is True

    def test_employee_not_ignored(self):
        ignored = parse_filter_values("bob@example.com")
        assert employee_is_ignored("Alice", "alice@example.com", None, ignored) is False

    def test_employee_by_combo(self):
        ignored = parse_filter_values("alice <alice@example.com>")
        assert employee_is_ignored("Alice", "alice@example.com", None, ignored) is True


# ---------------------------------------------------------------------------
# translate_placeholder
# ---------------------------------------------------------------------------


class TestTranslatePlaceholder:
    def test_with_default(self):
        assert translate_placeholder("key", "Hello {name}", name="World") == "Hello World"

    def test_without_default(self):
        assert translate_placeholder("some.key") == "some.key"


# ---------------------------------------------------------------------------
# load_settings / save_settings (with tmp dir)
# ---------------------------------------------------------------------------


class TestSettingsPersistence:
    def test_load_defaults_when_missing(self, tmp_path: Path):
        fake_settings_file = tmp_path / "app_settings.json"
        with patch("simple_org_chart.settings.SETTINGS_FILE", fake_settings_file):
            result = load_settings()
        # Should contain all default keys
        for key in DEFAULT_SETTINGS:
            assert key in result

    def test_roundtrip(self, tmp_path: Path):
        fake_settings_file = tmp_path / "app_settings.json"
        custom = {"chartTitle": "My Org", "headerColor": "#FF0000"}
        with patch("simple_org_chart.settings.SETTINGS_FILE", fake_settings_file):
            assert save_settings(custom) is True
            loaded = load_settings()
        assert loaded["chartTitle"] == "My Org"
        assert loaded["headerColor"] == "#FF0000"

    def test_invalid_json_falls_back(self, tmp_path: Path):
        fake_settings_file = tmp_path / "app_settings.json"
        fake_settings_file.write_text("NOT JSON", encoding="utf-8")
        with patch("simple_org_chart.settings.SETTINGS_FILE", fake_settings_file):
            result = load_settings()
        # Should fall back to defaults
        assert result["chartTitle"] == DEFAULT_SETTINGS["chartTitle"]

    def test_hex_color_normalisation(self, tmp_path: Path):
        fake_settings_file = tmp_path / "app_settings.json"
        custom = {"headerColor": "ff0000"}
        with patch("simple_org_chart.settings.SETTINGS_FILE", fake_settings_file):
            save_settings(custom)
            loaded = load_settings()
        assert loaded["headerColor"] == "#FF0000"
