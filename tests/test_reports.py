"""Tests for simple_org_chart.reports – filter logic & cache manager."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List
from unittest.mock import patch

import pytest
from simple_org_chart.reports import (
    ReportCacheManager,
    apply_disabled_filters,
    apply_filtered_user_filters,
    apply_last_login_filters,
    apply_missing_manager_filters,
    calculate_license_totals,
)


# ---------------------------------------------------------------------------
# ReportCacheManager
# ---------------------------------------------------------------------------


class TestReportCacheManager:
    def test_load_missing_file(self, tmp_path: Path):
        cache = ReportCacheManager()
        result = cache.load_json(
            str(tmp_path / "nonexistent.json"),
            description="test",
        )
        assert result == []

    def test_load_valid_json(self, tmp_path: Path):
        data = [{"name": "Alice"}, {"name": "Bob"}]
        file_path = tmp_path / "test.json"
        file_path.write_text(json.dumps(data), encoding="utf-8")

        cache = ReportCacheManager()
        result = cache.load_json(str(file_path), description="test")
        assert len(result) == 2
        assert result[0]["name"] == "Alice"

    def test_load_malformed_json(self, tmp_path: Path):
        file_path = tmp_path / "bad.json"
        file_path.write_text("NOT VALID JSON", encoding="utf-8")

        cache = ReportCacheManager()
        result = cache.load_json(str(file_path), description="test")
        assert result == []

    def test_load_wrong_type(self, tmp_path: Path):
        file_path = tmp_path / "dict.json"
        file_path.write_text('{"key": "value"}', encoding="utf-8")

        cache = ReportCacheManager()
        result = cache.load_json(str(file_path), description="test", expected_type=list)
        assert result == []

    def test_refresh_callback_invoked(self, tmp_path: Path):
        data = [{"id": 1}]
        file_path = tmp_path / "data.json"

        def write_data():
            file_path.write_text(json.dumps(data), encoding="utf-8")

        cache = ReportCacheManager(refresh_callback=write_data)
        result = cache.load_json(str(file_path), refresh=True, description="test")
        assert result == data


# ---------------------------------------------------------------------------
# calculate_license_totals
# ---------------------------------------------------------------------------


class TestCalculateLicenseTotals:
    def test_sum(self):
        records = [
            {"licenseCount": 3},
            {"licenseCount": 5},
            {"licenseCount": None},
            {"licenseCount": 0},
        ]
        assert calculate_license_totals(records) == 8

    def test_empty(self):
        assert calculate_license_totals([]) == 0

    def test_none(self):
        assert calculate_license_totals(None) == 0


# ---------------------------------------------------------------------------
# apply_last_login_filters
# ---------------------------------------------------------------------------


class TestApplyLastLoginFilters:
    def test_default_includes_all_enabled_members(self, sample_login_records):
        result = apply_last_login_filters(sample_login_records)
        assert len(result) >= 1

    def test_exclude_disabled(self, sample_login_records):
        result = apply_last_login_filters(
            sample_login_records,
            include_disabled=False,
        )
        for r in result:
            assert r["accountEnabled"] is True

    def test_exclude_guests(self, sample_login_records):
        result = apply_last_login_filters(
            sample_login_records,
            include_guests=False,
        )
        for r in result:
            assert (r.get("userType") or "").lower() != "guest"

    def test_exclude_shared_mailboxes(self, sample_login_records):
        result = apply_last_login_filters(
            sample_login_records,
            include_shared_mailboxes=False,
        )
        for r in result:
            assert r.get("mailboxType") != "shared"

    def test_exclude_room_mailboxes(self, sample_login_records):
        result = apply_last_login_filters(
            sample_login_records,
            include_room_equipment_mailboxes=False,
        )
        for r in result:
            assert r.get("mailboxType") != "room"

    def test_inactive_days_threshold(self, sample_login_records):
        result = apply_last_login_filters(
            sample_login_records,
            inactive_days="90",
        )
        for r in result:
            days = r.get("daysSinceLastActivity")
            if days is not None:
                assert days >= 90

    def test_inactive_days_max(self, sample_login_records):
        result = apply_last_login_filters(
            sample_login_records,
            inactive_days_max="30",
        )
        for r in result:
            days = r.get("daysSinceLastActivity")
            if days is not None:
                assert days <= 30

    def test_never_signed_in(self, sample_login_records):
        result = apply_last_login_filters(
            sample_login_records,
            inactive_days="never",
        )
        for r in result:
            assert r.get("neverSignedIn") is True

    def test_exclude_unlicensed(self, sample_login_records):
        result = apply_last_login_filters(
            sample_login_records,
            include_unlicensed=False,
        )
        for r in result:
            assert (r.get("licenseCount") or 0) > 0

    def test_exclude_licensed(self, sample_login_records):
        result = apply_last_login_filters(
            sample_login_records,
            include_licensed=False,
        )
        for r in result:
            assert (r.get("licenseCount") or 0) == 0

    def test_none_records(self):
        assert apply_last_login_filters(None) == []

    def test_empty_records(self):
        assert apply_last_login_filters([]) == []


# ---------------------------------------------------------------------------
# apply_filtered_user_filters
# ---------------------------------------------------------------------------


class TestApplyFilteredUserFilters:
    def test_exclude_guests(self, sample_login_records):
        result = apply_filtered_user_filters(
            sample_login_records,
            include_guests=False,
        )
        for r in result:
            assert (r.get("userType") or "").lower() != "guest"

    def test_exclude_disabled(self, sample_login_records):
        result = apply_filtered_user_filters(
            sample_login_records,
            include_disabled=False,
        )
        for r in result:
            assert r["accountEnabled"] is True

    def test_none_returns_empty(self):
        assert apply_filtered_user_filters(None) == []


# ---------------------------------------------------------------------------
# apply_missing_manager_filters (delegates to apply_filtered_user_filters)
# ---------------------------------------------------------------------------


class TestApplyMissingManagerFilters:
    def test_delegates(self, sample_login_records):
        a = apply_missing_manager_filters(
            sample_login_records,
            include_guests=False,
        )
        b = apply_filtered_user_filters(
            sample_login_records,
            include_guests=False,
        )
        assert a == b


# ---------------------------------------------------------------------------
# apply_disabled_filters
# ---------------------------------------------------------------------------


class TestApplyDisabledFilters:
    def test_none(self):
        assert apply_disabled_filters(None) == []

    def test_empty(self):
        assert apply_disabled_filters([]) == []

    def test_licensed_only(self):
        records = [
            {"licenseCount": 1, "userType": "Member"},
            {"licenseCount": 0, "userType": "Member"},
        ]
        result = apply_disabled_filters(records, licensed_only=True)
        assert len(result) == 1
        assert result[0]["licenseCount"] == 1

    def test_exclude_guests(self):
        records = [
            {"licenseCount": 1, "userType": "Guest"},
            {"licenseCount": 1, "userType": "Member"},
        ]
        result = apply_disabled_filters(records, include_guests=False)
        assert len(result) == 1
        assert result[0]["userType"] == "Member"
