"""Tests for simple_org_chart.hierarchy – tree building & missing-manager detection."""

from __future__ import annotations

from typing import Any, Dict, List

import pytest
from simple_org_chart.hierarchy import (
    build_org_hierarchy,
    collect_missing_manager_records,
    flatten_hierarchy_to_employee_list,
)
from simple_org_chart.settings import DEFAULT_SETTINGS


def _settings(**overrides) -> Dict[str, Any]:
    s = DEFAULT_SETTINGS.copy()
    s.update(overrides)
    return s


# ---------------------------------------------------------------------------
# build_org_hierarchy
# ---------------------------------------------------------------------------


class TestBuildOrgHierarchy:
    def test_empty_list(self):
        assert build_org_hierarchy([], settings=_settings()) is None

    def test_auto_detect_ceo(self, sample_employees):
        root = build_org_hierarchy(sample_employees, settings=_settings())
        assert root is not None
        assert root["name"] == "Alice CEO"

    def test_explicit_top_user(self, sample_employees):
        root = build_org_hierarchy(
            sample_employees,
            top_user_email_override="bob@example.com",
            settings=_settings(),
        )
        assert root is not None
        assert root["name"] == "Bob VP"

    def test_children_wired(self, sample_employees):
        root = build_org_hierarchy(sample_employees, settings=_settings())
        child_names = {c["name"] for c in root.get("children", [])}
        assert "Bob VP" in child_names

    def test_grandchildren(self, sample_employees):
        root = build_org_hierarchy(sample_employees, settings=_settings())
        bob = next(c for c in root["children"] if c["name"] == "Bob VP")
        grandchild_names = {c["name"] for c in bob.get("children", [])}
        assert "Carol Dev" in grandchild_names

    def test_fallback_to_settings_id(self, sample_employees):
        root = build_org_hierarchy(
            sample_employees,
            settings=_settings(topLevelUserId="2"),
        )
        assert root is not None
        assert root["id"] == "2"


# ---------------------------------------------------------------------------
# flatten_hierarchy_to_employee_list
# ---------------------------------------------------------------------------


class TestFlattenHierarchy:
    def test_flatten(self, sample_employees):
        root = build_org_hierarchy(sample_employees, settings=_settings())
        flat = flatten_hierarchy_to_employee_list(root)
        flat_ids = {e["id"] for e in flat}
        # Root + Bob + Carol should be in the tree (Dave is orphaned but may or
        # may not appear depending on auto-detect wiring)
        assert "1" in flat_ids
        assert "2" in flat_ids
        assert "3" in flat_ids

    def test_flatten_none(self):
        assert flatten_hierarchy_to_employee_list(None) == []


# ---------------------------------------------------------------------------
# collect_missing_manager_records
# ---------------------------------------------------------------------------


class TestMissingManagers:
    def test_orphan_detected(self, sample_employees):
        root = build_org_hierarchy(sample_employees, settings=_settings())
        missing = collect_missing_manager_records(
            sample_employees,
            hierarchy_root=root,
            settings=_settings(),
        )
        orphan_names = {r["name"] for r in missing}
        assert "Dave Orphan" in orphan_names

    def test_root_not_flagged(self, sample_employees):
        root = build_org_hierarchy(sample_employees, settings=_settings())
        missing = collect_missing_manager_records(
            sample_employees,
            hierarchy_root=root,
            settings=_settings(),
        )
        names = {r["name"] for r in missing}
        assert "Alice CEO" not in names

    def test_empty_input(self):
        assert collect_missing_manager_records([]) == []
