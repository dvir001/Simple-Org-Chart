"""Tests for reports.js REPORT_CONFIGS consistency.

Parses the JS config map and verifies that every referenced i18n key
exists in the locale file, and every dataPath has a matching route.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
REPORTS_JS = PROJECT_ROOT / "static" / "reports.js"
LOCALE_FILE = PROJECT_ROOT / "static" / "locales" / "en-US.json"
APP_MAIN = PROJECT_ROOT / "simple_org_chart" / "app_main.py"


def _load_locale() -> dict:
    with LOCALE_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


def _resolve_key(obj: dict, dotted_key: str):
    parts = dotted_key.split(".")
    current = obj
    for part in parts:
        if not isinstance(current, dict) or part not in current:
            return None
        current = current[part]
    return current


def _extract_report_config_keys() -> list[str]:
    """Extract i18n key strings from REPORT_CONFIGS in reports.js."""
    text = REPORTS_JS.read_text(encoding="utf-8")

    # Find the REPORT_CONFIGS object block
    start = text.find("const REPORT_CONFIGS")
    if start == -1:
        return []

    # Find all quoted strings that look like i18n keys.
    # This covers top-level keys, column labelKeys, filter group labelKeys,
    # filter option labelKeys, and any additional summary label keys.
    keys = []
    config_pattern = re.compile(
        r"""(?:labelKey|groupLabelKey|summaryLabelKey|licenseSummaryLabelKey|tableTitleKey|emptyKey|countSummaryKey)\s*:\s*['"]([^'"]+)['"]"""
    )
    for match in config_pattern.finditer(text, start):
        keys.append(match.group(1))
    return keys


def _extract_data_paths() -> list[str]:
    """Extract API dataPath values from REPORT_CONFIGS."""
    text = REPORTS_JS.read_text(encoding="utf-8")
    paths = []
    for match in re.finditer(r"""dataPath\s*:\s*['"]([^'"]+)['"]""", text):
        paths.append(match.group(1))
    return paths


def _extract_flask_routes() -> set[str]:
    """Extract route paths from app_main.py."""
    text = APP_MAIN.read_text(encoding="utf-8")
    routes = set()
    for match in re.finditer(r"""@app\.route\(\s*['"]([^'"]+)['"]""", text):
        routes.add(match.group(1))
    return routes


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


class TestReportConfigI18n:
    """Every i18n key in REPORT_CONFIGS should exist in the locale file."""

    @pytest.fixture(scope="class")
    def locale(self):
        return _load_locale()

    @pytest.fixture(scope="class")
    def config_keys(self):
        return _extract_report_config_keys()

    def test_has_config_keys(self, config_keys):
        assert len(config_keys) > 0, "Could not extract any i18n keys from REPORT_CONFIGS"

    def test_all_config_keys_resolve(self, locale, config_keys):
        missing = [k for k in config_keys if _resolve_key(locale, k) is None]
        assert not missing, "REPORT_CONFIGS i18n keys missing from locale:\n" + "\n".join(f"  {k}" for k in missing)


class TestReportConfigRoutes:
    """Every dataPath in REPORT_CONFIGS should correspond to a Flask route."""

    @pytest.fixture(scope="class")
    def data_paths(self):
        return _extract_data_paths()

    @pytest.fixture(scope="class")
    def flask_routes(self):
        return _extract_flask_routes()

    def test_has_data_paths(self, data_paths):
        assert len(data_paths) > 0

    def test_data_paths_have_routes(self, data_paths, flask_routes):
        missing = [p for p in data_paths if p not in flask_routes]
        assert not missing, "REPORT_CONFIGS dataPaths with no Flask route:\n" + "\n".join(f"  {p}" for p in missing)
