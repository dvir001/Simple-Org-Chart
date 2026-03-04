"""Tests for i18n locale consistency – ensure all data-i18n keys in templates have locale entries."""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOCALE_FILE = PROJECT_ROOT / "static" / "locales" / "en-US.json"
TEMPLATES_DIR = PROJECT_ROOT / "templates"
STATIC_DIR = PROJECT_ROOT / "static"

# Regex to capture i18n key attributes from HTML templates.
# Matches data-i18n and all data-i18n-* variants that hold keys
# (e.g. data-i18n-placeholder, data-i18n-title, data-i18n-html,
# data-i18n-aria-label, data-i18n-ariaLabel, data-i18n-alt).
# Excludes data-i18n-params which holds JSON parameter data, not a key.
_DATA_I18N_RE = re.compile(r'data-i18n(?:-(?!params)[a-zA-Z-]+)?="([^"]+)"')
# Regex to capture i18n key references in JS (translator calls like t('key'))
_JS_I18N_RE = re.compile(r"""(?:getTranslator\(\)|[^a-zA-Z]t)\(\s*['"]([a-zA-Z0-9_.]+)['"]\s*[,)]""")


def _load_locale() -> dict:
    with LOCALE_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


def _resolve_key(obj: dict, dotted_key: str):
    """Walk a nested dict by dotted key. Returns None if missing."""
    parts = dotted_key.split(".")
    current = obj
    for part in parts:
        if not isinstance(current, dict) or part not in current:
            return None
        current = current[part]
    return current


def _collect_html_i18n_keys() -> list[tuple[str, str]]:
    """Return (file, key) pairs for every data-i18n reference in templates."""
    results = []
    for html_file in sorted(TEMPLATES_DIR.glob("*.html")):
        text = html_file.read_text(encoding="utf-8")
        for match in _DATA_I18N_RE.finditer(text):
            key = match.group(1)
            # Skip keys with dynamic params expressions
            if "{" in key or "}" in key:
                continue
            results.append((html_file.name, key))
    return results


def _collect_js_i18n_keys() -> list[tuple[str, str]]:
    """Return (file, key) pairs for i18n references in JS files."""
    results = []
    for js_file in sorted(STATIC_DIR.glob("*.js")):
        if js_file.name in ("d3.min.js", "jspdf.umd.min.js", "i18n.js"):
            continue
        text = js_file.read_text(encoding="utf-8")
        for match in _JS_I18N_RE.finditer(text):
            key = match.group(1)
            results.append((js_file.name, key))
    return results


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


class TestLocaleFileStructure:
    def test_locale_file_exists(self):
        assert LOCALE_FILE.exists(), f"Locale file missing: {LOCALE_FILE}"

    def test_locale_file_valid_json(self):
        locale = _load_locale()
        assert isinstance(locale, dict)


class TestHTMLKeysExistInLocale:
    """Every data-i18n key referenced in templates should resolve in en-US.json."""

    @pytest.fixture(scope="class")
    def locale(self):
        return _load_locale()

    @pytest.fixture(scope="class")
    def html_keys(self):
        return _collect_html_i18n_keys()

    def test_has_html_keys(self, html_keys):
        assert len(html_keys) > 0, "No data-i18n keys found in templates"

    def test_all_html_keys_resolve(self, locale, html_keys):
        missing = []
        for filename, key in html_keys:
            if _resolve_key(locale, key) is None:
                missing.append(f"  {filename}: {key}")
        assert not missing, "HTML i18n keys missing from en-US.json:\n" + "\n".join(missing)


class TestJSKeysExistInLocale:
    """i18n keys referenced in JS files should resolve in en-US.json."""

    @pytest.fixture(scope="class")
    def locale(self):
        return _load_locale()

    @pytest.fixture(scope="class")
    def js_keys(self):
        return _collect_js_i18n_keys()

    def test_has_js_keys(self, js_keys):
        assert len(js_keys) > 0, "No i18n keys found in JS files"

    def test_all_js_keys_resolve(self, locale, js_keys):
        missing = []
        for filename, key in js_keys:
            if _resolve_key(locale, key) is None:
                missing.append(f"  {filename}: {key}")
        assert not missing, "JS i18n keys missing from en-US.json:\n" + "\n".join(missing)
