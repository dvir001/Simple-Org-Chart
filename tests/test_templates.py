"""Tests for template & static asset consistency."""

from __future__ import annotations

import re
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = PROJECT_ROOT / "templates"
STATIC_DIR = PROJECT_ROOT / "static"

# Regex to match things like href="/static/foo.css" or src="/static/bar.js"
_STATIC_REF_RE = re.compile(r'(?:href|src)="/static/([^"]+)"')
# Regex to match Jinja2 url_for references: url_for('static', filename='foo.css')
_URL_FOR_STATIC_RE = re.compile(r"""url_for\(\s*['"]static['"]\s*,\s*filename\s*=\s*['"]([^'"]+)['"]\s*\)""")
# ID references from JS: qs('someId') or getElementById('someId')
_QS_RE = re.compile(r"""qs\(\s*['"]([^'"]+)['"]\s*\)""")
_GETBYID_RE = re.compile(r"""getElementById\(\s*['"]([^'"]+)['"]\s*\)""")


def _collect_html_ids(filename: str) -> set[str]:
    """Return all element IDs declared in a template."""
    path = TEMPLATES_DIR / filename
    if not path.exists():
        return set()
    text = path.read_text(encoding="utf-8")
    return set(re.findall(r'\bid="([^"]+)"', text))


# ---------------------------------------------------------------------------
# Static asset references
# ---------------------------------------------------------------------------


class TestStaticAssets:
    """Verify that CSS/JS files referenced in templates exist on disk."""

    def _collect_references(self) -> list[tuple[str, str]]:
        refs = []
        patterns = [_STATIC_REF_RE, _URL_FOR_STATIC_RE]
        for html_file in TEMPLATES_DIR.glob("*.html"):
            text = html_file.read_text(encoding="utf-8")
            for pattern in patterns:
                for match in pattern.finditer(text):
                    refs.append((html_file.name, match.group(1)))
        return refs

    def test_all_static_refs_exist(self):
        missing = []
        for template, asset in self._collect_references():
            if not (STATIC_DIR / asset).exists():
                missing.append(f"  {template} references missing /static/{asset}")
        assert not missing, "Missing static assets:\n" + "\n".join(missing)


# ---------------------------------------------------------------------------
# Template existence
# ---------------------------------------------------------------------------


class TestTemplates:
    @pytest.mark.parametrize(
        "template",
        ["index.html", "login.html", "configure.html", "reports.html"],
    )
    def test_template_exists(self, template):
        assert (TEMPLATES_DIR / template).exists()


# ---------------------------------------------------------------------------
# reports.html ↔ reports.js DOM ID consistency
# ---------------------------------------------------------------------------


class TestReportsDOMConsistency:
    """Ensure element IDs referenced by reports.js exist in reports.html."""

    @pytest.fixture(scope="class")
    def html_ids(self) -> set[str]:
        return _collect_html_ids("reports.html")

    @pytest.fixture(scope="class")
    def js_ids(self) -> set[str]:
        js_path = STATIC_DIR / "reports.js"
        text = js_path.read_text(encoding="utf-8")
        ids: set[str] = set()
        for match in _QS_RE.finditer(text):
            ids.add(match.group(1))
        for match in _GETBYID_RE.finditer(text):
            ids.add(match.group(1))
        return ids

    def test_critical_ids_in_html(self, html_ids):
        """Spot-check critical IDs that must exist."""
        critical = {
            "reportTypeSelect",
            "userScannerPanel",
            "reportTableHead",
            "reportTableBody",
            "tableTitle",
            "tableStatus",
            "runUserScanBtn",
            "runFullScanBtn",
            "stopFullScanBtn",
            "fullScanStatus",
            "fullScanProgressWrap",
            "fullScanProgressFill",
            "fullScanProgressLabel",
            "fullScanTerminal",
            "fullScanTerminalBody",
            "fullScanHistory",
        }
        missing = critical - html_ids
        assert not missing, f"Critical IDs missing from reports.html: {missing}"

    def test_js_ids_mostly_exist_in_html(self, html_ids, js_ids):
        """Most IDs referenced by JS should exist in the HTML.
        
        We allow a small tolerance for dynamically-created IDs.
        """
        # Filter out IDs that look dynamically built (contain variable-like patterns)
        static_js_ids = {i for i in js_ids if not re.search(r"\$\{|\\|\/", i)}
        missing = static_js_ids - html_ids
        # Allow up to 15% dynamic IDs (generous tolerance)
        threshold = max(5, int(len(static_js_ids) * 0.15))
        assert len(missing) <= threshold, (
            f"Too many JS DOM IDs missing from reports.html ({len(missing)}/{len(static_js_ids)}):\n"
            + "\n".join(f"  {i}" for i in sorted(missing))
        )
