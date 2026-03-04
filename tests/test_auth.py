"""Tests for simple_org_chart.auth – decorators & path sanitisation."""

from __future__ import annotations

import pytest
from simple_org_chart.auth import sanitize_next_path


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
