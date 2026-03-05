"""Tests for email scheduling logic in email_config.py."""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timedelta, timezone
from unittest.mock import patch

import pytest

# Ensure the package is importable
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
os.environ.setdefault("ADMIN_PASSWORD", "test-password-for-ci-only")

from simple_org_chart.email_config import (
    DEFAULT_EMAIL_CONFIG,
    _already_sent_this_period,
    _is_scheduled_day,
    load_email_config,
    mark_email_sent,
    save_email_config,
    should_send_email_now,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _utc(year, month, day, hour=12, minute=0):
    """Shorthand for a UTC datetime."""
    return datetime(year, month, day, hour, minute, tzinfo=timezone.utc)


# ---------------------------------------------------------------------------
# _is_scheduled_day
# ---------------------------------------------------------------------------

class TestIsScheduledDay:
    """Verify that _is_scheduled_day correctly matches calendar days."""

    def test_daily_always_true(self):
        assert _is_scheduled_day(_utc(2026, 3, 5), "daily", "monday", "first") is True
        assert _is_scheduled_day(_utc(2026, 3, 7), "daily", "friday", "last") is True

    def test_weekly_correct_day(self):
        # 2026-03-02 is a Monday
        assert _is_scheduled_day(_utc(2026, 3, 2), "weekly", "monday", "first") is True

    def test_weekly_wrong_day(self):
        # 2026-03-03 is a Tuesday
        assert _is_scheduled_day(_utc(2026, 3, 3), "weekly", "monday", "first") is False

    def test_monthly_first(self):
        assert _is_scheduled_day(_utc(2026, 4, 1), "monthly", "monday", "first") is True
        assert _is_scheduled_day(_utc(2026, 4, 2), "monthly", "monday", "first") is False

    def test_monthly_last_day(self):
        # March has 31 days → 31st is last
        assert _is_scheduled_day(_utc(2026, 3, 31), "monthly", "monday", "last") is True
        assert _is_scheduled_day(_utc(2026, 3, 30), "monthly", "monday", "last") is False

    def test_monthly_last_feb(self):
        # February 2026 has 28 days
        assert _is_scheduled_day(_utc(2026, 2, 28), "monthly", "monday", "last") is True
        assert _is_scheduled_day(_utc(2026, 2, 27), "monthly", "monday", "last") is False

    def test_unknown_frequency(self):
        assert _is_scheduled_day(_utc(2026, 3, 5), "biweekly", "monday", "first") is False


# ---------------------------------------------------------------------------
# _already_sent_this_period
# ---------------------------------------------------------------------------

class TestAlreadySentThisPeriod:
    """Verify period-elapsed detection."""

    def test_daily_same_day(self):
        now = _utc(2026, 3, 5, 14, 0)
        last = _utc(2026, 3, 5, 8, 0)
        assert _already_sent_this_period(now, last, "daily") is True

    def test_daily_previous_day(self):
        now = _utc(2026, 3, 5, 8, 0)
        last = _utc(2026, 3, 4, 20, 0)
        assert _already_sent_this_period(now, last, "daily") is False

    def test_weekly_within_window(self):
        now = _utc(2026, 3, 9)  # Monday
        last = _utc(2026, 3, 5)  # 4 days ago
        assert _already_sent_this_period(now, last, "weekly") is True

    def test_weekly_outside_window(self):
        now = _utc(2026, 3, 9)  # Monday
        last = _utc(2026, 3, 1)  # 8 days ago
        assert _already_sent_this_period(now, last, "weekly") is False

    def test_monthly_same_month(self):
        now = _utc(2026, 4, 1)
        last = _utc(2026, 4, 1, 2, 0)  # earlier same day
        assert _already_sent_this_period(now, last, "monthly") is True

    def test_monthly_previous_month(self):
        now = _utc(2026, 4, 1)
        last = _utc(2026, 3, 1)
        assert _already_sent_this_period(now, last, "monthly") is False

    def test_monthly_previous_year(self):
        now = _utc(2026, 1, 1)
        last = _utc(2025, 12, 1)
        assert _already_sent_this_period(now, last, "monthly") is False


# ---------------------------------------------------------------------------
# should_send_email_now  (integration-level, mocks load_email_config & datetime)
# ---------------------------------------------------------------------------

def _mock_config(**overrides):
    cfg = DEFAULT_EMAIL_CONFIG.copy()
    cfg["enabled"] = True
    cfg["recipientEmail"] = "test@example.com"
    cfg.update(overrides)
    return cfg


class TestShouldSendEmailNow:
    """End-to-end tests for should_send_email_now."""

    @patch("simple_org_chart.email_config.load_email_config")
    def test_disabled(self, mock_load):
        mock_load.return_value = _mock_config(enabled=False)
        assert should_send_email_now() is False

    @patch("simple_org_chart.email_config.datetime")
    @patch("simple_org_chart.email_config.load_email_config")
    def test_monthly_first_never_sent_on_first(self, mock_load, mock_dt):
        mock_dt.now.return_value = _utc(2026, 4, 1)
        mock_dt.fromisoformat = datetime.fromisoformat
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)
        mock_load.return_value = _mock_config(frequency="monthly", dayOfMonth="first", lastSent=None)
        assert should_send_email_now() is True

    @patch("simple_org_chart.email_config.datetime")
    @patch("simple_org_chart.email_config.load_email_config")
    def test_monthly_first_never_sent_wrong_day(self, mock_load, mock_dt):
        """Even if never sent, don't fire on the 15th for a monthly-first schedule."""
        mock_dt.now.return_value = _utc(2026, 4, 15)
        mock_dt.fromisoformat = datetime.fromisoformat
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)
        mock_load.return_value = _mock_config(frequency="monthly", dayOfMonth="first", lastSent=None)
        assert should_send_email_now() is False

    @patch("simple_org_chart.email_config.datetime")
    @patch("simple_org_chart.email_config.load_email_config")
    def test_monthly_first_already_sent_this_month(self, mock_load, mock_dt):
        mock_dt.now.return_value = _utc(2026, 4, 1, 20, 0)
        mock_dt.fromisoformat = datetime.fromisoformat
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)
        mock_load.return_value = _mock_config(
            frequency="monthly", dayOfMonth="first",
            lastSent=_utc(2026, 4, 1, 8, 0).isoformat(),
        )
        assert should_send_email_now() is False

    @patch("simple_org_chart.email_config.datetime")
    @patch("simple_org_chart.email_config.load_email_config")
    def test_monthly_first_new_month(self, mock_load, mock_dt):
        mock_dt.now.return_value = _utc(2026, 5, 1)
        mock_dt.fromisoformat = datetime.fromisoformat
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)
        mock_load.return_value = _mock_config(
            frequency="monthly", dayOfMonth="first",
            lastSent=_utc(2026, 4, 1).isoformat(),
        )
        assert should_send_email_now() is True

    @patch("simple_org_chart.email_config.datetime")
    @patch("simple_org_chart.email_config.load_email_config")
    def test_daily_sends_every_day(self, mock_load, mock_dt):
        """Daily should send when last sent was yesterday."""
        mock_dt.now.return_value = _utc(2026, 3, 5)
        mock_dt.fromisoformat = datetime.fromisoformat
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)
        mock_load.return_value = _mock_config(
            frequency="daily",
            lastSent=_utc(2026, 3, 4).isoformat(),
        )
        assert should_send_email_now() is True

    @patch("simple_org_chart.email_config.datetime")
    @patch("simple_org_chart.email_config.load_email_config")
    def test_daily_no_double_send(self, mock_load, mock_dt):
        mock_dt.now.return_value = _utc(2026, 3, 5, 20, 0)
        mock_dt.fromisoformat = datetime.fromisoformat
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)
        mock_load.return_value = _mock_config(
            frequency="daily",
            lastSent=_utc(2026, 3, 5, 8, 0).isoformat(),
        )
        assert should_send_email_now() is False

    @patch("simple_org_chart.email_config.datetime")
    @patch("simple_org_chart.email_config.load_email_config")
    def test_weekly_correct_day_period_elapsed(self, mock_load, mock_dt):
        # 2026-03-09 is a Monday
        mock_dt.now.return_value = _utc(2026, 3, 9)
        mock_dt.fromisoformat = datetime.fromisoformat
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)
        mock_load.return_value = _mock_config(
            frequency="weekly", dayOfWeek="monday",
            lastSent=_utc(2026, 3, 2).isoformat(),
        )
        assert should_send_email_now() is True

    @patch("simple_org_chart.email_config.datetime")
    @patch("simple_org_chart.email_config.load_email_config")
    def test_weekly_wrong_day(self, mock_load, mock_dt):
        # 2026-03-10 is a Tuesday
        mock_dt.now.return_value = _utc(2026, 3, 10)
        mock_dt.fromisoformat = datetime.fromisoformat
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)
        mock_load.return_value = _mock_config(
            frequency="weekly", dayOfWeek="monday",
            lastSent=_utc(2026, 3, 2).isoformat(),
        )
        assert should_send_email_now() is False

    @patch("simple_org_chart.email_config.datetime")
    @patch("simple_org_chart.email_config.load_email_config")
    def test_invalid_last_sent_on_schedule_day(self, mock_load, mock_dt):
        mock_dt.now.return_value = _utc(2026, 4, 1)
        mock_dt.fromisoformat = datetime.fromisoformat
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)
        mock_load.return_value = _mock_config(
            frequency="monthly", dayOfMonth="first",
            lastSent="not-a-date",
        )
        assert should_send_email_now() is True


# ---------------------------------------------------------------------------
# save_email_config preserves lastSent
# ---------------------------------------------------------------------------

class TestSavePreservesLastSent:
    """Ensure saving from the configure page doesn't wipe lastSent."""

    def test_last_sent_preserved_on_save(self, tmp_path, monkeypatch):
        config_file = tmp_path / "email_config.json"
        monkeypatch.setattr("simple_org_chart.email_config.EMAIL_CONFIG_FILE", config_file)

        # Simulate a previous email send
        initial = DEFAULT_EMAIL_CONFIG.copy()
        initial["enabled"] = True
        initial["lastSent"] = "2026-03-01T12:00:00+00:00"
        config_file.write_text(json.dumps(initial))

        # Simulate a configure-page save (no lastSent in payload)
        save_email_config({"enabled": True, "frequency": "monthly", "dayOfMonth": "first"})

        reloaded = json.loads(config_file.read_text())
        assert reloaded["lastSent"] == "2026-03-01T12:00:00+00:00"

    def test_last_sent_overwritten_when_explicitly_provided(self, tmp_path, monkeypatch):
        config_file = tmp_path / "email_config.json"
        monkeypatch.setattr("simple_org_chart.email_config.EMAIL_CONFIG_FILE", config_file)

        initial = DEFAULT_EMAIL_CONFIG.copy()
        initial["lastSent"] = "2026-03-01T12:00:00+00:00"
        config_file.write_text(json.dumps(initial))

        # mark_email_sent passes lastSent explicitly
        save_email_config({"lastSent": "2026-04-01T08:00:00+00:00"})

        reloaded = json.loads(config_file.read_text())
        assert reloaded["lastSent"] == "2026-04-01T08:00:00+00:00"

    def test_last_sent_none_when_no_prior_config(self, tmp_path, monkeypatch):
        config_file = tmp_path / "email_config.json"
        monkeypatch.setattr("simple_org_chart.email_config.EMAIL_CONFIG_FILE", config_file)

        save_email_config({"enabled": True, "frequency": "weekly"})

        reloaded = json.loads(config_file.read_text())
        assert reloaded["lastSent"] is None
