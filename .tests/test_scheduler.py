"""Tests for the background scheduler loop."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

from simple_org_chart import scheduler


class _TickLimitedEvent:
    """Stand-in for threading.Event that stops the loop after a fixed number of ticks."""

    def __init__(self, ticks: int) -> None:
        self._remaining = ticks

    def is_set(self) -> bool:
        return self._remaining <= 0

    def wait(self, _timeout: float) -> None:
        self._remaining -= 1


def _run_loop(monkeypatch, settings_per_tick, callback, run_initial="false"):
    """Run the loop once per settings entry; an 08:00 update time counts as already due."""
    remaining = list(settings_per_tick)

    def fake_compute_next_run(update_time, _tz):
        offset = timedelta(minutes=-1) if update_time.hour == 8 else timedelta(days=1)
        return datetime.now(timezone.utc) + offset

    monkeypatch.setenv("RUN_INITIAL_UPDATE", run_initial)
    monkeypatch.setattr(scheduler, "load_settings", lambda: remaining.pop(0))
    monkeypatch.setattr(scheduler, "_compute_next_run", fake_compute_next_run)
    monkeypatch.setattr(scheduler, "_update_callback", callback)
    scheduler._schedule_loop(_TickLimitedEvent(len(settings_per_tick)))


def test_loop_picks_up_update_time_saved_while_running(monkeypatch):
    calls = []
    _run_loop(
        monkeypatch,
        [{"updateTime": "20:00"}, {"updateTime": "08:00"}],
        lambda source: calls.append(source),
    )
    assert calls == ["scheduled"]


def test_loop_stops_scheduling_when_auto_update_is_disabled(monkeypatch):
    calls = []
    _run_loop(
        monkeypatch,
        [{"updateTime": "20:00"}, {"updateTime": "08:00", "autoUpdateEnabled": False}],
        lambda source: calls.append(source),
    )
    assert calls == []


def test_failed_startup_update_does_not_stop_daily_schedule(monkeypatch):
    calls = []

    def callback(source):
        calls.append(source)
        if source == "startup":
            raise RuntimeError("Graph unavailable")

    _run_loop(monkeypatch, [{"updateTime": "08:00"}], callback, run_initial="true")
    assert calls == ["startup", "scheduled"]
