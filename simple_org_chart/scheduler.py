"""Background scheduler management for SimpleOrgChart."""

from __future__ import annotations

import logging
import os
import threading
import time
from datetime import datetime, timedelta, time as dt_time, timezone
from typing import Callable, Optional

import schedule

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover - Python < 3.9 not officially supported but guard anyway
    ZoneInfo = None  # type: ignore[assignment]

from simple_org_chart.settings import load_settings

logger = logging.getLogger(__name__)

_scheduler_running = False
_scheduler_lock = threading.Lock()
_scheduler_thread: Optional[threading.Thread] = None
_update_callback: Optional[Callable[[], None]] = None

DEFAULT_TIME_STRING = "20:00"
DEFAULT_TIMEZONE = "UTC"


def _resolve_timezone(tz_name: Optional[str]) -> timezone:
    if ZoneInfo is None:
        logger.warning("ZoneInfo module unavailable; falling back to server local timezone")
        return datetime.now().astimezone().tzinfo or timezone.utc

    if tz_name:
        try:
            return ZoneInfo(tz_name)
        except Exception:  # noqa: BLE001 - log and fall back
            logger.warning("Invalid timezone '%s'; defaulting to UTC", tz_name)
    return ZoneInfo(DEFAULT_TIMEZONE)


def _parse_time_string(value: Optional[str]) -> dt_time:
    candidate = (value or DEFAULT_TIME_STRING).strip()
    try:
        hour_str, minute_str = candidate.split(":", 1)
        hour = max(0, min(23, int(hour_str)))
        minute = max(0, min(59, int(minute_str)))
        return dt_time(hour=hour, minute=minute)
    except Exception:
        logger.warning("Invalid update time '%s'; defaulting to %s", value, DEFAULT_TIME_STRING)
        return dt_time(hour=20, minute=0)


def _compute_next_run(update_time: dt_time, tz: timezone) -> datetime:
    tz_now = datetime.now(tz)
    candidate = tz_now.replace(hour=update_time.hour, minute=update_time.minute, second=0, microsecond=0)
    if candidate <= tz_now:
        candidate += timedelta(days=1)
    return candidate.astimezone(timezone.utc)


def configure_scheduler(update_callback: Callable[[], None]) -> None:
    """Register the callback used to refresh employee data."""
    global _update_callback
    _update_callback = update_callback


def is_scheduler_running() -> bool:
    """Return True if the background scheduler loop is active."""
    return _scheduler_running


def _ensure_callback() -> Callable[[], None]:
    if _update_callback is None:
        raise RuntimeError("Scheduler update callback has not been configured")
    return _update_callback


def _schedule_loop() -> None:
    global _scheduler_running

    try:
        update_callback = _ensure_callback()
    except RuntimeError as exc:
        logger.error(str(exc))
        _scheduler_running = False
        return

    schedule.clear()
    settings = load_settings()

    run_initial = os.environ.get("RUN_INITIAL_UPDATE", "auto").lower()
    if run_initial == "true":
        logger.info("[%s] Running initial employee data update on startup (RUN_INITIAL_UPDATE=true)...", datetime.now())
        update_callback(source='startup')
    elif run_initial == "auto":
        # Only run initial update if no cached data exists
        from simple_org_chart.config import DATA_FILE
        if not os.path.exists(DATA_FILE):
            logger.info("[%s] No cached data found; running initial employee data update...", datetime.now())
            update_callback(source='startup-no-cache')
        else:
            logger.info("[%s] Cached data exists; skipping initial update (set RUN_INITIAL_UPDATE=true to force)", datetime.now())
    else:
        logger.info("[%s] Initial update skipped (RUN_INITIAL_UPDATE=%s)", datetime.now(), run_initial)

    update_time = _parse_time_string(settings.get("updateTime"))
    tz = timezone.utc
    next_run_utc: Optional[datetime]
    if settings.get("autoUpdateEnabled", True):
        next_run_utc = _compute_next_run(update_time, tz)
        logger.info(
            "Scheduled daily updates for %s UTC; next run at %s",
            update_time.strftime("%H:%M"),
            next_run_utc.strftime("%Y-%m-%d %H:%M UTC"),
        )
    else:
        next_run_utc = None
        logger.info("Automatic updates are disabled; skipping daily schedule")

    while _scheduler_running:
        if next_run_utc is not None and datetime.now(timezone.utc) >= next_run_utc:
            logger.info("Executing scheduled update at %s UTC", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"))
            try:
                update_callback()
            except Exception as exc:  # noqa: BLE001 - log and continue running loop
                logger.exception("Scheduled update callback failed: %s", exc)

            settings = load_settings()
            update_time = _parse_time_string(settings.get("updateTime"))
            tz = timezone.utc

            if settings.get("autoUpdateEnabled", True):
                next_run_utc = _compute_next_run(update_time, tz)
                logger.info(
                    "Next scheduled update at %s UTC",
                    next_run_utc.strftime("%Y-%m-%d %H:%M"),
                )
            else:
                next_run_utc = None
                logger.info("Automatic updates disabled; halting further scheduling")

        time.sleep(30)


def start_scheduler() -> None:
    """Start the background scheduler thread if it is not already running."""
    global _scheduler_running, _scheduler_thread

    with _scheduler_lock:
        if _scheduler_running:
            return
        _scheduler_running = True
        _scheduler_thread = threading.Thread(target=_schedule_loop, daemon=True)
        _scheduler_thread.start()
        logger.info("Scheduler started")


def stop_scheduler() -> None:
    """Stop the background scheduler loop."""
    global _scheduler_running

    with _scheduler_lock:
        if not _scheduler_running:
            return
        _scheduler_running = False
        logger.info("Scheduler stopped")


def restart_scheduler() -> None:
    """Restart the scheduler, reloading settings and timings."""
    stop_scheduler()
    time.sleep(2)
    schedule.clear()
    start_scheduler()


__all__ = [
    "configure_scheduler",
    "is_scheduler_running",
    "restart_scheduler",
    "start_scheduler",
    "stop_scheduler",
]
