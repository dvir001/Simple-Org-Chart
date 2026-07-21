"""Authentication helpers and decorators for SimpleOrgChart."""

import re
from functools import wraps
from typing import Any, Callable, Dict

from flask import jsonify, redirect, request, session, url_for

NextHandler = Callable[..., Any]
_next_path_pattern = re.compile(r"^[A-Za-z0-9_\-/]*$")
ROLE_READER = "reader"
ROLE_PRIVILEGED = "privileged"
ROLE_ADMIN = "admin"
_ROLE_LEVELS = {
    ROLE_READER: 1,
    ROLE_PRIVILEGED: 2,
    ROLE_ADMIN: 3,
}


def sanitize_next_path(raw_value: str | None) -> str:
    """Sanitize the "next" query parameter to prevent open redirects."""
    if not raw_value:
        return ""

    candidate = raw_value.strip()

    if candidate.startswith(("http://", "https://", "//")):
        return ""

    candidate = candidate.lstrip("/")

    if not _next_path_pattern.fullmatch(candidate):
        return ""

    return candidate


def is_authenticated() -> bool:
    return bool(session.get("authenticated"))


def current_role() -> str | None:
    if not is_authenticated():
        return None
    return session.get("role", ROLE_ADMIN)


def has_role(required_role: str) -> bool:
    return _ROLE_LEVELS.get(current_role() or "", 0) >= _ROLE_LEVELS[required_role]


def _require_api_role(required_role: str):
    def decorator(func: NextHandler) -> NextHandler:
        @wraps(func)
        def wrapper(*args: Any, **kwargs: Dict[str, Any]):
            if not is_authenticated():
                return jsonify({"error": "Authentication required"}), 401
            if not has_role(required_role):
                return jsonify({"error": "Insufficient permissions"}), 403
            return func(*args, **kwargs)

        return wrapper

    return decorator


def require_auth(func: NextHandler) -> NextHandler:
    """API decorator that requires full administrator access."""

    return _require_api_role(ROLE_ADMIN)(func)


def require_privileged(func: NextHandler) -> NextHandler:
    """API decorator for reports, sync, and privileged exports."""

    return _require_api_role(ROLE_PRIVILEGED)(func)


def login_required(func: NextHandler) -> NextHandler:
    """Route decorator that redirects to the login page when unauthenticated."""

    @wraps(func)
    def wrapper(*args: Any, **kwargs: Dict[str, Any]):
        if not is_authenticated():
            desired_path = sanitize_next_path(request.path)
            params: Dict[str, Any] = {"next": desired_path} if desired_path else {}
            return redirect(url_for("login", **params))
        return func(*args, **kwargs)

    return wrapper


def privileged_login_required(func: NextHandler) -> NextHandler:
    """Page decorator that requires report and sync access."""

    @wraps(func)
    def wrapper(*args: Any, **kwargs: Dict[str, Any]):
        if not is_authenticated():
            desired_path = sanitize_next_path(request.path)
            params: Dict[str, Any] = {"next": desired_path} if desired_path else {}
            return redirect(url_for("login", **params))
        if not has_role(ROLE_PRIVILEGED):
            return "Forbidden", 403
        return func(*args, **kwargs)

    return wrapper


__all__ = [
    "ROLE_ADMIN",
    "ROLE_PRIVILEGED",
    "ROLE_READER",
    "current_role",
    "has_role",
    "is_authenticated",
    "login_required",
    "privileged_login_required",
    "require_auth",
    "require_privileged",
    "sanitize_next_path",
]
