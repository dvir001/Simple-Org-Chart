#!/bin/bash
set -e

APP_UID=$(id -u app)
APP_GID=$(id -g app)

# Fix ownership of bind-mounted directories so the 'app' user can write to them.
# Only chown when the top-level directory owner doesn't already match to avoid
# slow recursive traversal of large bind mounts on every container start.
for dir in /app/data /app/config /app/repositories; do
    if [ -d "$dir" ]; then
        current_owner=$(stat -c '%u:%g' "$dir" 2>/dev/null || echo "")
        if [ -z "$current_owner" ] || [ "$current_owner" != "$APP_UID:$APP_GID" ]; then
            chown -R app:app "$dir" 2>/dev/null || true
        fi
    fi
done

# Drop privileges and exec the main process
exec gosu app "$@"
