#!/bin/bash
set -e

# Fix ownership of bind-mounted directories so the 'app' user can write to them
for dir in /app/data /app/config /app/repositories; do
    if [ -d "$dir" ]; then
        chown -R app:app "$dir" 2>/dev/null || true
    fi
done

# Drop privileges and exec the main process
exec gosu app "$@"
