import multiprocessing
import os

# Server socket
_default_port = os.getenv("APP_PORT", "5000")
if not (_default_port.isdigit() and 1 <= int(_default_port) <= 65535):
    _default_port = "5000"
bind = f"0.0.0.0:{_default_port}"
backlog = 2048

# Worker processes
workers = multiprocessing.cpu_count() * 2 + 1
worker_class = 'sync'
worker_connections = 1000
timeout = int(os.getenv("GUNICORN_TIMEOUT", "600"))  # Default 10 minutes for PNG screenshot generation
keepalive = 2

# Restart workers after this many requests, to help prevent memory leaks
max_requests = 1000
max_requests_jitter = 50

# Preload the application before forking worker processes
preload_app = True

# Logging - Option to filter access logs
import logging
from datetime import datetime, timezone

class AccessLogFilter(logging.Filter):
    def filter(self, record):
        # Filter out noisy requests to reduce log spam
        try:
            # Access the raw message and args separately to avoid format errors
            if hasattr(record, 'args') and record.args:
                # For gunicorn access logs, args is a dict with request info
                args = record.args
                if isinstance(args, dict):
                    path = args.get('U', '') or ''
                    status = str(args.get('s', ''))
                    remote_addr = args.get('h', '') or ''
                    user_agent = args.get('a', '') or ''
                    
                    # Filter out photo requests
                    if '/api/photo/' in path:
                        return False
                    
                    # Filter out static file requests
                    if '/static/' in path:
                        return False
                    
                    # Filter out settings polling
                    if '/api/settings' in path and status == '200':
                        return False
                    
                    # Filter out auth-check requests
                    if '/api/auth-check' in path and status == '200':
                        return False
                    
                    # Filter out health check requests from localhost
                    if path == '/' and status == '200' and (remote_addr == '127.0.0.1' or 'curl/' in user_agent):
                        return False
        except Exception:
            # If anything goes wrong, allow the log through
            pass
                
        return True

# Apply custom filter to reduce access log spam
def when_ready(server):
    access_logger = logging.getLogger('gunicorn.access')
    access_logger.addFilter(AccessLogFilter())

# Disable access logs - app-level logging provides sufficient visibility
# and avoids format string conflicts with Python's logging module
accesslog = None
errorlog = '-'
loglevel = 'info'

# Custom log date format for error logs
logconfig_dict = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'generic': {
            'format': '[%(asctime)s.000Z] %(levelname)s: %(message)s',
            'datefmt': '%Y-%m-%dT%H:%M:%S',
        },
    },
    'handlers': {
        'console': {
            'class': 'logging.StreamHandler',
            'formatter': 'generic',
            'stream': 'ext://sys.stdout',
        },
    },
    'root': {
        'level': 'INFO',
        'handlers': ['console'],
    },
    'loggers': {
        'gunicorn.error': {
            'level': 'INFO',
            'handlers': ['console'],
            'propagate': False,
        },
    },
}

# Process naming
proc_name = 'db_autorgchart'

# Server mechanics
daemon = False
pidfile = None
user = None
group = None
tmp_upload_dir = None

# SSL (uncomment if you want to use SSL)
# keyfile = 'path/to/keyfile'
# certfile = 'path/to/certfile'