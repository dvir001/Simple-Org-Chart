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
timeout = 30
keepalive = 2

# Restart workers after this many requests, to help prevent memory leaks
max_requests = 1000
max_requests_jitter = 50

# Preload the application before forking worker processes
preload_app = True

# Logging - Option to filter access logs
import logging

class AccessLogFilter(logging.Filter):
    def filter(self, record):
        # Filter out noisy requests to reduce log spam
        if hasattr(record, 'getMessage'):
            message = record.getMessage()
            
            # Filter out photo requests (both 200 and 304 responses)
            if '/api/photo/' in message and ('304' in message or '200' in message):
                return False
                
            # Filter out static file requests with 304 responses (cached)
            if '/static/' in message and '304' in message:
                return False
                
            # Filter out health check requests (Docker health checks with curl)
            if '"GET / HTTP/1.1" 200' in message and 'curl/' in message:
                return False
                
        return True

# Apply custom filter to reduce access log spam
def when_ready(server):
    access_logger = logging.getLogger('gunicorn.access')
    access_logger.addFilter(AccessLogFilter())

accesslog = '-'  # Enable access logging with filtering
errorlog = '-'
loglevel = 'info'
access_log_format = '%(h)s %(l)s %(u)s %(t)s "%(r)s" %(s)s %(b)s "%(f)s" "%(a)s"'

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