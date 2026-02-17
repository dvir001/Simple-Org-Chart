# Use Ubuntu 22.04 with Python 3.10 for better Playwright support
FROM ubuntu:22.04

# Set working directory
WORKDIR /app

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV FLASK_APP=app.py
ENV FLASK_ENV=production
ENV APP_PORT=5000
ENV DEBIAN_FRONTEND=noninteractive
ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

# Install Python 3.10 and system dependencies
RUN apt-get update && apt-get install -y \
    python3.10 \
    python3.10-venv \
    python3-pip \
    gcc \
    curl \
    && rm -rf /var/lib/apt/lists/* \
    && groupadd --system app \
    && useradd --system --gid app --create-home app

# Create symlinks for python/pip
RUN ln -sf /usr/bin/python3.10 /usr/bin/python && \
    ln -sf /usr/bin/python3.10 /usr/bin/python3

# Copy requirements first for better Docker layer caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Install Playwright browsers to system location and set ownership
RUN playwright install --with-deps chromium && \
    chown -R app:app /ms-playwright

# Copy application code
COPY . .

# Create necessary directories for data persistence and adjust ownership
RUN mkdir -p static data data/photos && \
    chown -R app:app /app

# Expose port
EXPOSE ${APP_PORT}

# Health check
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:${APP_PORT:-5000}/ || exit 1

# Switch to non-root user
USER app

# Run the application using gunicorn
CMD ["gunicorn", "--config", "deploy/gunicorn.conf.py", "simple_org_chart:app"]
