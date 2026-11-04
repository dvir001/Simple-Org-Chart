# Use Ubuntu 24.04 with Python 3.12 for Pillow 13 and Playwright support
FROM ubuntu:24.04

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
ENV PATH="/opt/venv/bin:$PATH"

# Install Python 3.12 and system dependencies
RUN apt-get update && apt-get install -y \
    python3 \
    python3-venv \
    python3-pip \
    gcc \
    curl \
    gosu \
    && rm -rf /var/lib/apt/lists/* \
    && groupadd --system app \
    && useradd --system --gid app --create-home app

# Create an isolated environment for application dependencies
RUN python3 -m venv /opt/venv

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
RUN mkdir -p static data data/photos config repositories && \
    chmod 775 data config repositories && \
    chown -R app:app /app

# Make entrypoint executable
COPY deploy/entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

# Expose port
EXPOSE ${APP_PORT}

# Health check
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:${APP_PORT:-5000}/ || exit 1

# Entrypoint fixes bind-mount permissions then drops to 'app' user
ENTRYPOINT ["/entrypoint.sh"]

# Run the application using gunicorn
CMD ["gunicorn", "--config", "deploy/gunicorn.conf.py", "simple_org_chart:app"]
