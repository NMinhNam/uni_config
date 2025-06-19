# Sử dụng multi-stage build để giảm kích thước image
FROM python:3.10-slim AS builder

# Cài đặt các dependencies cần thiết cho việc build
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    gcc \
    g++ \
    libssl-dev \
    libffi-dev \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements trước để tận dụng cache
COPY requirements.txt .

# Tạo virtual environment và cài đặt dependencies
RUN python -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# Stage 2: Runtime image
FROM python:3.10-slim

# Set working directory
WORKDIR /app

# Thiết lập environment variables
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive \
    PYTHONPATH=/app \
    PATH="/opt/venv/bin:$PATH" \
    # Ensure ChromeDriver knows where Chrome is
    CHROME_BIN=/usr/bin/google-chrome \
    # Set a specific stable Chrome version
    CHROME_VERSION="114.0.5735.90"

# Cài đặt các runtime dependencies cần thiết
RUN apt-get update && apt-get install -y --no-install-recommends \
    # Chỉ giữ lại các packages thực sự cần thiết cho runtime
    curl \
    wget \
    gnupg \
    libjpeg-dev \
    libpng-dev \
    tesseract-ocr \
    tesseract-ocr-vie \
    poppler-utils \
    # Add dependencies for Chrome
    ca-certificates \
    fonts-liberation \
    libappindicator3-1 \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libc6 \
    libcairo2 \
    libcups2 \
    libdbus-1-3 \
    libexpat1 \
    libfontconfig1 \
    libgbm1 \
    libgcc1 \
    libglib2.0-0 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libstdc++6 \
    libx11-6 \
    libx11-xcb1 \
    libxcb1 \
    libxcomposite1 \
    libxcursor1 \
    libxdamage1 \
    libxext6 \
    libxfixes3 \
    libxi6 \
    libxrandr2 \
    libxrender1 \
    libxss1 \
    libxtst6 \
    lsb-release \
    xdg-utils \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    # Cài đặt Chrome
    && wget -q -O - https://dl.google.com/linux/linux_signing_key.pub | apt-key add - \
    && echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends google-chrome-stable \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    # Make directories writable for ChromeDriver
    && mkdir -p /home/appuser/.wdm \
    && chmod -R 777 /home/appuser/.wdm

# Copy virtual environment từ builder stage
COPY --from=builder /opt/venv /opt/venv

# Tạo các thư mục cần thiết
RUN mkdir -p uploads DOWNLOAD logs

# Copy application code
COPY . .

# Make ChromeDriver files executable
RUN chmod -R +x /app

# Tạo non-root user
RUN adduser --disabled-password --gecos '' appuser \
    && chown -R appuser:appuser /app \
    && chown -R appuser:appuser /home/appuser

# Create startup script for Railway
RUN echo '#!/bin/bash\n\n# Fix ChromeDriver permissions\nfind ~/.wdm -name "chromedriver*" -type f -exec chmod +x {} \; 2>/dev/null\nfind /tmp -name "chromedriver*" -type f -exec chmod +x {} \; 2>/dev/null\n\n# Fix potential THIRD_PARTY_NOTICES issue\nfind ~/.wdm -name "THIRD_PARTY_NOTICES.chromedriver" -exec chmod -x {} \; 2>/dev/null\n\n# Start the application\nexec gunicorn --bind 0.0.0.0:$PORT --workers 4 --threads 2 --worker-class gthread --timeout 120 --keep-alive 5 --max-requests 1000 --max-requests-jitter 50 wsgi:application' > /app/railway_startup.sh \
    && chmod +x /app/railway_startup.sh

# Switch to non-root user
USER appuser

# Expose port
EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:${PORT:-8000}/ || exit 1

# Use ENTRYPOINT for Railway to ensure script runs
ENTRYPOINT ["/bin/bash", "/app/railway_startup.sh"]
