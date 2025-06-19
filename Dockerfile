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
    PATH="/opt/venv/bin:$PATH"

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
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    # Cài đặt Chrome
    && wget -q -O - https://dl.google.com/linux/linux_signing_key.pub | apt-key add - \
    && echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends google-chrome-stable \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Copy virtual environment từ builder stage
COPY --from=builder /opt/venv /opt/venv

# Tạo các thư mục cần thiết
RUN mkdir -p uploads DOWNLOAD logs

# Copy application code
COPY . .

# Tạo non-root user
RUN adduser --disabled-password --gecos '' appuser \
    && chown -R appuser:appuser /app
USER appuser

# Expose port
EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:8000/ || exit 1

# Start application
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--workers", "4", "--threads", "2", "--worker-class", "gthread", "--timeout", "120", "--keep-alive", "5", "--max-requests", "1000", "--max-requests-jitter", "50", "wsgi:application"]
