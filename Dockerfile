# Base image
FROM python:3.10-slim

# Set working directory
WORKDIR /app

# Install system dependencies if cần thiết
RUN apt-get update && apt-get install -y \
    build-essential \
    libssl-dev \
    libffi-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy source code
COPY . .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose port (Railway sẽ gán cổng qua biến PORT)
EXPOSE 8000

# Start Flask using gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "app:app"]
