# Use official Python runtime as base image
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies for PDF processing
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    g++ \
    libjpeg-dev \
    zlib1g-dev \
    libpng-dev \
    libfreetype6-dev \
    mupdf-tools \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements file
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Expose port (Cloud Run uses PORT env variable)
ENV PORT=8080
EXPOSE 8080

# Run the application
CMD exec uvicorn app.main:app --host 0.0.0.0 --port ${PORT}
