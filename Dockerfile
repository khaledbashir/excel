# Stage 1: Build Frontend
FROM node:18-alpine AS frontend-builder
WORKDIR /app/frontend
COPY demo/ExcelAssistant/package.json demo/ExcelAssistant/package-lock.json ./
RUN npm ci
COPY demo/ExcelAssistant ./
RUN npm run build

# Stage 2: Run Backend
FROM python:3.11-slim
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Copy built frontend assets from stage 1
COPY --from=frontend-builder /app/frontend/dist /app/demo/ExcelAssistant/dist

# Set environment variables
ENV PYTHONPATH=/app
ENV PYTHONUNBUFFERED=1

# Expose port
EXPOSE 8000

# Run the server
CMD ["python", "demo/ExcelAssistant/server.py"]
