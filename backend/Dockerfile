FROM python:3.12-slim

# Set the working directory
WORKDIR /app

# Install OS dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    libpq-dev \
    python3-dev \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy app source
COPY . .

# Expose Cloud Run port
EXPOSE 8080

# Run app using uvicorn (ensure file is named main.py and app instance is named "app")
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080"]
