# Use official Python 3.11 slim base image
FROM python:3.11-slim

# Set environment variables to avoid prompts during package install
ENV DEBIAN_FRONTEND=noninteractive

# Set the working directory inside the container
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    wget \
    xfonts-75dpi \
    xfonts-base \
    fontconfig \
    libxrender1 \
    libxext6 \
    libfreetype6 \
    wkhtmltopdf \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Copy all app files into the container
COPY . .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose the port your Flask app runs on
EXPOSE 7000

# Set default environment variable for Render or similar platforms
ENV PORT=7000

# Command to run your Flask app
CMD ["python", "app.py"]
