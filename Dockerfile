# Use official Playwright image that matches version in requirements.txt (1.40.0)
FROM mcr.microsoft.com/playwright/python:v1.40.0-jammy

WORKDIR /app

# Copy requirements and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Create necessary directories and set permissions
RUN mkdir -p results uploads && chmod 777 results uploads

# Expose port (Render/HF often use 7860 or random)
EXPOSE 7860
ENV PORT=7860

# Run the application
CMD ["python", "app.py"]

