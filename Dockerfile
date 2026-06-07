FROM python:3.12-slim

# Install LibreOffice for DOCX → PDF conversion
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer libreoffice-core \
    fonts-dejavu fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY app.py .
COPY templates/ ./templates/
COPY src/main/resources/static/ ./src/main/resources/static/
COPY jrsu_word_file.docx .

# Create output directories
RUN mkdir -p output uploads

# Expose port (uses PORT env variable, defaults to 8080)
EXPOSE 8080

ENV TEMPLATE_PATH=jrsu_word_file.docx
ENV OUTPUT_DIR=output
ENV UPLOAD_DIR=uploads
ENV LIBREOFFICE=soffice
ENV PORT=8080

# Start the Flask app
CMD ["python", "app.py"]
