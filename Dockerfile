FROM python:3.11-slim

WORKDIR /app

# System libs required by WeasyPrint (Pango/Cairo) + build tools for pycairo (xhtml2pdf dep)
RUN apt-get update && apt-get install -y --no-install-recommends \
      libpangocairo-1.0-0 \
      libpango-1.0-0 \
      libcairo2 \
      libcairo2-dev \
      libgdk-pixbuf-2.0-0 \
      libffi8 \
      shared-mime-info \
      build-essential \
      python3-dev \
    && rm -rf /var/lib/apt/lists/*

# Install requirements
COPY requirements_web.txt .
RUN pip install --no-cache-dir -r requirements_web.txt

# Copy application files
COPY app.py .
COPY gavel_eps_generator.py .
COPY web_templates/ web_templates/
COPY fonts/ fonts/
COPY templates/ templates/

# Output directory (ephemeral — SVGs are uploaded to Trello)
RUN mkdir -p gavel_eps

EXPOSE 8080

# Single worker + threads so APScheduler runs in one process
CMD ["gunicorn", \
     "--bind", "0.0.0.0:8080", \
     "--workers", "1", \
     "--threads", "8", \
     "--timeout", "0", \
     "--access-logfile", "-", \
     "app:app"]
