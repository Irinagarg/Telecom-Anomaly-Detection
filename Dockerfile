FROM python:3.12-slim

WORKDIR /app

RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

COPY kpi_tracker_app/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY kpi_tracker_app/app.py .
COPY kpi_tracker_app/templates/ templates/

ENV FLASK_ENV=production
ENV PYTHONUNBUFFERED=1

EXPOSE 5000

CMD ["gunicorn", \
     "--bind", "0.0.0.0:5000", \
     "--workers", "1", \
     "--threads", "4", \
     "--timeout", "600", \
     "app:app"]