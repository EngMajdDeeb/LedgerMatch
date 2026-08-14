# syntax=docker/dockerfile:1
FROM python:3.12-slim

# Prevent .pyc files and force stdout/stderr to be unbuffered (so `docker logs` is live)
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# System deps needed by Pillow (jpeg/zlib) - kept minimal on purpose.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libjpeg62-turbo \
    zlib1g \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
RUN chmod +x docker-entrypoint.sh

# Static files are baked into the image at build time (WhiteNoise serves
# them); media/ and db.sqlite3 are bind-mounted at runtime (see
# docker-compose.yml) so uploads and the database survive a rebuild.
RUN python manage.py collectstatic --noinput

EXPOSE 8000

ENTRYPOINT ["./docker-entrypoint.sh"]

# 3 workers is plenty for an internal tool with ~10 pharmacies. Each request
# can take up to a couple of minutes (sequential OpenAI calls), so the
# worker timeout is raised well above gunicorn's 30s default.
CMD ["gunicorn", "dawak_finance.wsgi:application", \
     "--bind", "0.0.0.0:8000", \
     "--workers", "3", \
     "--timeout", "180"]
