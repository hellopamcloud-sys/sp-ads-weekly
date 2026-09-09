FROM python:3.11-slim
ENV DEBIAN_FRONTEND=noninteractive PYTHONUNBUFFERED=1 TZ=Asia/Shanghai
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-calc-nogui fonts-noto-cjk fonts-liberation ca-certificates && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY *.py ./
ENV PORT=8080 OUT_DIR=/tmp/reports
EXPOSE 8080
CMD ["sh", "-c", "gunicorn -w 1 --threads 4 --timeout 900 -b 0.0.0.0:${PORT} app:app"]
