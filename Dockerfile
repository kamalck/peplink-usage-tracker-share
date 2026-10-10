FROM python:3.12-slim
WORKDIR /app
COPY app/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app/ .
ENV PORT=5020
EXPOSE 5020
# One worker + threads: the login rate-limit counter lives in memory and SQLite is single-writer anyway.
CMD ["gunicorn", "--bind", "0.0.0.0:5020", "--workers", "1", "--threads", "4", "--timeout", "300", "--access-logfile", "-", "app:app"]
