FROM python:3.12-slim
WORKDIR /app
COPY app/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app/ .
ENV PORT=5020
EXPOSE 5020
CMD ["python", "app.py"]
