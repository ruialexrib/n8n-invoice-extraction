FROM python:3.12-slim

RUN pip install --no-cache-dir python-docx==1.1.2

WORKDIR /app
COPY app.py /app/app.py

EXPOSE 8080
CMD ["python", "/app/app.py"]

