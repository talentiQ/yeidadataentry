FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV FLASK_SECRET_KEY=change-this-secret-key
EXPOSE 5000
CMD ["python", "app.py"]
