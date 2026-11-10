FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN useradd --create-home --uid 10001 app
USER app
EXPOSE 8770
CMD ["python3", "app.py", "--no-browser", "--host", "0.0.0.0", "--port", "8770"]
