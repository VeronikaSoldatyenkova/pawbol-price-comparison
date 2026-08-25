FROM python:3.12-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PIP_NO_CACHE_DIR=1

COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

COPY . .

# Refuse to build/deploy an image if imports or core workflows are broken.
RUN python -m py_compile main.py runtime_main.py core.py app.py smoke_test.py && python smoke_test.py

EXPOSE 8501

HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8501/health')" || exit 1

CMD ["uvicorn", "runtime_main:app", "--host", "0.0.0.0", "--port", "8501", "--workers", "1"]
