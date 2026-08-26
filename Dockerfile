FROM python:3.12-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PIP_NO_CACHE_DIR=1

COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

COPY . .

# Refuse to build/deploy an image if imports, core workflows, optimized
# large-file routes or special-price calculations are broken.
RUN python -m py_compile main.py runtime_main.py powerbi_purchase_fields.py performance_optimization.py special_prices.py production_app.py core.py smoke_test.py performance_smoke_test.py special_prices_smoke_test.py && python smoke_test.py && python performance_smoke_test.py && python special_prices_smoke_test.py

EXPOSE 8501

HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8501/health')" || exit 1

CMD ["uvicorn", "production_app:app", "--host", "0.0.0.0", "--port", "8501", "--workers", "1"]
