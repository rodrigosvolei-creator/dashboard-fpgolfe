FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY dashboard_web.py .
COPY RAZAO_\ 01\ a\ 03\ 2026_\ Oficial.xlsx .
COPY orcamento_2026.xlsx .
COPY assets/ ./assets/

EXPOSE 8050

CMD ["gunicorn", "dashboard_web:server", "--bind", "0.0.0.0:8050", "--workers", "2", "--timeout", "120"]
