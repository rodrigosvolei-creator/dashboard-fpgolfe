FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY dashboard_web.py .
COPY razao.xlsx .
COPY orcamento.xlsx .

RUN mkdir -p /app/assets && \
    printf '@media print {\n    body > div > div > div:first-child { display: none !important; }\n    .print-hide { display: none !important; }\n    body { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }\n    @page { size: A4 landscape; margin: 10mm; }\n}\n' > /app/assets/print.css && \
    printf 'document.addEventListener("click", function(e) {\n    if (e.target && e.target.id === "btn-print-oxr") {\n        setTimeout(function() { window.print(); }, 300);\n    }\n});\n' > /app/assets/print.js

EXPOSE 8050

CMD ["gunicorn", "dashboard_web:server", "--bind", "0.0.0.0:8050", "--workers", "2", "--timeout", "120"]
