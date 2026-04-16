FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY dashboard_web.py .
COPY gerar_pptx.py .
COPY razao.xlsx .
COPY orcamento.xlsx .
COPY capa_fundo.jpg .
COPY capa_logo.png .
COPY golfista_preto.png .
COPY logo_fpgolfe_vermelho.png .
COPY logo_fpg_golf_center.png .
COPY logo_final_slide.png .

RUN mkdir -p /app/assets && \
    printf '@media print {\n    body > div > div > div:first-child { display: none !important; }\n    .print-hide { display: none !important; }\n    body { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }\n    @page { size: A4 landscape; margin: 10mm; }\n}\n' > /app/assets/print.css && \
    printf 'document.addEventListener("click", function(e) {\n    if (e.target && e.target.id === "btn-print-oxr") {\n        setTimeout(function() { window.print(); }, 300);\n    }\n});\n' > /app/assets/print.js

EXPOSE 8050

CMD ["gunicorn", "dashboard_web:server", "--bind", "0.0.0.0:8050", "--workers", "2", "--timeout", "120"]
