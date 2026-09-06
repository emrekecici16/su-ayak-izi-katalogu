FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY suayakizi.py su_ayak_izi_sektorleri.json ./
COPY static ./static

ENV SUAYAK_HOST=0.0.0.0
ENV SUAYAK_PORT=8081
EXPOSE 8081

CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:8081", "suayakizi:app"]
