# Użyj lekkiego obrazu Python
FROM python:3.11-slim

# Wyłącz pliki .pyc i bufferowanie (dobre praktyki)
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Ustaw katalog roboczy
WORKDIR /app

# Zainstaluj czcionki DejaVu i narzędzia systemowe
RUN apt-get update && apt-get install -y \
    fonts-dejavu-core \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

# Skopiuj plik z zależnościami
COPY requirements.txt .

# Zainstaluj zależności Python
RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Skopiuj całą aplikację
COPY . .

# Utwórz katalogi dla uploads i output
RUN mkdir -p /app/uploads /app/output /app/static

# Eksponuj port 5000 (wewnętrzny port kontenera)
EXPOSE 5000

# Uruchom aplikację z Gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "3", "--timeout", "120", "--access-logfile", "-", "app:app"]
