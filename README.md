# Generator Raportów Wpłat z Librusa

## Opis
Aplikacja Flask do przetwarzania raportów wpłat z systemu Librus. Generuje raporty w formacie Excel i PDF z podsumowaniem wpłat według grup zajęciowych i handlowców.

## Funkcjonalności

### 1. Przetwarzanie danych
- Import plików CSV z Librusa
- Filtrowanie według zakresu dat i godzin
- Automatyczne przeliczanie prowizji (10%)
- Przypisywanie handlowców do grup

### 2. Generowane pliki Excel
**Zakładka 1:** Cały plik CSV
**Zakładka 2:** DANE DO WYLICZEŃ - przefiltrowane dane
**Zakładka 3:** PODSUMOWANIE GRUP - suma wpłat, prowizje, handlowcy
**Zakładka 4:** PODSUMOWANIE HANDLOWCÓW - suma prowizji per handlowiec

### 3. Raporty PDF
- Podsumowanie grup zajęciowych
- Podsumowanie handlowców
- Zakres dat widoczny w raporcie

### 4. Zarządzanie handlowcami
- Interfejs web do edycji przypisań
- Dodawanie nowych wpisów
- Usuwanie istniejących

## Instalacja

### Lokalna instalacja

```bash
# Rozpakuj archiwum
unzip librus_raport_app.zip
cd librus_raport_app

# Utwórz środowisko wirtualne
python -m venv venv

# Aktywuj środowisko
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# Zainstaluj zależności
pip install -r requirements.txt

# Uruchom aplikację
python app.py
```

Aplikacja będzie dostępna pod adresem: `http://localhost:5000`

### Instalacja na VPS (Ubuntu/Debian)

```bash
# Aktualizacja systemu
sudo apt update && sudo apt upgrade -y

# Instalacja Python i pip
sudo apt install python3 python3-pip python3-venv -y

# Przesłanie plików na serwer
scp librus_raport_app.zip user@your-vps-ip:/home/user/

# Logowanie na serwer
ssh user@your-vps-ip

# Rozpakowanie
unzip librus_raport_app.zip
cd librus_raport_app

# Utworzenie środowiska wirtualnego
python3 -m venv venv
source venv/bin/activate

# Instalacja zależności
pip install -r requirements.txt

# Instalacja Gunicorn (serwer produkcyjny)
pip install gunicorn

# Test uruchomienia
python app.py
```

### Konfiguracja jako usługa systemd

Utwórz plik `/etc/systemd/system/librus-raport.service`:

```ini
[Unit]
Description=Librus Raport App
After=network.target

[Service]
User=user
WorkingDirectory=/home/user/librus_raport_app
Environment="PATH=/home/user/librus_raport_app/venv/bin"
ExecStart=/home/user/librus_raport_app/venv/bin/gunicorn --workers 3 --bind 0.0.0.0:5000 app:app

[Install]
WantedBy=multi-user.target
```

Uruchomienie usługi:

```bash
sudo systemctl daemon-reload
sudo systemctl start librus-raport
sudo systemctl enable librus-raport
sudo systemctl status librus-raport
```

### Konfiguracja Nginx (opcjonalnie)

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

## Użytkowanie

1. **Uruchom aplikację** - otwórz przeglądarkę i przejdź do adresu aplikacji
2. **Wybierz plik CSV** - eksportowany z Librusa
3. **Ustaw zakres dat** - data od/do + godziny
4. **Generuj raport** - kliknij "Generuj Raport"
5. **Pobierz pliki** - Excel i PDF

### Zarządzanie handlowcami

1. Kliknij "Zarządzaj Handlowcami"
2. Edytuj istniejące wpisy lub dodaj nowe
3. Zapisz zmiany

## Struktura projektu

```
librus_raport_app/
├── app.py                 # Główna aplikacja Flask
├── requirements.txt       # Zależności Python
├── README.md             # Dokumentacja
├── handlowcy.xlsx        # Plik z przypisaniami handlowców
├── templates/            # Szablony HTML
│   ├── index.html       # Strona główna
│   └── handlowcy.html   # Zarządzanie handlowcami
├── uploads/             # Przesłane pliki CSV
└── static/              # Wygenerowane raporty
```

## Wymagania

- Python 3.8+
- Flask 3.0+
- pandas 2.1+
- openpyxl 3.1+
- reportlab 4.0+

## Bezpieczeństwo

⚠️ **Uwaga:** Ta aplikacja nie zawiera zaawansowanych mechanizmów bezpieczeństwa!

Przed wdrożeniem produkcyjnym zaleca się:
- Dodanie uwierzytelniania użytkowników
- Ograniczenie dostępu przez firewall
- Konfiguracja HTTPS (SSL)
- Regularne kopie zapasowe danych

## Wsparcie

W przypadku problemów sprawdź:
1. Logi aplikacji: `sudo journalctl -u librus-raport -f`
2. Uprawnienia do plików i folderów
3. Poprawność formatu pliku CSV z Librusa

## Autor

Aplikacja stworzona dla potrzeb automatyzacji raportowania wpłat.
