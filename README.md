# Alstom NDT — Skaner Prób v3.1
## Wdrożenie na Render.com (darmowe)

### Krok 1 — GitHub
1. Wejdź na github.com i załóż darmowe konto
2. Kliknij "New repository" → nazwa: `ndt-app` → Public → Create
3. Wgraj WSZYSTKIE pliki z tego folderu (przeciągnij do przeglądarki)

### Krok 2 — Render
1. Wejdź na render.com → "Get Started for Free"
2. Zaloguj się przez GitHub
3. Kliknij "New +" → "Web Service"
4. Wybierz repozytorium `ndt-app`
5. Ustaw:
   - **Name:** ndt-app
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn server:app`
6. Kliknij "Advanced" → "Add Environment Variable":
   - Key: `ANTHROPIC_API_KEY`
   - Value: [Twój klucz API]
7. Kliknij "Create Web Service"

### Krok 3 — Instalacja na iPhone
1. Otwórz Safari na iPhone
2. Wejdź na link który da Render (np. https://ndt-app.onrender.com)
3. Kliknij przycisk "Udostępnij" (kwadrat ze strzałką)
4. Wybierz "Dodaj do ekranu głównego"
5. Gotowe!

### Co nowego w v3.1
- Jasny motyw (biało-niebieski)
- Obrazki w szablonach Excel są w 100% zachowane (logo Alstom, pieczątki)
- Nie wymaga openpyxl — czystsza instalacja

### Uwaga
Darmowy plan Render "zasypia" po 15 min bezczynności.
Pierwsze uruchomienie może trwać ~30 sekund.
