# Generator Instrukcji BHP

Streamlit application that generates Polish BHP (workplace safety) instruction Excel files from PDF safety data sheets.

## Stack
- Python 3.12
- Streamlit (frontend)
- anthropic (ekstrakcja danych z karty charakterystyki przez Claude API)
- pdfplumber, openpyxl, Pillow (przetwarzanie / generowanie Excela)

## Konfiguracja (wymagane)
Ekstrakcja danych z karty charakterystyki działa przez Claude API i obsługuje
także skany (OCR). Wymagany sekret:

- `ANTHROPIC_API_KEY` — klucz API Anthropic. W Replit ustaw w zakładce
  **Secrets**. Bez klucza aplikacja pokaże czytelny błąd i nie wygeneruje
  instrukcji.

Opcjonalnie:

- `CLAUDE_MODEL` — model do ekstrakcji (domyślnie `claude-sonnet-4-6`).

## Logowanie (kontrola dostępu)
Aplikacja jest zabezpieczona logowaniem (biblioteka `streamlit-authenticator`).
Niezalogowani użytkownicy widzą tylko formularz logowania — generator jest
niedostępny.

Konta i hasła (HASH bcrypt) trzymane są w sekcji `[auth]` sekretów aplikacji
(Streamlit Cloud → Settings → Secrets), lokalnie w `.streamlit/secrets.toml`.
Wzór: `.streamlit/secrets.toml.example`. Plik `secrets.toml` jest w `.gitignore`
i nie trafia do repozytorium.

Dodanie / zmiana użytkownika:
1. `python tworz_uzytkownika.py` — podaj login, nazwę, e-mail i hasło.
2. Skopiuj wygenerowany fragment `[auth.credentials.usernames.<login>]`
   do sekretów aplikacji (lub `.streamlit/secrets.toml`).
3. Zapisz — Streamlit przeładuje aplikację z nowym kontem.

Usunięcie użytkownika = usunięcie jego sekcji z sekretów.

Limit dla automatycznej ekstrakcji: PDF do ~30 MB / 100 stron. Większe skany
poza zakresem (do podziału ręcznego). Instalacja zależności:
`pip install -r requirements.txt`.

## Project Layout
- `app.py` — Streamlit UI entry point
- `auth.py` — bramka logowania (konta użytkowników, hasła bcrypt)
- `tworz_uzytkownika.py` — skrypt do generowania danych nowego użytkownika
- `.streamlit/secrets.toml.example` — wzór sekretów (klucz API + konta [auth])
- `extraction_ai.py` — ekstrakcja danych z karty przez Claude API (PDF/skan)
- `main.py` — orkiestracja ekstrakcji i wypełnianie szablonu Excel
- `assets/` — Static assets used in generated workbooks
- `DOCS/` — Sample input PDFs
- `temp_images/` — Temporary image storage during processing
- `wzor.xlsx` — Excel template used as the basis for generated instructions
- `.streamlit/config.toml` — Streamlit server configuration (port 5000, host 0.0.0.0)

## Running Locally
The "Start application" workflow runs:
```
streamlit run app.py --server.port 5000 --server.address 0.0.0.0 --server.headless true --server.enableCORS false --server.enableXsrfProtection false
```

## Deployment
Configured for autoscale deployment running the same Streamlit command.
