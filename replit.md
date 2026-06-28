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

Konta pochodzą z dwóch źródeł (łączonych):
1. Sekrety `[auth]` — konta awaryjne (bootstrap), np. `admin`. Działają zawsze,
   nawet gdy baza jest pusta/niedostępna. Każde konto z sekretów = administrator.
2. Supabase / PostgreSQL (`[supabase]`) — użytkownicy dodawani z panelu
   administratora w aplikacji; zapis trwały, przeżywa restart. Konto może mieć
   flagę administratora (`is_admin`).

### Panel administratora (w aplikacji)
Po zalogowaniu administrator widzi sekcję **„Zarządzanie użytkownikami"**:
dodawanie kont, reset hasła, nadawanie/odbieranie uprawnień admina, usuwanie.
Nowe konta zapisują się w Supabase (PostgreSQL) i są od razu gotowe do logowania.
Wymaga skonfigurowanej sekcji `[supabase]` w sekretach (connection string z
Supabase → Connect → tryb "Session pooler"). Tabela `app_users` tworzy się sama.

### Konto awaryjne przez skrypt
`python tworz_uzytkownika.py` generuje blok `[auth.credentials.usernames.<login>]`
do wklejenia w sekretach — przydatne do utworzenia pierwszego konta admina.

Limit dla automatycznej ekstrakcji: PDF do ~30 MB / 100 stron. Większe skany
poza zakresem (do podziału ręcznego). Instalacja zależności:
`pip install -r requirements.txt`.

## Project Layout
- `app.py` — Streamlit UI entry point
- `auth.py` — bramka logowania (konta z sekretów + Supabase, hasła bcrypt)
- `users_store.py` — trwały magazyn użytkowników w Supabase (PostgreSQL)
- `admin_panel.py` — panel administratora (dodawanie/edycja/usuwanie kont)
- `tworz_uzytkownika.py` — skrypt do generowania konta awaryjnego w sekretach
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
