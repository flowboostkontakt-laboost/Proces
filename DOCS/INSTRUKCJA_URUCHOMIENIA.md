# Instrukcja uruchomienia generatora instrukcji BHP

## Wymagania

Zainstalowany **Python 3.11+** (sprawdź komendą `python --version`).

## Struktura plików

W głównym folderze projektu (`proces/`) muszą znajdować się:

- `main.py` — silnik przetwarzający PDF-y
- `app.py` — interfejs webowy (Streamlit)
- `requirements.txt` — lista bibliotek Python
- `wzor.xlsx` — szablon Excel
- `assets/` — folder z ikonami (GHS, piktogramy ochrony)
- `DANE_WEJSCIOWE/` — **główny folder** na karty charakterystyki PDF
- `DOCS/` — **folder zapasowy** na PDF-y (używany, gdy `DANE_WEJSCIOWE` jest pusta)
- `GOTOWE_INSTRUKCJE/` — tutaj trafiają wygenerowane pliki Excel

## Krok po kroku — pierwsze uruchomienie

### 1. Zainstaluj zależności

Otwórz terminal w folderze `proces/` i wpisz:

```bash
pip install -r requirements.txt
```

### 2. Przygotuj dane wejściowe

Wrzuć karty charakterystyki w formacie **PDF** do folderu:

```
DANE_WEJSCIOWE/
```

> Jeśli folder `DANE_WEJSCIOWE` jest pusty, program automatycznie szuka PDF-ów w folderze `DOCS/`.

### 3. Uruchom aplikację

W terminalu w folderze `proces/`:

```bash
python -m streamlit run app.py
```

Aplikacja wystartuje i wyświetli adres:

```
Local URL: http://localhost:8501
```

Otwórz ten link w przeglądarce.

### 4. Generowanie instrukcji

1. Zaznacz checkboxami pliki PDF, które chcesz przetworzyć.
2. Kliknij przycisk **Generuj instrukcje**.
3. Poczekaj na pasek postępu.
4. Pobierz gotowe pliki `.xlsx` — zostaną też zapisane w folderze `GOTOWE_INSTRUKCJE/`.

## Uruchamianie kolejny raz

Każde następne uruchomienie to tylko jedna komenda (zakładając, że `DANE_WEJSCIOWE` zawiera nowe PDF-y):

```bash
python -m streamlit run app.py
```

## Co zrobić, gdy coś nie działa

| Problem | Rozwiązanie |
|---------|-------------|
| `pip` nie działa | Upewnij się, że Python jest dodany do zmiennej PATH |
| Brak plików w UI | Sprawdź, czy w `DANE_WEJSCIOWE` lub `DOCS` są pliki `.pdf` |
| Brak szablonu | Upewnij się, że `wzor.xlsx` istnieje w głównym folderze |
| Brak obrazków w Excelu | Sprawdź, czy folder `assets/` zawiera pliki PNG |
| Port 8501 jest zajęty | Dodaj flagę: `python -m streamlit run app.py --server.port 8502` |
