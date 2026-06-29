"""Ekstrakcja danych z karty charakterystyki (SDS) przez Claude API.

Zastępuje kruchą ekstrakcję regex z `main.py`. PDF (również skanowany /
obrazowy) jest wysyłany do modelu jako dokument base64 — Anthropic wykonuje
OCR stron skanowanych automatycznie. Model zwraca ustrukturyzowany JSON przez
wymuszone użycie narzędzia (tool use, strict), niezależnie od formatu karty.

Wymaga zmiennej środowiskowej ANTHROPIC_API_KEY (sekret w Replit).
Model konfigurowalny przez CLAUDE_MODEL (domyślnie claude-sonnet-4-6).
"""

from __future__ import annotations

import base64
import os
from pathlib import Path

# Limit Anthropic dla dokumentu PDF base64: 32 MB / 100 stron.
MAX_PDF_BYTES = 30 * 1024 * 1024

DEFAULT_MODEL = "claude-sonnet-4-6"

TOOL_NAME = "zapisz_karte"

SYSTEM_PROMPT = (
    "Jesteś ekspertem ds. kart charakterystyki (SDS) i klasyfikacji CLP/GHS. "
    "Otrzymujesz pełną kartę charakterystyki substancji lub mieszaniny w PDF "
    "(może to być skan/obraz — odczytaj go). Twoim zadaniem jest wypełnić "
    "narzędzie `zapisz_karte` danymi przepisanymi WIERNIE z karty.\n\n"
    "Zasady:\n"
    "- Mapuj treści po NUMERZE sekcji (1, 2, 4, 5, 7, 8), nie po nazwie — "
    "różni producenci nazywają sekcje inaczej.\n"
    "- Przepisuj treść w języku polskim dokładnie tak, jak w karcie. Nie "
    "skracaj, nie streszczaj, nie dodawaj informacji, których nie ma w karcie.\n"
    "- Jeśli danego pola nie ma w karcie, zwróć pusty string \"\" lub pustą "
    "listę [] — niczego nie zmyślaj.\n"
    "- Zachowaj wszystkie punkty/podpunkty (np. każdy myślnik osobno na liście)."
)

# Wskazówki per-pole (1:1 do pól ExtractedData w main.py).
_PROPS: dict[str, dict] = {
    "product_name": {
        "type": "string",
        "description": "Sekcja 1.1 — identyfikator produktu / nazwa handlowa.",
    },
    "producer": {
        "type": "string",
        "description": (
            "Sekcja 1.3 — dostawca karty: WYŁĄCZNIE nazwa firmy (producenta/"
            "dostawcy). Pomiń adres, ulicę, kod pocztowy, miasto, kraj, "
            "telefony, faksy, e-maile i numery alarmowe — sama nazwa firmy."
        ),
    },
    "revision_date": {
        "type": "string",
        "description": "Data aktualizacji / numer wersji karty (format jak w karcie).",
    },
    "hazard_statements": {
        "type": "array",
        "description": (
            "Sekcja 2.2 (oznakowanie) — zwroty wskazujące rodzaj zagrożenia. "
            "Każdy zwrot jako osobny obiekt: kod (H/EUH) i pełna treść po polsku."
        ),
        "items": {
            "type": "object",
            "properties": {
                "code": {"type": "string", "description": "np. H280, EUH066"},
                "text": {"type": "string", "description": "Pełna treść zwrotu po polsku"},
            },
            "required": ["code", "text"],
            "additionalProperties": False,
        },
    },
    "hazard_pictograms": {
        "type": "array",
        "description": (
            "Sekcja 2.2 — kody piktogramów GHS obecne w oznakowaniu, "
            "wyłącznie z zakresu GHS01..GHS09 (wielkie litery)."
        ),
        "items": {"type": "string"},
    },
    "hand_protection": {
        "type": "string",
        "description": "Sekcja 8.2 — ochrona rąk.",
    },
    "respiratory_protection": {
        "type": "string",
        "description": "Sekcja 8.2 — ochrona dróg oddechowych.",
    },
    "skin_protection": {
        "type": "string",
        "description": "Sekcja 8.2 — ochrona skóry / ciała.",
    },
    "eye_protection": {
        "type": "string",
        "description": "Sekcja 8.2 — ochrona oczu lub twarzy.",
    },
    "first_aid_inhalation": {
        "type": "string",
        "description": "Sekcja 4.1 — pierwsza pomoc po narażeniu drogą oddechową (wdychanie).",
    },
    "first_aid_skin": {
        "type": "string",
        "description": "Sekcja 4.1 — pierwsza pomoc po kontakcie ze skórą.",
    },
    "first_aid_eyes": {
        "type": "string",
        "description": "Sekcja 4.1 — pierwsza pomoc po kontakcie z oczami.",
    },
    "first_aid_ingestion": {
        "type": "string",
        "description": "Sekcja 4.1 — pierwsza pomoc po połknięciu.",
    },
    "first_aid_general": {
        "type": "string",
        "description": "Sekcja 4.2 — najważniejsze ostre i opóźnione objawy / uwagi ogólne.",
    },
    "fire_suitable": {
        "type": "string",
        "description": "Sekcja 5.1 — odpowiednie środki gaśnicze.",
    },
    "fire_unsuitable": {
        "type": "string",
        "description": "Sekcja 5.1 — niewłaściwe środki gaśnicze.",
    },
    "fire_procedure": {
        "type": "string",
        "description": (
            "Postępowanie w przypadku pożaru — POŁĄCZ sekcję 5.2 (szczególne "
            "zagrożenia powodowane przez substancję/mieszaninę) oraz 5.3 "
            "(informacje/zalecenia dla straży pożarnej, sprzęt ochronny). "
            "Nie wpisuj tu środków gaśniczych."
        ),
    },
    "environmental_release": {
        "type": "string",
        "description": (
            "Postępowanie w przypadku niezamierzonego uwolnienia do środowiska "
            "— POŁĄCZ sekcje 6.1 (indywidualne środki ostrożności, wyposażenie "
            "ochronne, procedury w sytuacji awaryjnej), 6.2 (środki ostrożności "
            "w zakresie ochrony środowiska) oraz 6.3 (metody i materiały "
            "zapobiegające skażeniu / oczyszczanie). Przepisz istotne treści ze "
            "wszystkich trzech podsekcji."
        ),
    },
    "handling": {
        "type": "array",
        "description": (
            "Sekcja 7.1 — środki ostrożności dotyczące bezpiecznego "
            "postępowania. PEŁNA lista wszystkich punktów (po polsku)."
        ),
        "items": {"type": "string"},
    },
    "storage": {
        "type": "array",
        "description": (
            "Sekcja 7.2 (i 7.3, jeśli dotyczy) — warunki bezpiecznego "
            "magazynowania. PEŁNA lista wszystkich punktów (po polsku)."
        ),
        "items": {"type": "string"},
    },
    "nds": {
        "type": "string",
        "description": (
            "Sekcja 8.1 — najwyższe dopuszczalne stężenie / wartości "
            "dopuszczalne narażenia zawodowego (NDS, mg/m3 itp.)."
        ),
    },
}

EXTRACT_TOOL = {
    "name": TOOL_NAME,
    "description": "Zapisz dane przepisane z karty charakterystyki.",
    # Bez `strict: True` — kompilator gramatyki odrzuca schemat tej wielkości
    # ("Schema is too complex for compilation"). `tool_choice` poniżej i tak
    # wymusza wywołanie narzędzia, a model Sonnet 4.6 produkuje strukturę
    # zgodną ze schematem niezawodnie.
    "input_schema": {
        "type": "object",
        "properties": _PROPS,
        "required": list(_PROPS.keys()),
        "additionalProperties": False,
    },
}


def _api_key() -> str:
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        raise RuntimeError(
            "Brak klucza ANTHROPIC_API_KEY. Ustaw sekret ANTHROPIC_API_KEY "
            "(w Replit: Secrets) i uruchom ponownie."
        )
    return key


def _call_model(pdf_bytes: bytes) -> dict:
    try:
        import anthropic
    except ModuleNotFoundError as exc:  # pragma: no cover - zależność z requirements
        raise RuntimeError(
            "Brak pakietu 'anthropic'. Zainstaluj: pip install -r requirements.txt"
        ) from exc

    client = anthropic.Anthropic(api_key=_api_key())
    model = os.environ.get("CLAUDE_MODEL", "").strip() or DEFAULT_MODEL
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("ascii")

    try:
        response = client.messages.create(
            model=model,
            max_tokens=8000,
            # Render order: tools -> system -> messages. Breakpoint na bloku
            # system cache'uje narzędzia + system. PDF (zmienny, duży) jest w
            # messages — celowo NIE cache'owany.
            system=[
                {
                    "type": "text",
                    "text": SYSTEM_PROMPT,
                    "cache_control": {"type": "ephemeral"},
                }
            ],
            tools=[EXTRACT_TOOL],
            tool_choice={"type": "tool", "name": TOOL_NAME},
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": (
                                "Przeanalizuj tę kartę charakterystyki i wypełnij "
                                "narzędzie zapisz_karte."
                            ),
                        },
                        {
                            "type": "document",
                            "source": {
                                "type": "base64",
                                "media_type": "application/pdf",
                                "data": pdf_b64,
                            },
                        },
                    ],
                }
            ],
        )
    except anthropic.AuthenticationError as exc:
        raise RuntimeError(
            "Nieprawidłowy klucz ANTHROPIC_API_KEY (401). Sprawdź sekret."
        ) from exc
    except anthropic.PermissionDeniedError as exc:
        raise RuntimeError(
            "Klucz API nie ma dostępu do modelu (403). Sprawdź uprawnienia / model."
        ) from exc
    except anthropic.RateLimitError as exc:
        raise RuntimeError(
            "Przekroczono limit zapytań do API (429). Spróbuj ponownie za chwilę."
        ) from exc
    except anthropic.APIStatusError as exc:
        raise RuntimeError(
            f"Błąd API Anthropic ({exc.status_code}). Spróbuj ponownie."
        ) from exc
    except anthropic.APIConnectionError as exc:
        raise RuntimeError(
            "Błąd połączenia z API Anthropic. Sprawdź sieć i spróbuj ponownie."
        ) from exc

    for block in response.content:
        if block.type == "tool_use" and block.name == TOOL_NAME:
            return dict(block.input)
    raise RuntimeError(
        "Model nie zwrócił ustrukturyzowanych danych z karty (brak tool_use)."
    )


def extract_data_ai(pdf_path: Path):
    """Wyciągnij dane z karty charakterystyki przez Claude API.

    Zwraca obiekt ExtractedData (z main.py). Rzuca RuntimeError z czytelnym
    komunikatem przy braku klucza, błędzie API lub zbyt dużym pliku.
    """
    from main import ExtractedData, H_TO_GHS  # leniwy import — brak cyklu

    pdf_bytes = pdf_path.read_bytes()
    if len(pdf_bytes) > MAX_PDF_BYTES:
        raise RuntimeError(
            f"Plik {pdf_path.name} jest za duży ({len(pdf_bytes) // (1024 * 1024)} MB). "
            "Limit dla automatycznej ekstrakcji to 30 MB."
        )

    data = _call_model(pdf_bytes)

    def s(key: str) -> str:
        value = data.get(key, "")
        return value.strip() if isinstance(value, str) else ""

    def lst(key: str) -> list[str]:
        value = data.get(key, [])
        if not isinstance(value, list):
            return []
        out: list[str] = []
        for item in value:
            if isinstance(item, str) and item.strip():
                out.append(item.strip())
        return out

    statements = data.get("hazard_statements", [])
    hazard_lines: list[str] = []
    hazard_codes: list[str] = []
    if isinstance(statements, list):
        for st in statements:
            if not isinstance(st, dict):
                continue
            code = str(st.get("code", "")).strip().upper()
            text = str(st.get("text", "")).strip()
            # Wg uwag: zostawiamy sam opis słowny zagrożenia (bez kodu H/EUH).
            # Kody zbieramy osobno — służą do doboru piktogramów GHS.
            line = text or code
            if line:
                hazard_lines.append(line)
            if code:
                hazard_codes.append(code)
    hazard_text = "\n".join(hazard_lines)
    hazard_codes = sorted({c for c in hazard_codes if c})

    pictograms = []
    raw_pict = data.get("hazard_pictograms", [])
    if isinstance(raw_pict, list):
        for code in raw_pict:
            norm = str(code).strip().upper()
            if norm in {f"GHS0{i}" for i in range(1, 10)}:
                pictograms.append(norm)
    pictograms = sorted(set(pictograms))
    # Awaryjnie: jeśli model nie zwrócił piktogramów, wyprowadź je z kodów H.
    if not pictograms and hazard_codes:
        derived: set[str] = set()
        for code in hazard_codes:
            derived |= H_TO_GHS.get(code, set())
        pictograms = sorted(derived)

    # Producent: dla pewności bierzemy samą nazwę firmy (pierwsza niepusta linia),
    # gdyby model mimo instrukcji dołączył adres w kolejnych wierszach.
    producer = s("producer")
    if producer:
        producer = next((ln.strip() for ln in producer.splitlines() if ln.strip()), "")

    return ExtractedData(
        product_name=s("product_name"),
        producer=producer,
        revision_date=s("revision_date"),
        hazard_text=hazard_text,
        hazard_codes=hazard_codes,
        hazard_pictograms=pictograms,
        hand_protection=s("hand_protection"),
        respiratory_protection=s("respiratory_protection"),
        skin_protection=s("skin_protection"),
        eye_protection=s("eye_protection"),
        first_aid_inhalation=s("first_aid_inhalation"),
        first_aid_skin=s("first_aid_skin"),
        first_aid_eyes=s("first_aid_eyes"),
        first_aid_ingestion=s("first_aid_ingestion"),
        first_aid_general=s("first_aid_general"),
        # A48 "Postępowanie w przypadku pożaru" = sekcja 5.2 + 5.3
        fire_overview=s("fire_procedure"),
        fire_suitable=s("fire_suitable"),
        fire_unsuitable=s("fire_unsuitable"),
        environmental_release=s("environmental_release"),
        # before/during/after_work zostają jak w szablonie (nie nadpisujemy).
        before_work=[],
        during_work=[],
        after_work=[],
        storage=lst("storage"),
        handling=lst("handling"),
        nds=s("nds"),
    )
