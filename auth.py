"""Bramka logowania dla aplikacji (Streamlit) — konta użytkowników z hasłami.

Użytkownicy i hasła (HASH bcrypt) są przechowywane w sekretach Streamlit
(`st.secrets["auth"]`), a lokalnie w pliku `.streamlit/secrets.toml`. Dzięki
temu lista kont i hasła NIE trafiają do repozytorium ani do gita.

Format sekcji [auth] — patrz `.streamlit/secrets.toml.example`.
Nowe konto dodajesz skryptem: `python tworz_uzytkownika.py`.
"""
from __future__ import annotations

import streamlit as st


def _to_plain(obj):
    """st.secrets zwraca obiekty typu AttrDict — zamień rekurencyjnie na dict."""
    if hasattr(obj, "items") and not isinstance(obj, (str, bytes)):
        return {k: _to_plain(v) for k, v in obj.items()}
    return obj


def _load_auth_config() -> dict:
    """Wczytuje i waliduje konfigurację logowania z sekretów. Stop przy błędzie."""
    if "auth" not in st.secrets:
        st.error(
            "Logowanie nie jest skonfigurowane. Administrator musi dodać sekcję "
            "[auth] w sekretach aplikacji (Streamlit Cloud → Settings → Secrets) "
            "lub w pliku .streamlit/secrets.toml. "
            "Wzór: .streamlit/secrets.toml.example."
        )
        st.stop()

    auth = _to_plain(st.secrets["auth"])
    credentials = auth.get("credentials") or {}
    if not credentials.get("usernames"):
        st.error(
            "Sekcja [auth] nie zawiera żadnych użytkowników "
            "(auth.credentials.usernames). Dodaj co najmniej jedno konto."
        )
        st.stop()

    cookie_key = auth.get("cookie_key")
    if not cookie_key:
        st.error(
            "Brak `cookie_key` w sekcji [auth]. Wygeneruj losowy klucz "
            '(python -c "import secrets; print(secrets.token_hex(32))") '
            "i dodaj do sekretów."
        )
        st.stop()

    return {
        "credentials": credentials,
        "cookie_name": auth.get("cookie_name", "instrukcje_bhp_auth"),
        "cookie_key": str(cookie_key),
        "cookie_expiry_days": int(auth.get("cookie_expiry_days", 30)),
    }


def require_login():
    """Wyświetla formularz logowania i blokuje aplikację dla niezalogowanych.

    Po udanym logowaniu zwraca (name, username, authenticator).
    W przeciwnym razie zatrzymuje renderowanie strony (st.stop()).
    """
    try:
        import streamlit_authenticator as stauth
    except ModuleNotFoundError:
        st.error(
            "Brak biblioteki `streamlit-authenticator`. "
            "Zainstaluj zależności: pip install -r requirements.txt"
        )
        st.stop()

    config = _load_auth_config()

    authenticator = stauth.Authenticate(
        config["credentials"],
        config["cookie_name"],
        config["cookie_key"],
        config["cookie_expiry_days"],
    )

    # Renderuje formularz logowania (login + hasło) w głównym obszarze.
    authenticator.login(location="main")

    status = st.session_state.get("authentication_status")
    if status is False:
        st.error("Nieprawidłowy login lub hasło.")
        st.stop()
    if status is None:
        st.info("Zaloguj się, aby korzystać z generatora.")
        st.stop()

    # Zalogowany — informacja + przycisk wylogowania w panelu bocznym.
    with st.sidebar:
        st.write(f"Zalogowano jako **{st.session_state.get('name')}**")
        authenticator.logout("Wyloguj", "sidebar")

    return (
        st.session_state.get("name"),
        st.session_state.get("username"),
        authenticator,
    )
