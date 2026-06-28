"""Bramka logowania + źródło kont użytkowników.

Konta pochodzą z dwóch miejsc i są łączone:
1. Sekrety [auth] — konta "awaryjne" (bootstrap), np. `admin`. Zawsze działają,
   nawet gdy baza jest niedostępna lub pusta — chronią przed zablokowaniem się.
2. MongoDB (users_store) — użytkownicy dodawani z panelu administratora; zapis
   trwały, przeżywa restart aplikacji.

Konta z sekretów mają priorytet — nie da się ich nadpisać kontem z bazy.
Administratorem jest: każde konto z sekretów [auth] ORAZ konto z bazy z flagą
`is_admin`.

Format sekcji [auth] — patrz `.streamlit/secrets.toml.example`.
"""
from __future__ import annotations

import sys

import streamlit as st

import users_store


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


def _merged_credentials(bootstrap: dict) -> dict:
    """Łączy konta z bazy (MongoDB) z kontami awaryjnymi z sekretów.

    Sekrety wygrywają — zapewnia to dostęp nawet gdy baza jest niedostępna.
    """
    usernames: dict = {}

    # 1. Użytkownicy z bazy (jeśli skonfigurowana i dostępna).
    if users_store.is_configured():
        try:
            for u in users_store.list_users():
                usernames[u["username"]] = {
                    "name": u["name"],
                    "email": u["email"],
                    "password": u["password"],
                }
        except Exception as exc:  # baza chwilowo niedostępna — logujemy, nie blokujemy
            print(f"[AUTH] Nie udało się pobrać użytkowników z bazy: {exc}", file=sys.stderr)

    # 2. Konta awaryjne z sekretów — nadpisują ewentualne kolizje z bazy.
    usernames.update(bootstrap.get("usernames", {}))
    return {"usernames": usernames}


def is_admin(username: str | None, bootstrap: dict) -> bool:
    if not username:
        return False
    if username in bootstrap.get("usernames", {}):
        return True  # konta z sekretów = administratorzy
    try:
        u = users_store.get_user(username)
        return bool(u and u["is_admin"])
    except Exception:
        return False


def require_login():
    """Wyświetla formularz logowania i blokuje aplikację dla niezalogowanych.

    Po udanym logowaniu zwraca (name, username, is_admin, authenticator).
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
    bootstrap = config["credentials"]
    credentials = _merged_credentials(bootstrap)

    authenticator = stauth.Authenticate(
        credentials,
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

    username = st.session_state.get("username")
    admin = is_admin(username, bootstrap)

    # Zalogowany — informacja + przycisk wylogowania w panelu bocznym.
    with st.sidebar:
        st.write(f"Zalogowano jako **{st.session_state.get('name')}**")
        if admin:
            st.caption("🔑 Administrator")
        authenticator.logout("Wyloguj", "sidebar")

    return st.session_state.get("name"), username, admin, authenticator
