"""Trwały magazyn użytkowników w MongoDB (Atlas).

Konta dodawane z panelu administratora zapisują się tutaj i przeżywają
restart aplikacji (w przeciwieństwie do plików na Streamlit Cloud).

Konfiguracja w sekretach:

    [mongo]
    uri = "mongodb+srv://user:haslo@cluster.xxxx.mongodb.net/?retryWrites=true&w=majority"
    db  = "instrukcje_bhp"   # opcjonalnie, domyślnie "instrukcje_bhp"

Dokument użytkownika: { _id: <login>, name, email, password(<hash bcrypt>), is_admin }.
"""
from __future__ import annotations

import bcrypt
import streamlit as st

try:
    from pymongo import MongoClient
except ModuleNotFoundError:  # pragma: no cover
    MongoClient = None

DEFAULT_DB = "instrukcje_bhp"
COLLECTION = "users"
MIN_PASSWORD_LEN = 6


def _get_conf() -> tuple[str | None, str]:
    """Zwraca (uri, nazwa_bazy). uri = None gdy brak konfiguracji."""
    if "mongo" in st.secrets:
        m = st.secrets["mongo"]
        return m.get("uri") or None, m.get("db", DEFAULT_DB)
    return st.secrets.get("MONGODB_URI") or None, DEFAULT_DB


def is_configured() -> bool:
    uri, _ = _get_conf()
    return bool(uri) and MongoClient is not None


@st.cache_resource(show_spinner=False)
def _client(uri: str):
    return MongoClient(uri, serverSelectionTimeoutMS=8000, appname="instrukcje-bhp")


def _col():
    uri, db = _get_conf()
    if not uri:
        raise RuntimeError("Brak konfiguracji MongoDB (sekcja [mongo] uri w sekretach).")
    if MongoClient is None:
        raise RuntimeError("Brak biblioteki pymongo. Zainstaluj: pip install -r requirements.txt")
    return _client(uri)[db][COLLECTION]


def hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode(), bcrypt.gensalt()).decode()


def _doc_to_user(d: dict) -> dict:
    return {
        "username": d["_id"],
        "name": d.get("name") or d["_id"],
        "email": d.get("email", ""),
        "is_admin": bool(d.get("is_admin", False)),
        "password": d.get("password", ""),
    }


def list_users() -> list[dict]:
    return [_doc_to_user(d) for d in _col().find({}).sort("_id", 1)]


def get_user(username: str) -> dict | None:
    d = _col().find_one({"_id": username})
    return _doc_to_user(d) if d else None


def add_user(username: str, name: str, email: str, plain: str, is_admin: bool = False) -> None:
    username = (username or "").strip()
    if not username or " " in username:
        raise ValueError("Login nie może być pusty ani zawierać spacji.")
    if not plain or len(plain) < MIN_PASSWORD_LEN:
        raise ValueError(f"Hasło musi mieć co najmniej {MIN_PASSWORD_LEN} znaków.")
    if _col().find_one({"_id": username}):
        raise ValueError(f"Użytkownik '{username}' już istnieje.")
    _col().insert_one(
        {
            "_id": username,
            "name": (name or "").strip() or username,
            "email": (email or "").strip(),
            "password": hash_password(plain),
            "is_admin": bool(is_admin),
        }
    )


def set_password(username: str, plain: str) -> None:
    if not plain or len(plain) < MIN_PASSWORD_LEN:
        raise ValueError(f"Hasło musi mieć co najmniej {MIN_PASSWORD_LEN} znaków.")
    res = _col().update_one({"_id": username}, {"$set": {"password": hash_password(plain)}})
    if res.matched_count == 0:
        raise ValueError(f"Nie ma użytkownika '{username}'.")


def set_admin(username: str, is_admin: bool) -> None:
    _col().update_one({"_id": username}, {"$set": {"is_admin": bool(is_admin)}})


def delete_user(username: str) -> None:
    res = _col().delete_one({"_id": username})
    if res.deleted_count == 0:
        raise ValueError(f"Nie ma użytkownika '{username}'.")
